// Parameter contract sweep: what each tool SENDS against what the bridge READS.
//
// MCP does not validate parameter names, and PowerShell ignores object properties nobody reads, so
// a mismatch is silent: delete_sheet(name) deleted the active sheet, close_workbook(save) ignored
// save=false, find_replace sent keys the bridge never looked at.
//
// Two defect classes are checked, both static (no COM call is made):
//   A. the handler sends a parameter the bridge never reads  -> silently ignored at runtime
//   B. the schema advertises a parameter the handler never sends -> silently ignored by the tool
//
// The bridge side comes from its own generated key table (the __validateParams control frame),
// so it always matches the shipped script. The tool side is parsed out of the handler source.
//
// Failure behaviour (FIXES 117): a host that reports `ready:false` (the single-instance lease is
// held by another DSH session) or exits mid-sweep now aborts with that message and exit code 2.
// Before, the `ready:false` frame was mistaken for a successful start and the first invoke waited
// forever on a child that had already exited - `gates:code` hung with no output at all.
//
// Usage: node scripts/param-contract.mjs [--json] [--check]
import { spawn } from "node:child_process";
import { analyseToolSource } from './lib/tool-action-map.mjs';
import { readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

const PS = path.join(process.env.SystemRoot || "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
const jsonOut = process.argv.includes("--json");
// --check：只比对已入库的 docs/param-contract.md，不写盘（供只读门禁 gates:verify 使用）。
const checkOnly = process.argv.includes("--check");

// The MCP server and the host are local processes; anything slower than this is a stuck child, not
// a slow action. The host budget is generous because a cold start pays dot-source + JIT + orphan
// reclaim once (measured ~0.4s, but a busy machine or a slow COM reclaim can exceed it).
const SERVER_REQUEST_TIMEOUT_MS = 10000;
const HOST_READY_TIMEOUT_MS = 30000;
const HOST_INVOKE_TIMEOUT_MS = 30000;

const { map: TOOLS, unparsed } = analyseToolSource();

let server = null;
let host = null;

/** Fail loudly instead of hanging: kill both children, say why, exit non-zero. */
function abort(error) {
  console.error("param-contract: " + (error && error.message ? error.message : String(error)));
  console.error("param-contract: sweep aborted; docs/param-contract.md was not written.");
  try { if (host && host.child) host.child.kill(); } catch { }
  try { if (server && server.child) server.child.kill(); } catch { }
  process.exit(2);
}

function spawnServer() {
  const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true, env: { ...process.env, WPS_OFFICE_TOOLSET: "full" } });
  let buf = "";
  const pending = new Map();
  const settle = (id, kind, value) => {
    const entry = pending.get(id);
    if (!entry) return;
    pending.delete(id);
    entry[kind](value);
  };
  const failAll = (error) => { for (const id of [...pending.keys()]) settle(id, "reject", error); };
  child.stdout.on("data", (d) => {
    buf += d.toString();
    let i;
    while ((i = buf.indexOf("\n")) >= 0) {
      const line = buf.slice(0, i).trim();
      buf = buf.slice(i + 1);
      if (!line) continue;
      let m; try { m = JSON.parse(line); } catch { continue; }
      if (m.id && pending.has(m.id)) settle(m.id, "resolve", m);
    }
  });
  child.stderr.on("data", () => {});
  child.on("exit", (code, signal) => failAll(new Error("MCP server exited before answering (code=" + code + ", signal=" + signal + ")")));
  child.on("error", (error) => failAll(new Error("MCP server spawn failed: " + error.message)));
  let id = 0;
  const req = (method, params) => new Promise((resolve, reject) => {
    const n = ++id;
    const timer = setTimeout(() => { pending.delete(n); reject(new Error("timeout after " + SERVER_REQUEST_TIMEOUT_MS + "ms waiting for " + method)); }, SERVER_REQUEST_TIMEOUT_MS);
    pending.set(n, { resolve: (m) => { clearTimeout(timer); resolve(m); }, reject: (e) => { clearTimeout(timer); reject(e); } });
    child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: n, method, params }) + "\n");
  });
  return { child, req };
}

function spawnHost() {
  const child = spawn(PS, ["-NoProfile", "-NoLogo", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-File", "host/wps-com-host.ps1"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
  let buf = "";
  const pending = new Map();
  let ready = null;
  let startupError = null;
  const settle = (id, kind, value) => {
    const entry = pending.get(id);
    if (!entry) return;
    pending.delete(id);
    entry[kind](value);
  };
  const failAll = (error) => {
    if (!startupError) startupError = error;
    for (const id of [...pending.keys()]) settle(id, "reject", startupError);
  };
  child.stdout.on("data", (d) => {
    buf += d.toString();
    let i;
    while ((i = buf.indexOf("\n")) >= 0) {
      const line = buf.slice(0, i).trim();
      buf = buf.slice(i + 1);
      if (!line) continue;
      let m; try { m = JSON.parse(line); } catch { continue; }
      // Only `ready: true` means the host loop is up. A `ready: false` frame is the lease refusal
      // and carries the human-readable reason; treating it as readiness is what used to hang.
      if (m.id === undefined || m.id === null) {
        if (m.ready === true) { ready = m; continue; }
        if (m.ready === false) { failAll(new Error("COM host refused to start: " + String(m.error || "unknown reason"))); continue; }
      }
      if (m.id && pending.has(m.id)) settle(m.id, "resolve", m);
    }
  });
  child.stderr.on("data", () => {});
  // A host that dies mid-sweep must reject every queued action; otherwise each pending invoke waits
  // on a child that will never answer (there is no reply for a dead process).
  child.on("exit", (code, signal) => failAll(new Error("COM host exited (code=" + code + ", signal=" + signal + ")")));
  child.on("error", (error) => failAll(new Error("COM host spawn failed: " + error.message)));
  let id = 0;
  const invoke = (action, params) => new Promise((resolve, reject) => {
    if (startupError) { reject(startupError); return; }
    const n = ++id;
    const timer = setTimeout(() => { pending.delete(n); reject(new Error("timeout after " + HOST_INVOKE_TIMEOUT_MS + "ms waiting for action " + action)); }, HOST_INVOKE_TIMEOUT_MS);
    pending.set(n, { resolve: (m) => { clearTimeout(timer); resolve(m); }, reject: (e) => { clearTimeout(timer); reject(e); } });
    child.stdin.write(JSON.stringify({ id: n, action, params }) + "\n");
  });
  const waitReady = () => new Promise((resolve, reject) => {
    const t0 = Date.now();
    const tick = () => {
      if (ready) return resolve(ready);
      if (startupError) return reject(startupError);
      if (Date.now() - t0 > HOST_READY_TIMEOUT_MS) return reject(new Error("COM host never reported ready within " + HOST_READY_TIMEOUT_MS + "ms"));
      setTimeout(tick, 50);
    };
    tick();
  });
  return { child, invoke, waitReady };
}

server = spawnServer();
await server.req("initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "param-contract", version: "1" } }).catch(abort);
server.child.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) + "\n");
const list = await server.req("tools/list", {}).catch(abort);
const tools = new Map(list.result.tools.map((t) => [t.name, t]));

host = spawnHost();
await host.waitReady().catch(abort);

// Cache the bridge answer per action: one hop per action, not per tool.
const bridgeCache = new Map();
async function bridgeInfo(action) {
  if (bridgeCache.has(action)) return bridgeCache.get(action);
  let res;
  try { res = await host.invoke("__validateParams", { action, keys: [] }); } catch (error) { abort(error); }
  const data = res && res.result && res.result.data;
  const answer = data && data.validated
    ? { accepted: new Set(data.accepted), containers: data.containers || [] }
    : null;
  bridgeCache.set(action, answer);
  return answer;
}

const classA = []; // handler sends what the bridge never reads
const classB = []; // schema advertises what the handler never sends
const classC = []; // a nested object carries properties the action never reads
const classD = []; // a pass-through handler advertises a parameter the bridge never reads
const unvalidated = [];
let checked = 0;

for (const [toolName, info] of TOOLS) {
  const tool = tools.get(toolName);
  if (!tool) continue;
  const bridge = await bridgeInfo(info.action);
  if (!bridge) { unvalidated.push({ tool: toolName, action: info.action }); continue; }
  const accepted = bridge.accepted;
  checked++;
  const ignored = info.keys.filter((k) => !accepted.has(k));
  if (ignored.length) classA.push({ tool: toolName, action: info.action, ignored, accepted: [...accepted] });
  const schema = Object.keys((tool.inputSchema && tool.inputSchema.properties) || {});
  const dropped = schema.filter((k) => !info.mentioned(k));
  if (dropped.length) classB.push({ tool: toolName, action: info.action, dropped, sent: info.keys });


  // Nested objects are merged onto the flat key set before the check, so their property names are
  // part of the contract too - and the tool schemas are the only place they are declared.
  const props = (tool.inputSchema && tool.inputSchema.properties) || {};
  for (const container of bridge.containers) {
    const declaration = props[container];
    if (!declaration || declaration.type !== "object" || !declaration.properties) continue;
    const subKeys = Object.keys(declaration.properties);
    const unknown = subKeys.filter((k) => !accepted.has(k));
    if (unknown.length) classC.push({ tool: toolName, action: info.action, container, unknown, accepted: [...accepted] });
  }
}

// A handler whose arguments cannot be read statically is normally a pass-through: it forwards the
// caller's object unchanged, so its schema describes what reaches the bridge. Checking schema ->
// bridge is the only static coverage those tools can get - and it is how evaluate_formula's 'cell'
// was found advertised but never read.
for (const entry of unparsed) {
  if (!entry.action) continue;
  const tool = tools.get(entry.tool);
  if (!tool) continue;
  const bridge = await bridgeInfo(entry.action);
  if (!bridge) continue;
  const schema = Object.keys((tool.inputSchema && tool.inputSchema.properties) || {});
  const unknown = schema.filter((k) => !bridge.accepted.has(k));
  if (unknown.length) classD.push({ tool: entry.tool, action: entry.action, unknown, accepted: [...bridge.accepted] });
}

await host.invoke("__shutdown", {}).catch(() => { });
host.child.kill();
server.child.kill();

const report = { tools: tools.size, checked, classA, classB, classC, classD, unvalidated, unparsed };

// The doc is generated so the remaining work is always the real, current list.
const doc = [
  "# Parameter contract report",
  "",
  "Generated by `node scripts/param-contract.mjs` - do not edit by hand.",
  "",
  "Every tool's own handler source is parsed for the parameter object it hands to the bridge, and",
  "that key set is compared with the keys the action actually reads (the bridge derives those from",
  "its own switch in `scripts/build-host-actions.ps1`). No COM call is made.",
  "",
  "| metric | count |",
  "| --- | --- |",
  "| tools in the full catalog | " + tools.size + " |",
  "| tool/action pairs checked | " + checked + " |",
  "| **A. handler sends a parameter the bridge never reads** | **" + classA.length + "** |",
  "| B. schema advertises a parameter the handler never uses | " + classB.length + " |",
  "| **C. nested object carries a property the action never reads** | **" + classC.length + "** |",
  "| **D. pass-through handler advertises a parameter the bridge never reads** | **" + classD.length + "** |",
  "| actions with no key table (guard skipped) | " + unvalidated.length + " |",
  "| handlers whose arguments are not statically readable | " + unparsed.length + " |",
  "| of those, still covered by the D check below | " + unparsed.filter((u) => u.action).length + " |",
  "",
  "## A. Sent by the tool, never read by the bridge",
  "",
  "These parameters are dropped today: the tool appears to accept them and the action quietly",
  "proceeds without them. The bridge now rejects them with an explicit error instead, so each row",
  "is either an alias to reconcile or a capability to implement.",
  "",
  "| tool | action | parameter(s) dropped | bridge reads |",
  "| --- | --- | --- | --- |",
  ...classA.map((m) => "| `" + m.tool + "` | `" + m.action + "` | " + m.ignored.map((k) => "`" + k + "`").join(", ") + " | " + (m.accepted.length ? m.accepted.map((k) => "`" + k + "`").join(", ") : "(none)") + " |"),
  "",
  "## B. Advertised by the schema, never used by the handler",
  "",
  classB.length === 0 ? "None." : "| tool | action | parameter(s) unused |",
  ...(classB.length === 0 ? [] : ["| --- | --- | --- |", ...classB.map((m) => "| `" + m.tool + "` | `" + m.action + "` | " + m.dropped.map((k) => "`" + k + "`").join(", ") + " |")]),
  "",
  "## C. Nested object properties the action never reads",
  "",
  "The bridge merges a nested container onto the flat key set before checking, so these property",
  "names are part of the contract too - and the tool schema is the only place they are declared.",
  "",
  classC.length === 0 ? "None." : "| tool | action | container | unused property |",
  ...(classC.length === 0 ? [] : ["| --- | --- | --- | --- |", ...classC.map((m) => "| `" + m.tool + "` | `" + m.action + "` | `" + m.container + "` | " + m.unknown.map((k) => "`" + k + "`").join(", ") + " |")]),
  "",
  "## Not checked",
  "",
  "A handler listed here is not necessarily unchecked: when it forwards the caller's object unchanged",
  "(the common pass-through shape), the D check above compares its schema with the bridge directly.",
  "Only an entry with no action name has no coverage at all.",
  "",
  "| tool | reason |",
  "| --- | --- |",
  ...unvalidated.map((u) => "| `" + u.tool + "` | bridge has no key table for `" + u.action + "` |"),
  ...unparsed.map((u) => "| `" + u.tool + "` | " + u.reason + " |"),
  "",
].join("\n");
let docStale = false;
if (checkOnly) {
  let onDisk = null;
  try { onDisk = readFileSync("docs/param-contract.md", "utf8"); } catch { }
  docStale = onDisk !== doc;
  console.log(docStale
    ? "param-contract: docs/param-contract.md is STALE (run `node scripts/param-contract.mjs` to regenerate)"
    : "param-contract: docs/param-contract.md is up to date");
} else {
  writeFileSync("docs/param-contract.md", doc);
}

if (jsonOut) console.log(JSON.stringify(report, null, 2));
else {
  console.log("tools=" + tools.size + " checked=" + checked);
  console.log("");
  console.log("A. handler SENDS a parameter the bridge NEVER READS  (silent no-op): " + classA.length);
  for (const m of classA) console.log("  " + m.tool + " -> " + m.action + "\n      sent-not-read: " + m.ignored.join(", ") + "\n      bridge reads: " + m.accepted.join(", "));
  console.log("");
  console.log("B. schema ADVERTISES a parameter the handler NEVER USES  (silent no-op): " + classB.length);
  for (const m of classB) console.log("  " + m.tool + " -> " + m.action + "\n      advertised-but-unused: " + m.dropped.join(", ") + "\n      object sent to the bridge: " + m.sent.join(", "));
  console.log("");
  console.log("C. nested object carries a property the action NEVER READS: " + classC.length);
  console.log("D. pass-through handler advertises what the bridge NEVER READS: " + classD.length);
  for (const m of classD) console.log("  " + m.tool + " -> " + m.action + "\n      never read: " + m.unknown.join(", ") + "\n      bridge reads: " + m.accepted.join(", "));
  for (const m of classC) console.log("  " + m.tool + " -> " + m.action + "  [" + m.container + "]\n      unused: " + m.unknown.join(", "));
  console.log("");
  console.log("UNVALIDATED (bridge has no key table for the action): " + unvalidated.length);
  for (const u of unvalidated) console.log("  " + u.tool + " -> " + u.action);
  console.log("UNPARSED (handler arguments not statically readable): " + unparsed.length);
  for (const u of unparsed) console.log("  " + u.tool + " (" + u.reason + ")");
}
process.exit(classA.length + classB.length + classC.length + classD.length === 0 && !docStale ? 0 : 1);

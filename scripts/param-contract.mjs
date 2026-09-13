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
// Usage: node scripts/param-contract.mjs [--json]
import { spawn } from "node:child_process";
import { readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

const PS = path.join(process.env.SystemRoot || "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
const jsonOut = process.argv.includes("--json");

/**
 * Keys of the object literal handed to executeMethod, or null when it cannot be read statically.
 * `from` must be the offset just past the action string, so that a result-type generic such as
 * executeMethod<{ success: boolean }>(...) is never mistaken for the argument object.
 */
function sentKeys(source, from) {
  const after = source.slice(from);
  // The argument object may be preceded by a comment, and may be a variable holding it.
  const m = /^\s*,\s*(?:\/\/[^\n]*\n\s*)*\{/.exec(after);
  if (!m) {
    const idm = /^\s*,\s*([A-Za-z_$][\w$]*)\s*(?:,|\))/.exec(after);
    if (!idm) return null;
    const decl = new RegExp("(?:const|let|var)\\s+" + idm[1] + "\\s*(?::[^=]*)?=\\s*\\{").exec(source);
    if (!decl) return null;
    return collectKeys(source, decl.index + decl[0].length - 1);
  }
  return collectKeys(after, m[0].length - 1);
}

function collectKeys(text, braceAt) {
  const after = text;
  let depth = 0;
  let end = -1;
  for (let i = braceAt; i < after.length; i++) {
    const ch = after[i];
    if (ch === "{") depth++;
    else if (ch === "}") { depth--; if (depth === 0) { end = i; break; } }
  }
  if (end < 0) return null;
  const body = after.slice(braceAt + 1, end);
  const keys = new Set();
  let d = 0;
  let segment = "";
  const segments = [];
  for (const ch of body) {
    if ("{[(".includes(ch)) d++;
    else if ("}])".includes(ch)) d--;
    if (ch === "," && d === 0) { segments.push(segment); segment = ""; continue; }
    segment += ch;
  }
  segments.push(segment);
  for (const raw of segments) {
    const s = raw.replace(/\/\/[^\n]*/g, "").trim();
    if (!s) continue;
    if (s.startsWith("...")) return null; // a spread hides the real key set
    const m = /^([A-Za-z_$][\w$]*)\s*(?::([\s\S]*))?$/.exec(s);
    if (!m) return null;
    keys.add(m[1]);
  }
  return keys;
}

function analyseToolSource() {
  const map = new Map();
  const unparsed = [];
  const walk = (dir) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, entry.name);
      if (entry.isDirectory()) { walk(p); continue; }
      if (!entry.name.endsWith(".ts")) continue;
      const src = readFileSync(p, "utf8");
      const names = [...src.matchAll(/name:\s*'(wps_[a-z0-9_]+)'/g)];
      for (let i = 0; i < names.length; i++) {
        const from = names[i].index;
        const to = i + 1 < names.length ? names[i + 1].index : src.length;
        const block = src.slice(from, to);
        // Only a call with an explicitly named action can be traced without executing it.
        const calls = [...block.matchAll(/executeMethod[\s\S]{0,600}?'([A-Za-z][A-Za-z0-9]*)'/g)];
        if (calls.length !== 1) { unparsed.push({ tool: names[i][1], reason: calls.length + " executeMethod calls" }); continue; }
        const keys = sentKeys(block, calls[0].index + calls[0][0].length);
        if (!keys) { unparsed.push({ tool: names[i][1], reason: "argument object not statically readable" }); continue; }
        // A schema key is dropped only if the handler never mentions it at all; renaming it (for
        // example filePath -> path) is a legitimate adaptation, not a defect.
        const mentioned = (key) => new RegExp("\\b" + key + "\\b").test(block);
        map.set(names[i][1], { action: calls[0][1], keys: [...keys], mentioned });
      }
    }
  };
  walk("mcp/src/tools");
  return { map, unparsed };
}
const { map: TOOLS, unparsed } = analyseToolSource();

function spawnServer() {
  const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true, env: { ...process.env, WPS_OFFICE_TOOLSET: "full" } });
  let buf = "";
  const pending = new Map();
  child.stdout.on("data", (d) => {
    buf += d.toString();
    let i;
    while ((i = buf.indexOf("\n")) >= 0) {
      const line = buf.slice(0, i).trim();
      buf = buf.slice(i + 1);
      if (!line) continue;
      let m; try { m = JSON.parse(line); } catch { continue; }
      if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); }
    }
  });
  child.stderr.on("data", () => {});
  let id = 0;
  const req = (method, params) => new Promise((r) => { const n = ++id; pending.set(n, r); child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: n, method, params }) + "\n"); });
  return { child, req };
}

function spawnHost() {
  const child = spawn(PS, ["-NoProfile", "-NoLogo", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-File", "host/wps-com-host.ps1"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
  let buf = "";
  const pending = new Map();
  let ready = null;
  child.stdout.on("data", (d) => {
    buf += d.toString();
    let i;
    while ((i = buf.indexOf("\n")) >= 0) {
      const line = buf.slice(0, i).trim();
      buf = buf.slice(i + 1);
      if (!line) continue;
      let m; try { m = JSON.parse(line); } catch { continue; }
      if (!m.id && m.ready !== undefined) { ready = m; continue; }
      if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); }
    }
  });
  child.stderr.on("data", () => {});
  let id = 0;
  const invoke = (action, params) => new Promise((r) => { const n = ++id; pending.set(n, r); child.stdin.write(JSON.stringify({ id: n, action, params }) + "\n"); });
  const waitReady = () => new Promise((resolve, reject) => {
    const t0 = Date.now();
    const tick = () => { if (ready) return resolve(ready); if (Date.now() - t0 > 30000) return reject(new Error("host never reported ready")); setTimeout(tick, 50); };
    tick();
  });
  return { child, invoke, waitReady };
}

const server = spawnServer();
await server.req("initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "param-contract", version: "1" } });
server.child.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) + "\n");
const list = await server.req("tools/list", {});
const tools = new Map(list.result.tools.map((t) => [t.name, t]));

const host = spawnHost();
await host.waitReady();

// Cache the bridge answer per action: one hop per action, not per tool.
const bridgeCache = new Map();
async function bridgeInfo(action) {
  if (bridgeCache.has(action)) return bridgeCache.get(action);
  const res = await host.invoke("__validateParams", { action, keys: [] });
  const data = res && res.result && res.result.data;
  const answer = data && data.validated
    ? { accepted: new Set(data.accepted), containers: data.containers || [] }
    : null;
  bridgeCache.set(action, answer);
  return answer;
}
async function bridgeAccepts(action) {
  const info = await bridgeInfo(action);
  return info ? info.accepted : null;
}

const classA = []; // handler sends what the bridge never reads
const classB = []; // schema advertises what the handler never sends
const classC = []; // a nested object carries properties the action never reads
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

await host.invoke("__shutdown", {});
host.child.kill();
server.child.kill();

const report = { tools: tools.size, checked, classA, classB, classC, unvalidated, unparsed };

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
  "| actions with no key table (guard skipped) | " + unvalidated.length + " |",
  "| handlers whose arguments are not statically readable | " + unparsed.length + " |",
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
  "| tool | reason |",
  "| --- | --- |",
  ...unvalidated.map((u) => "| `" + u.tool + "` | bridge has no key table for `" + u.action + "` |"),
  ...unparsed.map((u) => "| `" + u.tool + "` | " + u.reason + " |"),
  "",
].join("\n");
writeFileSync("docs/param-contract.md", doc);

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
  for (const m of classC) console.log("  " + m.tool + " -> " + m.action + "  [" + m.container + "]\n      unused: " + m.unknown.join(", "));
  console.log("");
  console.log("UNVALIDATED (bridge has no key table for the action): " + unvalidated.length);
  for (const u of unvalidated) console.log("  " + u.tool + " -> " + u.action);
  console.log("UNPARSED (handler arguments not statically readable): " + unparsed.length);
  for (const u of unparsed) console.log("  " + u.tool + " (" + u.reason + ")");
}
process.exit(classA.length + classB.length + classC.length === 0 ? 0 : 1);

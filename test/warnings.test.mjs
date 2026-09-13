// Best-effort failures must not vanish: every swallowed error in the bridge is collected and
// attached to the action's result as "warnings" (see Add-WpsWarning / Output-Json).
// Run: node test/warnings.test.mjs
import { spawn } from "node:child_process";

const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function payload(res) { try { return JSON.parse(res.result.content[0].text); } catch { return {}; } }
function ok(res) { return !!(res && res.result && !res.result.isError); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "warn", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {} } });
const raw = async (method, params) => payload(await viaAction(method, params));

const created = await viaAction("createDocument", {});
check("scratch document created", ok(created), "");

// A style name that does not exist makes the bridge's optional style assignment fail. The text is
// still inserted, so the action succeeds - and the failure must be reported, not swallowed.
const seeded = await raw("insertText", { text: "warning probe", position: "start", style: "NoSuchStyleName" });
check("action with a best-effort failure still succeeds", seeded.success === true, JSON.stringify(seeded).slice(0, 90));
check("the failure is reported as a warning", Array.isArray(seeded.warnings) && seeded.warnings.length > 0, JSON.stringify(seeded.warnings || []).slice(0, 110));
check("the warning carries the actual error text", typeof seeded.warnings?.[0] === "string" && seeded.warnings[0].length > 8, String(seeded.warnings?.[0]).slice(0, 80));
check("warnings are also inside data for pass-through tools", Array.isArray(seeded.data?.warnings), JSON.stringify(seeded.data || {}).slice(0, 90));

// Warnings are per action: a clean call must not inherit the previous one.
const clean = await raw("ping", {});
check("a clean action reports no warnings", clean.warnings === undefined, JSON.stringify(clean).slice(0, 80));

// And a first-class tool whose optional step succeeds reports nothing either.
const text = await raw("getDocumentText", {});
check("a successful action after a warned one stays clean", text.warnings === undefined, JSON.stringify(text).slice(0, 90));

for (let i = 0; i < 3; i++) {
  const res = await call("wps_call", { tool: "wps_execute_method", args: { method: "closeDocument", params: { save: false } } });
  if (!ok(res)) break;
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "WARNING TESTS OK (" + results.length + ")" : "WARNING TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

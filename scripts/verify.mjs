import { spawn } from "node:child_process";

const entry = process.argv[2];
if (!entry) { console.error("usage: node scripts/verify.mjs <mcp-entry.js>"); process.exit(2); }

const child = spawn(process.execPath, [entry], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
let stderr = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((resolve) => { pending.set(id, resolve); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", (d) => { stderr += d.toString(); });
child.on("error", (e) => { console.log("FAIL spawn " + e.message); process.exit(1); });

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function textOf(res) { return res && res.result && res.result.content && res.result.content[0] ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }

const timeout = setTimeout(() => { console.log("FAIL timeout waiting for server"); child.kill(); process.exit(1); }, 60000);

const init = await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "verify", version: "1.0.0" } });
const si = init.result && init.result.serverInfo;
check("initialize", !!si, si ? si.name + "@" + si.version : "no serverInfo");
send({ jsonrpc: "2.0", method: "notifications/initialized" });

const list = await req(2, "tools/list", {});
const tools = (list.result && list.result.tools) || [];
let bytes = 0;
for (const t of tools) bytes += Buffer.byteLength(JSON.stringify(t), "utf8");
check("tools/list", tools.length > 0, "tools=" + tools.length + " schemaBytes=" + bytes + " approxTokens=" + Math.round(bytes / 3.5));

const names = new Set(tools.map((t) => t.name));
for (const required of ["wps_common_ping", "wps_common_wire_check", "wps_execute_method"]) {
  check("catalog contains " + required, names.has(required));
}

const ping = await req(3, "tools/call", { name: "wps_common_ping", arguments: {} });
check("call wps_common_ping", !!(ping.result && !ping.result.isError), textOf(ping).replace(/\s+/g, " ").slice(0, 100));

const wire = await req(4, "tools/call", { name: "wps_common_wire_check", arguments: {} });
check("call wps_common_wire_check", !!(wire.result && !wire.result.isError), textOf(wire).replace(/\s+/g, " ").slice(0, 100));

clearTimeout(timeout);
if (stderr.trim()) console.log("server stderr tail: " + stderr.trim().split("\n").slice(-3).join(" | ").slice(0, 300));
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "VERIFY OK (" + results.length + " checks)" : "VERIFY FAILED (" + failed + "/" + results.length + ")");
child.kill();
process.exit(failed === 0 ? 0 : 1);

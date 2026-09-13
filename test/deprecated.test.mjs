// Verifies the merged/duplicate tool collapse: deprecated names stay callable but
// disappear from wps_help discovery and forward to the canonical tool.
// Run: node test/deprecated.test.mjs
import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import path from "node:path";

const require = createRequire(import.meta.url);
const { DEPRECATED_TOOLS } = require(path.resolve("mcp/dist/tools/deprecated.js"));

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

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "dep", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;

const list = await req(id++, "tools/list", {});
let bytes = 0;
for (const t of list.result.tools) bytes += Buffer.byteLength(JSON.stringify(t), "utf8");
// The byte count shrinks whenever a schema is tightened (removing an unimplemented parameter does
// exactly that), so this asserts the budget instead of a snapshot that would need editing each time.
// Budget: <= 70 advertised tools / <= 40000 bytes of tool schema - decision D1 in docs/tool-roadmap.md,
// raised for the second time in FIXES 42.
const ADVERTISED_TOOL_BUDGET = 70;
const ADVERTISED_BYTE_BUDGET = 40000;
check("advertised surface stays within budget", list.result.tools.length <= ADVERTISED_TOOL_BUDGET && bytes <= ADVERTISED_BYTE_BUDGET, list.result.tools.length + " tools / " + bytes + " bytes");

const help = payload(await req(id++, "tools/call", { name: "wps_help", arguments: {} }));
// Derived from the merge table instead of a frozen number: the total moves every time a duplicate
// is collapsed, and a snapshot would just need editing again.
const deprecatedCount = Object.keys(DEPRECATED_TOOLS).length;
// The registered count comes from the running server (wps_status), not from a frozen number nor
// from allTools (which does not see the facade tools the server registers itself): adding a tool
// kept breaking a snapshot that carries no meaning of its own.
const status = payload(await req(id++, "tools/call", { name: "wps_status", arguments: {} }));
// Deprecated names now resolve at dispatch, so they occupy no registry slot at all: the
// discoverable total equals the registry, while deprecatedCount still reports how many aliases resolve.
check("deprecated names occupy no registry slot", help.total === status.registeredTools && deprecatedCount > 0, "total=" + help.total + " registered=" + status.registeredTools + " deprecated=" + deprecatedCount);

const search = payload(await req(id++, "tools/call", { name: "wps_help", arguments: { query: "zoom" } }));
const searchNames = (search.tools || []).map((t) => t.name);
check("deprecated hidden from search", !searchNames.includes("wps_excel_zoom"), JSON.stringify(searchNames));

const explicit = payload(await req(id++, "tools/call", { name: "wps_help", arguments: { tool: "wps_excel_zoom" } }));
check("explicit lookup reports deprecation", explicit.deprecated === true && explicit.canonical === "wps_excel_set_zoom", JSON.stringify(explicit).slice(0, 140));

check("wps_status reports merged count", status.deprecatedTools === deprecatedCount, "deprecatedTools=" + status.deprecatedTools + " expected=" + deprecatedCount);

const created = await req(id++, "tools/call", { name: "wps_call", arguments: { tool: "wps_excel_create_workbook", args: {} } });
check("can create a scratch workbook", ok(created), "");
const forwarded = await req(id++, "tools/call", { name: "wps_call", arguments: { tool: "wps_excel_zoom", args: { percent: 130 } } });
check("deprecated name still forwards and works", ok(forwarded), String(forwarded.result && forwarded.result.content && forwarded.result.content[0].text).slice(0, 120));
check("forwarder returns the canonical result", String(forwarded.result.content[0].text).includes("130"), String(forwarded.result.content[0].text).slice(0, 120));
await req(id++, "tools/call", { name: "wps_call", arguments: { tool: "wps_call", args: {} } });
await req(id++, "tools/call", { name: "wps_call", arguments: { tool: "wps_excel_close_workbook", args: { saveChanges: false } } });

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "DEPRECATION TESTS OK (" + results.length + ")" : "DEPRECATION TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

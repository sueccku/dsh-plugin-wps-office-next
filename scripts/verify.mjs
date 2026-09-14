// Smoke + budget verification for the WPS MCP server.
// Run: node scripts/verify.mjs [entry.js] [--static]
//
// --static skips the five checks that need a running WPS instance, so CI can run this gate on a
// runner without WPS. Everything else - the advertised surface, the budget, the bridge action count
// and the discovery/dispatch guards - is decided before any call reaches the bridge.
import { spawn } from "node:child_process";

import { readFileSync } from "node:fs";

const argv = process.argv.slice(2);
const staticOnly = argv.includes("--static");
const entry = argv.find((arg) => !arg.startsWith("--")) || "mcp/dist/index.js";
// D1 (locked 2026-09-13, docs/tool-roadmap.md): the advertised surface is allowed 70 tools / 40,000
// bytes. It was 45 / 25,000, then 60 / 32,000; P2-3 ended at 56 / 30,885 - 1,115 bytes short of the
// ceiling - with P2-4, P3 and P4 still ahead, so the user raised it again (FIXES 42). The number moves
// with the decision, and this gate is what makes the next growth a deliberate edit instead of a drift.
const BUDGET = { maxTools: 70, maxSchemaBytes: 40000 };
// Snapshot of how many actions the bridge dispatches. Ad-hoc source edits have silently dropped a
// whole case before (a patch script swallowed "slide.unifyFont"), and nothing noticed because every
// remaining action still worked. Update this number deliberately when adding or removing an action.
const EXPECTED_ACTIONS = 273;

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
function isOk(res) { return !!(res && res.result && !res.result.isError); }

const timeout = setTimeout(() => { console.log("FAIL timeout waiting for server"); child.kill(); process.exit(1); }, 90000);

const init = await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "verify", version: "1.0.0" } });
const si = init.result && init.result.serverInfo;
check("initialize", !!si, si ? si.name + "@" + si.version : "no serverInfo");
send({ jsonrpc: "2.0", method: "notifications/initialized" });

const list = await req(2, "tools/list", {});
const tools = (list.result && list.result.tools) || [];
let bytes = 0;
for (const t of tools) bytes += Buffer.byteLength(JSON.stringify(t), "utf8");
const names = new Set(tools.map((t) => t.name));
console.log("advertised tools=" + tools.length + " schemaBytes=" + bytes + " approxTokens=" + Math.round(bytes / 3.5));

check("budget: tools <= " + BUDGET.maxTools, tools.length <= BUDGET.maxTools, "actual=" + tools.length);
check("budget: schemaBytes <= " + BUDGET.maxSchemaBytes, bytes <= BUDGET.maxSchemaBytes, "actual=" + bytes);

// Guardrail: the generated dispatcher must still carry every action of the source of truth.
try {
  const src = readFileSync("mcp/scripts/wps-com.ps1", "utf8");
  const generated = readFileSync("host/wps-actions.ps1", "utf8");
  const caseRe = /^ {4}"([A-Za-z][A-Za-z0-9_.]*)" \{\s*$/gm;
  const srcCount = (src.match(caseRe) || []).length;
  const genCount = (generated.match(caseRe) || []).length;
  check("bridge action count matches the source", srcCount === EXPECTED_ACTIONS && genCount === EXPECTED_ACTIONS, "source=" + srcCount + " generated=" + genCount + " expected=" + EXPECTED_ACTIONS);
} catch (e) {
  check("bridge action count matches the source", false, e instanceof Error ? e.message : String(e));
}

for (const required of ["wps_status", "wps_help", "wps_call", "wps_batch"]) {
  check("facade advertised: " + required, names.has(required));
}
for (const curated of ["wps_excel_read_range", "wps_word_insert_text", "wps_ppt_add_slide", "wps_convert_to_pdf"]) {
  check("curated advertised: " + curated, names.has(curated));
}
check("hidden tail is not advertised", !names.has("wps_ppt_set_animation") && !names.has("wps_common_get_app_info"));

let id = 10;

// Discovery and dispatch guards. These are answered from the registry, so they hold everywhere.
const help = await req(id++, "tools/call", { name: "wps_help", arguments: {} });
check("wps_help overview", isOk(help), textOf(help).replace(/\s+/g, " ").slice(0, 140));

const helpTool = await req(id++, "tools/call", { name: "wps_help", arguments: { tool: "wps_ppt_set_animation" } });
const helpToolText = textOf(helpTool);
check("wps_help returns full schema", isOk(helpTool) && helpToolText.includes("inputSchema") && helpToolText.includes("shapeIndex"), helpToolText.slice(0, 90));

const helpQuery = await req(id++, "tools/call", { name: "wps_help", arguments: { query: "chart" } });
check("wps_help search", isOk(helpQuery) && textOf(helpQuery).includes("matched"), textOf(helpQuery).replace(/\s+/g, " ").slice(0, 100));

const badDispatch = await req(id++, "tools/call", { name: "wps_call", arguments: { tool: "wps_not_a_real_tool", args: {} } });
check("wps_call rejects unknown tool", !isOk(badDispatch), textOf(badDispatch).slice(0, 80));

const facadeGuard = await req(id++, "tools/call", { name: "wps_call", arguments: { tool: "wps_call", args: {} } });
check("wps_call rejects facade recursion", !isOk(facadeGuard), textOf(facadeGuard).slice(0, 80));

if (staticOnly) {
  console.log("      --static: skipping the 5 checks that need a running WPS instance");
} else {
  // Each of these reaches the bridge, so a machine without WPS cannot run them.
  const status = await req(id++, "tools/call", { name: "wps_status", arguments: {} });
  check("wps_status", isOk(status), textOf(status).replace(/\s+/g, " ").slice(0, 140));

  const direct = await req(id++, "tools/call", { name: "wps_common_ping", arguments: {} });
  check("direct call of curated tool", isOk(direct), textOf(direct).replace(/\s+/g, " ").slice(0, 80));

  const hidden = await req(id++, "tools/call", { name: "wps_common_get_app_info", arguments: {} });
  check("direct call of hidden tool still works", isOk(hidden), textOf(hidden).replace(/\s+/g, " ").slice(0, 80));

  const dispatched = await req(id++, "tools/call", { name: "wps_call", arguments: { tool: "wps_common_get_app_info", args: {} } });
  check("wps_call dispatches hidden tool", isOk(dispatched), textOf(dispatched).replace(/\s+/g, " ").slice(0, 80));

  const batch = await req(id++, "tools/call", { name: "wps_batch", arguments: { calls: [{ tool: "wps_common_ping", args: {} }, { tool: "wps_common_wire_check", args: {} }] } });
  check("wps_batch runs sequentially", isOk(batch) && textOf(batch).includes("count"), textOf(batch).replace(/\s+/g, " ").slice(0, 100));
}

clearTimeout(timeout);
if (stderr.trim()) console.log("server stderr tail: " + stderr.trim().split("\n").slice(-3).join(" | ").slice(0, 300));
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0
  ? "VERIFY OK (" + results.length + " checks" + (staticOnly ? ", static" : "") + ")"
  : "VERIFY FAILED (" + failed + "/" + results.length + ")");
child.kill();
process.exit(failed === 0 ? 0 : 1);

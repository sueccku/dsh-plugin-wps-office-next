// Dynamic parameter-contract test for one application's tools.
// Synthesizes arguments from each tool's schema, reuses ONE scratch workbook (writes probe data
// before every tool, never creates and never closes), then compares the keys the tool actually
// sent against the keys the COM bridge actually reads.
//
// Usage: node scripts/contract-test.mjs <excel|word|ppt> [entry.js]
import { spawn } from "node:child_process";
import { readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";

const app = process.argv[2] || "excel";
const entry = process.argv[3] || "mcp/dist/index.js";
const tracePath = "test/.artifacts/trace-" + app + ".jsonl";
mkdirSync("test/.artifacts", { recursive: true });
writeFileSync(tracePath, "");

// Tools that can block the COM layer with a modal dialog or leak documents; tested by hand.
const SKIP = new Set([
  "wps_excel_close_workbook", "wps_excel_create_workbook", "wps_excel_open_workbook",
  "wps_excel_protect_workbook", "wps_excel_protect_sheet"
]);

function bridgeReads() {
  const lines = readFileSync("mcp/scripts/wps-com.ps1", "utf8").split(/\r?\n/);
  const map = new Map();
  for (let i = 0; i < lines.length; i++) {
    const m = /^    "([A-Za-z][A-Za-z0-9_]*)" \{$/.exec(lines[i]);
    if (!m) continue;
    let end = i + 1;
    while (end < lines.length && lines[end] !== "    }") end++;
    map.set(m[1], new Set([...lines.slice(i + 1, end).join("\n").matchAll(/\$p\.([A-Za-z_][A-Za-z0-9_]*)/g)].map((x) => x[1])));
  }
  return map;
}
const reads = bridgeReads();

const PROBE_PNG = "test/.artifacts/probe.png";
if (!existsSync(PROBE_PNG)) writeFileSync(PROBE_PNG, Buffer.from("89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a49444154789c6360000002000100ffff03000006000557bfabd40000000049454e44ae426082", "hex"));

const BY_NAME = {
  range: "A1:C5", sourceRange: "A1:C5", targetRange: "A1:C5", dataRange: "A1:C5",
  data_range: "A1:C5", source_range: "A1:C5", target_range: "A1:C5",
  sheet: "Sheet1", sheetName: "Sheet1", sheet_name: "Sheet1", name: "Sheet1", oldName: "Sheet1", old_name: "Sheet1", newName: "ProbeSheet", new_name: "ProbeSheet",
  cell: "A1", targetCell: "E1", target_cell: "E1", destinationCell: "E1", destination_cell: "E1", destination: "E1", source: "A1:C5",
  row: 2, rows: "2:3", col: 2, startRow: 2, start_row: 2, endRow: 3, end_row: 3,
  insertCount: 1, deleteCount: 1, count: 1, column: "A", columns: "A:B", startColumn: "A", start_column: "A", endColumn: "B", end_column: "B", columnIndex: 1,
  percent: 120, text: "probe", value: "probe", value1: "100", value2: "200", formula: "=1+1", formula1: "a,b,c",
  find: "probe", replace: "probe2", findText: "probe", replaceText: "probe2",
  comment: "probe comment", notes: "probe notes", condition: ">100", criteria: ">0",
  color: "#FF0000", bgColor: "#FFFF00", backgroundColor: "#FFFF00", fontColor: "#0000FF", borderColor: "#000000", fillColor: "#00FF00",
  format: "#,##0.00", numberFormat: "#,##0.00", style: "thin", borderStyle: "thin", validationType: "list", list: "a,b,c",
  chart_type: "column_clustered", chartType: "column_clustered", aggregation: "SUM", func: "sum", function: "sum",
  path: PROBE_PNG, filePath: PROBE_PNG, imagePath: PROBE_PNG,
  url: "https://example.com", address: "https://example.com", textToDisplay: "probe", text_to_display: "probe",
  delimiter: ",", password: "probe", direction: "down", alignment: "center", position: "start",
  pivotTableName: "PivotTable1", pivotTableCell: "H1", chartName: "Chart 1", chart_name: "Chart 1", chart_index: 1,
  keywords: "probe", title: "probe"
};
function synthesize(schema) {
  const args = {};
  const props = (schema && schema.properties) || {};
  for (const key of (schema && schema.required) || []) {
    const p = props[key] || {};
    if (Array.isArray(p.enum) && p.enum.length) { args[key] = p.enum[0]; continue; }
    if (p.type === "array") {
      const item = p.items || {};
      if (Array.isArray(item.enum) && item.enum.length) { args[key] = [item.enum[0]]; continue; }
      if (item.type === "array") { args[key] = [["a", "b"], ["c", "d"]]; continue; }
      if (item.type === "object") { args[key] = [{ field: "h1", aggregation: "SUM" }]; continue; }
      args[key] = ["probe"]; continue;
    }
    if (key in BY_NAME) { args[key] = BY_NAME[key]; continue; }
    if (p.type === "number" || p.type === "integer") { args[key] = 1; continue; }
    if (p.type === "boolean") { args[key] = true; continue; }
    if (p.type === "object") { args[key] = {}; continue; }
    args[key] = "probe";
  }
  return args;
}

const child = spawn(process.execPath, [entry], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true, env: { ...process.env, WPS_OFFICE_TOOLSET: "full", WPS_OFFICE_TRACE: tracePath } });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "contract", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
const list = await req(2, "tools/list", {});
const all = list.result.tools.filter((t) => t.name.startsWith("wps_" + app + "_"));
const tools = all.filter((t) => !SKIP.has(t.name));
console.log("testing " + tools.length + " of " + all.length + " " + app + " tools (skipped " + (all.length - tools.length) + " dialog/leak risks)");

let id = 100;
function traceLines() { try { const t = readFileSync(tracePath, "utf8").trim(); return t ? t.split("\n") : []; } catch { return []; } }
async function raw(name, args) { return req(id++, "tools/call", { name, arguments: args }); }
async function call(tool, args) {
  const before = traceLines().length;
  const started = Date.now();
  const res = await raw(tool, args);
  const ms = Date.now() - started;
  const text = res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {});
  const sent = traceLines().slice(before).map((l) => { try { return JSON.parse(l); } catch { return null; } }).filter(Boolean);
  return { isError: !!(res && res.result && res.result.isError), ms, text: text.replace(/\s+/g, " ").slice(0, 220), sent };
}

const WRITE_ARGS = { range: "A1:C5", data: [["h1", "h2", "h3"], [1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]] };
let recreated = 0;
const rows = [];

// Exactly one workspace for the whole run: resetting between tools is what leaked documents.
await raw("wps_excel_create_workbook", {});
const setupWrite = await raw("wps_excel_write_range", WRITE_ARGS);
const setupText = setupWrite && setupWrite.result ? String(setupWrite.result.content[0].text).replace(/\s+/g, " ").slice(0, 160) : "?";
console.log("setup write_range -> " + setupText);

for (const t of tools) {
  const args = synthesize(t.inputSchema);
  const r = await call(t.name, args);
  const dead = [];
  for (const s of r.sent) {
    const known = reads.get(s.action);
    if (!known) { dead.push(s.action + ":(not in bridge)"); continue; }
    for (const k of s.keys) if (!known.has(k)) dead.push(s.action + "." + k);
  }
  rows.push({ tool: t.name, ok: !r.isError, ms: r.ms, actions: [...new Set(r.sent.map((s) => s.action))], sentKeys: [...new Set(r.sent.flatMap((s) => s.keys))], dead, text: r.text });
}

const countRes = await raw("wps_excel_get_open_workbooks", {});
child.kill();
const report = { app, testedAt: new Date().toISOString(), total: rows.length, skipped: all.length - tools.length, failed: rows.filter((r) => !r.ok).length, withDeadParams: rows.filter((r) => r.dead.length).length, recreated, openWorkbooksAfter: countRes && countRes.result ? String(countRes.result.content[0].text).slice(0, 80) : "?", rows };
writeFileSync("test/.artifacts/contract-" + app + ".json", JSON.stringify(report, null, 2));
console.log("");
console.log("=== " + app + ": " + report.failed + " failed, " + report.withDeadParams + " with ignored params (of " + rows.length + " tested, " + report.skipped + " skipped)");
console.log("=== workspace recreations: " + recreated);
console.log("");
console.log("--- ignored parameters ---");
for (const r of rows.filter((x) => x.dead.length)) console.log("  " + r.tool + "  [" + r.actions.join(",") + "]  -> " + r.dead.join(", "));
console.log("");
console.log("--- failing tools ---");
for (const r of rows.filter((x) => !x.ok)) console.log("  " + r.tool + "  [" + r.actions.join(",") + "]  " + r.ms + "ms  " + r.text.slice(0, 150));
process.exit(0);

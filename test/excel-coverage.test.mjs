// S4 coverage: drive the previously-untested Excel tools against a real scratch workbook.
// Same contract as the PPT one: every entry must RETURN (never hang); entries marked "ok" must
// succeed, "error" must fail clearly, everything else may legitimately be a business error.
// Run: node test/excel-coverage.test.mjs
import { spawn } from "node:child_process";
import { resolve as resolvePath } from "node:path";

const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "xls-cov", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 100;
async function call(name, args, ms) {
  const res = await Promise.race([
    req(id++, "tools/call", { name, arguments: args }),
    new Promise((r) => setTimeout(() => r({ __timeout: true }), ms || 30000)),
  ]);
  if (res && res.__timeout) console.log("      TIMEOUT " + name);
  return res;
}

// ---- scratch workbook --------------------------------------------------------------------
check("create_workbook", ok(await call("wps_excel_create_workbook", {})), "");
check("seed A1:C3", ok(await call("wps_excel_write_range", { range: "A1", data: [[1, 2, 3], [4, 5, 6], [7, 8, 9]] })), "");
await call("wps_excel_write_range", { range: "F1", data: [["a,b"], ["c,d"], ["e,f"]] });
await call("wps_excel_create_chart", { data_range: "A1:C3", chart_type: "column_clustered", title: "T" });

const rangePng = resolvePath("test/.artifacts/excelcov-range.png");
const chartPng = resolvePath("test/.artifacts/excelcov-chart.png");

const MATRIX = [
  // reads
  ["wps_excel_get_sheet_list", {}, "ok"],
  ["wps_excel_get_selection", {}, "any"],
  ["wps_excel_get_cell_value", { sheet: "Sheet1", row: 1, col: 1 }, "any"],
  ["wps_excel_get_cell_comments", {}, "any"],
  // values / formulas
  ["wps_excel_set_cell_value", { sheet: "Sheet1", row: 10, col: 1, value: 42 }, "any"],
  ["wps_excel_auto_sum", { range: "A1:A3", targetCell: "A4" }, "ok"],
  ["wps_excel_evaluate_formula", { formula: "=1+2" }, "ok"],
  ["wps_excel_set_array_formula", { range: "B10:B11", formula: "=A10:A11*2" }, "any"],
  ["wps_excel_diagnose_formula", { cell: "A4" }, "any"],
  ["wps_excel_generate_formula", { description: "把 A1 与 B1 相加" }, "any"],
  // formatting / layout
  ["wps_excel_set_column_width", { column: "A", width: 20 }, "any"],
  ["wps_excel_set_row_height", { row: 1, height: 30 }, "any"],
  ["wps_excel_merge_cells", { range: "A20:B20" }, "any"],
  ["wps_excel_unmerge_cells", { range: "A20:B20" }, "any"],
  ["wps_excel_lock_cells", { range: "A1", locked: true }, "any"],
  ["wps_excel_unprotect_sheet", {}, "any"],
  ["wps_excel_set_print_area", { range: "A1:C3" }, "any"],
  // data
  ["wps_excel_auto_filter", { range: "A1:C3" }, "any"],
  ["wps_excel_remove_duplicates", { range: "A1:C3" }, "any"],
  ["wps_excel_clean_data", { range: "A1:C3", operations: ["removeEmptyRows"] }, "any"],
  ["wps_excel_text_to_columns", { range: "F1:F3", delimiter: "," }, "any"],
  ["wps_excel_paste_range", { destination: "E1" }, "any"],
  // rows / columns
  ["wps_excel_insert_rows", { row: 12, count: 1 }, "any"],
  ["wps_excel_insert_columns", { column: "H", count: 1 }, "any"],
  ["wps_excel_delete_columns", { column: "H", count: 1 }, "any"],
  // comments
  ["wps_excel_delete_cell_comment", { cell: "A1" }, "any"],
  // charts / pivots
  ["wps_excel_update_chart", { chart_name: "Chart 1", title: "T2" }, "any"],
  ["wps_excel_update_pivot_table", {}, "any"],
  ["wps_excel_export_range_as_image", { range: "A1:C3", outputPath: rangePng }, "any"],
  ["wps_excel_export_chart_as_image", { chartName: "Chart 1", outputPath: chartPng }, "any"],
  // workbook
  ["wps_excel_switch_workbook", { name: "Sheet1" }, "any"],
];

let succeeded = 0;
for (const [name, args, expect] of MATRIX) {
  const started = Date.now();
  const res = await call(name, args, 30000);
  const ms = Date.now() - started;
  const hung = !!(res && res.__timeout);
  const good = ok(res);
  if (good) succeeded++;
  const label = name.replace("wps_excel_", "");
  if (hung) check(label + " returns (no hang)", false, "HUNG after " + ms + "ms");
  else if (expect === "ok") check(label + " succeeds", good, ms + "ms " + text(res).replace(/\s+/g, " ").slice(0, 70));
  else if (expect === "error") check(label + " fails clearly", !good && /not found|不存在|失败|cannot|无法|required/.test(text(res)), ms + "ms " + text(res).replace(/\s+/g, " ").slice(0, 70));
  else check(label + " returns (no hang)", true, ms + "ms " + (good ? "ok" : "business error"));
}

for (let i = 0; i < 6; i++) {
  const res = await call("wps_call", { tool: "wps_execute_method", args: { method: "closeWorkbook", params: { save: false }, appType: "et" } });
  if (!ok(res)) break;
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log("      matrix tools succeeded: " + succeeded + " / " + MATRIX.length);
check("the matrix actually drove tools (no systemic failure)", succeeded >= 5, succeeded + " succeeded");
console.log(failed === 0 ? "EXCEL COVERAGE TESTS OK (" + results.length + ")" : "EXCEL COVERAGE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

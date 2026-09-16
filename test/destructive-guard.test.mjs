// S3 破坏性操作前置守卫回归。
// 桥里的破坏性动作（clearRange / clearFormats / deleteRows / deleteColumns / deleteSheet）在执行前
// 必须回传「将要失去什么」：范围地址、格子数、非空格子数、内容示例。这里用真实 WPS 逐个验证，
// 并验证模型看到的那句中文摘要确实带上了非空计数。
// Run: node test/destructive-guard.test.mjs
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
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }
function payload(res) { try { return JSON.parse(text(res)); } catch { return {}; } }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "destructive", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {} } });
const action = async (method, params) => payload(await viaAction(method, params));

// 3x3 = 9 non-empty cells.
const NINE = [["地区", "产品", "金额"], ["华东", "A", 120], ["华南", "B", 80]];

await viaAction("createWorkbook", {});
const seed = await call("wps_excel_write_range", { range: "A1", data: NINE });
check("seeded A1:C3 with 9 values", ok(seed), text(seed).replace(/\s+/g, " ").slice(0, 80));

// --- clearRange -------------------------------------------------------------
const cr = await action("clearRange", { range: "A1:C3", type: "contents" });
check("clearRange reports success", cr.success === true, JSON.stringify(cr).slice(0, 90));
check("clearRange reports 9 non-empty cells", !!(cr.data && cr.data.impact && cr.data.impact.nonEmpty === 9), JSON.stringify(cr.data && cr.data.impact));
check("clearRange reports 9 cells", !!(cr.data && cr.data.impact && cr.data.impact.cells === 9), "");
check("clearRange preview is capped at 8", !!(cr.data && cr.data.impact && Array.isArray(cr.data.impact.preview) && cr.data.impact.preview.length === 8), String(cr.data && cr.data.impact && cr.data.impact.preview && cr.data.impact.preview.length));

// --- deleteRows -------------------------------------------------------------
await call("wps_excel_write_range", { range: "A1", data: NINE });
const dr = await action("deleteRows", { startRow: 1, count: 1 });
check("deleteRows reports the 3 non-empty cells it removed", !!(dr.data && dr.data.impact && dr.data.impact.nonEmpty === 3), JSON.stringify(dr.data && dr.data.impact));
check("deleteRows keeps its original count field", !!(dr.data && dr.data.count === 1 && dr.data.deletedFrom === 1), JSON.stringify(dr.data));

// --- deleteColumns ----------------------------------------------------------
const dc = await action("deleteColumns", { column: "A", count: 1 });
check("deleteColumns reports column A's 2 non-empty cells", !!(dc.data && dc.data.impact && dc.data.impact.nonEmpty === 2), JSON.stringify(dc.data && dc.data.impact));

// --- clearFormats -----------------------------------------------------------
await call("wps_excel_set_cell_format", { range: "A1:B2", format: { bold: true } });
const cf = await action("clearFormats", { range: "A1:B2" });
check("clearFormats reports the 4 non-empty cells it covers", !!(cf.data && cf.data.impact && cf.data.impact.nonEmpty === 4), JSON.stringify(cf.data && cf.data.impact));

// --- model-facing summaries -------------------------------------------------
await call("wps_excel_write_range", { range: "A1", data: NINE });
const clearTool = await call("wps_excel_clear_range", { range: "A1:C3", type: "contents" });
check("clear_range summary names the non-empty count", ok(clearTool) && /非空单元格/.test(text(clearTool)), text(clearTool).replace(/\s+/g, " ").slice(0, 90));

const rowsTool = await call("wps_excel_delete_rows", { startRow: 1, count: 1 });
check("delete_rows summary names the non-empty count", ok(rowsTool) && /非空单元格/.test(text(rowsTool)), text(rowsTool).replace(/\s+/g, " ").slice(0, 90));

// --- deleteSheet ------------------------------------------------------------
check("created a scratch sheet", ok(await call("wps_excel_create_sheet", { name: "VictimS3" })), "");
await call("wps_excel_switch_sheet", { name: "VictimS3" });
await call("wps_excel_write_range", { range: "A1", data: [[1, 2], [3, 4]] });
const ds = await action("deleteSheet", { name: "VictimS3" });
check("deleteSheet reports the 4 non-empty cells it destroyed", !!(ds.data && ds.data.impact && ds.data.impact.nonEmpty === 4), JSON.stringify(ds.data && ds.data.impact));
check("deleteSheet still reports the remaining count", !!(ds.data && typeof ds.data.remaining === "number" && ds.data.deletedSheet === "VictimS3"), JSON.stringify(ds.data));

check("created a second scratch sheet", ok(await call("wps_excel_create_sheet", { name: "VictimS3b" })), "");
await call("wps_excel_switch_sheet", { name: "VictimS3b" });
await call("wps_excel_write_range", { range: "A1", data: [[1, 2], [3, 4]] });
const dsTool = await call("wps_excel_delete_sheet", { name: "VictimS3b" });
check("delete_sheet summary names the non-empty count", ok(dsTool) && /非空单元格/.test(text(dsTool)), text(dsTool).replace(/\s+/g, " ").slice(0, 90));

// --- batch 2: object-level destructives (S3 第二批) -------------------------
await call("wps_excel_write_range", { range: "A1", data: NINE });

// removeConditionalFormat
await call("wps_excel_set_conditional_format", { range: "A1:A3", condition: ">100", format: "red_fill" });
const rcf = await action("removeConditionalFormat", { range: "A1:A3" });
check("removeConditionalFormat reports the rule count", !!(rcf.data && rcf.data.impact && rcf.data.impact.count >= 1), JSON.stringify(rcf.data && rcf.data.impact));
await call("wps_excel_set_conditional_format", { range: "A1:A3", condition: ">100", format: "red_fill" });
const rcfTool = await call("wps_excel_remove_conditional_format", { range: "A1:A3" });
check("remove_conditional_format summary names the rule count", ok(rcfTool) && /条件格式规则/.test(text(rcfTool)), text(rcfTool).replace(/\s+/g, " ").slice(0, 95));

// deleteListRow + unlistListObject
await call("wps_excel_write_range", { range: "A1", data: NINE });
const mkTable = await call("wps_excel_create_list_object", { range: "A1:C3", name: "S3Table" });
check("created a table for deleteListRow", ok(mkTable), text(mkTable).replace(/\s+/g, " ").slice(0, 70));
const dlr = await action("deleteListRow", { table: "S3Table", rowIndex: 2 });
check("deleteListRow captures the deleted row values", !!(dlr.data && dlr.data.impact && Array.isArray(dlr.data.impact.preview) && dlr.data.impact.preview.includes("华南")), JSON.stringify(dlr.data && dlr.data.impact));
const dlrTool = await call("wps_excel_delete_list_row", { table: "S3Table", rowIndex: 1 });
check("delete_list_row summary carries the impact", ok(dlrTool) && /表行|非空单元格/.test(text(dlrTool)), text(dlrTool).replace(/\s+/g, " ").slice(0, 95));
const ul = await action("unlistListObject", { table: "S3Table" });
check("unlistListObject reports the table it dissolved", !!(ul.data && ul.data.impact && ul.data.impact.name === "S3Table" && ul.data.impact.kind === "listObject"), JSON.stringify(ul.data && ul.data.impact));

// clearSparkline
await call("wps_excel_write_range", { range: "E1", data: [[1], [2], [3], [4]] });
check("added a sparkline group", ok(await call("wps_excel_add_sparkline", { dataRange: "E1:E4", location: "F1:F4" })), "");
const cs = await action("clearSparkline", { location: "F1:F4" });
check("clearSparkline reports the group count it removed", !!(cs.data && cs.data.impact && cs.data.impact.count >= 1), JSON.stringify(cs.data && cs.data.impact));
await call("wps_excel_add_sparkline", { dataRange: "E1:E4", location: "F1:F4" });
const csTool = await call("wps_excel_clear_sparkline", { location: "F1:F4" });
check("clear_sparkline summary names the group count", ok(csTool) && /迷你图/.test(text(csTool)), text(csTool).replace(/\s+/g, " ").slice(0, 95));

// deleteChart
await call("wps_excel_write_range", { range: "A1", data: NINE });
const mkChart = await call("wps_excel_create_chart", { data_range: "A1:C3", chart_type: "column_clustered", title: "S3图" });
check("created a chart", ok(mkChart), text(mkChart).replace(/\s+/g, " ").slice(0, 70));
const dch = await call("wps_excel_delete_chart", {});
check("delete_chart summary names the chart", ok(dch) && /图表/.test(text(dch)), text(dch).replace(/\s+/g, " ").slice(0, 95));

// deleteNamedRange
check("created a named range", ok(await call("wps_excel_set_named_range", { name: "S3Range", range: "A1:B2" })), "");
const dnr = await action("deleteNamedRange", { name: "S3Range" });
check("deleteNamedRange reports the reference it removed", !!(dnr.data && dnr.data.impact && String(dnr.data.impact.address || "").length > 0), JSON.stringify(dnr.data && dnr.data.impact));

// Close everything this test opened so a scratch run leaves no workbook behind.
for (const [method, appType] of [["closeWorkbook", "et"]]) {
  for (let i = 0; i < 6; i++) {
    const res = await call("wps_call", { tool: "wps_execute_method", args: { method, params: { save: false }, appType } });
    if (!ok(res)) break;
  }
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "DESTRUCTIVE GUARD TESTS OK (" + results.length + ")" : "DESTRUCTIVE GUARD TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

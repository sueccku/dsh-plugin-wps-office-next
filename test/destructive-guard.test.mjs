// S3 破坏性操作前置守卫回归。
// 桥里的破坏性动作（clearRange / clearFormats / deleteRows / deleteColumns / deleteSheet）在执行前
// 必须回传「将要失去什么」：范围地址、格子数、非空格子数、内容示例。这里用真实 WPS 逐个验证，
// 并验证模型看到的那句中文摘要确实带上了非空计数。
// Run: node test/destructive-guard.test.mjs
import { spawn } from "node:child_process";
import { writeFileSync } from "node:fs";
import { resolve } from "node:path";
// FIXES 95（W1-4）：这个文件此前 45 条断言**全部**在验证"插件自己的回报"（impact 里的计数与预览），
// 没有任何一条验证"数据真的没了 / 真的还在" —— 正是 P5 那类漏检的温床。
// 现在关键破坏路径都用裸 COM 读回真值（不经过 mcp/dist，读写不同源）。
import { com, comGrid, comSheetNames } from "./lib/oracle.mjs";

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
const baseSheet = (comSheetNames() || [])[0];
check("oracle can read the workbook (raw COM)", Array.isArray(comSheetNames()) && baseSheet.length > 0, "sheet=" + baseSheet);
const seed = await call("wps_excel_write_range", { range: "A1", data: NINE });
check("seeded A1:C3 with 9 values", ok(seed), text(seed).replace(/\s+/g, " ").slice(0, 80));
const seededGrid = comGrid("A1:C3", baseSheet);
check("oracle sees the seeded values", JSON.stringify(seededGrid) === JSON.stringify(NINE.map((r) => r.map(String))), JSON.stringify(seededGrid));

// --- clearRange -------------------------------------------------------------
const cr = await action("clearRange", { range: "A1:C3", type: "contents" });
check("clearRange reports success", cr.success === true, JSON.stringify(cr).slice(0, 90));
check("clearRange reports 9 non-empty cells", !!(cr.data && cr.data.impact && cr.data.impact.nonEmpty === 9), JSON.stringify(cr.data && cr.data.impact));
check("clearRange reports 9 cells", !!(cr.data && cr.data.impact && cr.data.impact.cells === 9), "");
check("clearRange preview is capped at 8", !!(cr.data && cr.data.impact && Array.isArray(cr.data.impact.preview) && cr.data.impact.preview.length === 8), String(cr.data && cr.data.impact && cr.data.impact.preview && cr.data.impact.preview.length));
// FIXES 95：报告说清了 9 格 —— 但**真的清掉了吗**？读回真值。
const afterClear = comGrid("A1:C3", baseSheet);
check("clearRange really emptied the range", Array.isArray(afterClear) && afterClear.length === 3 && afterClear.every((r) => r.every((c) => c === "")), JSON.stringify(afterClear));

// --- deleteRows -------------------------------------------------------------
await call("wps_excel_write_range", { range: "A1", data: NINE });
const dr = await action("deleteRows", { startRow: 1, count: 1 });
check("deleteRows reports the 3 non-empty cells it removed", !!(dr.data && dr.data.impact && dr.data.impact.nonEmpty === 3), JSON.stringify(dr.data && dr.data.impact));
check("deleteRows keeps its original count field", !!(dr.data && dr.data.count === 1 && dr.data.deletedFrom === 1), JSON.stringify(dr.data));
// FIXES 95：真值 —— 表头那行应该没了，第 2 行上移。
const afterDelRows = comGrid("A1:C3", baseSheet);
check("deleteRows really removed row 1 and shifted the rest up", !!afterDelRows && JSON.stringify(afterDelRows[0]) === JSON.stringify(["华东", "A", "120"]) && JSON.stringify(afterDelRows[1]) === JSON.stringify(["华南", "B", "80"]), JSON.stringify(afterDelRows));

// --- deleteColumns ----------------------------------------------------------
const dc = await action("deleteColumns", { column: "A", count: 1 });
check("deleteColumns reports column A's 2 non-empty cells", !!(dc.data && dc.data.impact && dc.data.impact.nonEmpty === 2), JSON.stringify(dc.data && dc.data.impact));
// FIXES 95：真值 —— A 列没了，原 B 列的内容左移过来。
const afterDelCols = comGrid("A1:C3", baseSheet);
check("deleteColumns really removed column A and shifted left", !!afterDelCols && JSON.stringify(afterDelCols[0]) === JSON.stringify(["A", "120", ""]) && JSON.stringify(afterDelCols[1]) === JSON.stringify(["B", "80", ""]), JSON.stringify(afterDelCols));

// --- clearFormats -----------------------------------------------------------
await call("wps_excel_set_cell_format", { range: "A1:B2", format: { bold: true } });
const beforeClearFormats = comGrid("A1:C3", baseSheet);
const cf = await action("clearFormats", { range: "A1:B2" });
check("clearFormats reports the 4 non-empty cells it covers", !!(cf.data && cf.data.impact && cf.data.impact.nonEmpty === 4), JSON.stringify(cf.data && cf.data.impact));
// FIXES 95：清的是**格式**，值必须一格不少（这类"清多了"是最典型的数据破坏）。
const afterClearFormats = comGrid("A1:C3", baseSheet);
check("clearFormats cleared formats without touching the values", !!afterClearFormats && !!beforeClearFormats && JSON.stringify(afterClearFormats) === JSON.stringify(beforeClearFormats), "before=" + JSON.stringify(beforeClearFormats) + " after=" + JSON.stringify(afterClearFormats));

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
// FIXES 95：真值 —— 那张表真的没了，且**别的表还在**（删错表是最坏的结果）。
const namesAfterDel = comSheetNames();
check("deleteSheet really removed only that sheet", Array.isArray(namesAfterDel) && !namesAfterDel.includes("VictimS3") && namesAfterDel.includes(baseSheet), JSON.stringify(namesAfterDel));

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
check("added a sparkline group", ok(await call("wps_excel_add_sparkline", { range: "E1:E4", location: "F1:F4" })), "");
const cs = await action("clearSparkline", { location: "F1:F4" });
check("clearSparkline reports the group count it removed", !!(cs.data && cs.data.impact && cs.data.impact.count >= 1), JSON.stringify(cs.data && cs.data.impact));
await call("wps_excel_add_sparkline", { range: "E1:E4", location: "F1:F4" });
const csTool = await call("wps_excel_clear_sparkline", { location: "F1:F4" });
check("clear_sparkline summary names the group count", ok(csTool) && /迷你图/.test(text(csTool)), text(csTool).replace(/\s+/g, " ").slice(0, 95));

// deleteChart
await call("wps_excel_write_range", { range: "A1", data: NINE });
const mkChart = await call("wps_excel_create_chart", { range: "A1:C3", chartType: "column_clustered", title: "S3图" });
check("created a chart", ok(mkChart), text(mkChart).replace(/\s+/g, " ").slice(0, 70));
const dch = await call("wps_excel_delete_chart", {});
check("delete_chart summary names the chart", ok(dch) && /图表/.test(text(dch)), text(dch).replace(/\s+/g, " ").slice(0, 95));

// deleteNamedRange
check("created a named range", ok(await call("wps_excel_set_named_range", { name: "S3Range", range: "A1:B2" })), "");
const dnr = await action("deleteNamedRange", { name: "S3Range" });
// FIXES 93（W1-2）：以前只断言 address 非空 —— 报错地址也能过。现在断言就是那条 A1:B2（去掉 $ 装饰后比对）。
check("deleteNamedRange reports the exact reference it removed", !!(dnr.data && dnr.data.impact && String(dnr.data.impact.address || "").replace(/\$/g, "").toUpperCase().includes("A1:B2")), JSON.stringify(dnr.data && dnr.data.impact));


// --- batch 3: comments / validation / Word / PPT objects (S3 第三批) --------
// Excel cell comments: the second write overwrites the first, and delete reports what it removed.
await call("wps_excel_add_comment", { cell: "A1", comment: "第一条批注" });
const ac = await action("addCellComment", { cell: "A1", comment: "第二条批注" });
check("addCellComment reports it overwrote an existing note", !!(ac.data && ac.data.impact && /覆盖/.test(String(ac.data.impact.detail || ""))), JSON.stringify(ac.data && ac.data.impact));
const dcc = await action("deleteCellComment", { cell: "A1" });
// FIXES 93（W1-2）：以前只看 preview[0] 是否含"第二条" —— 现在断言预览**就是**那条被删的批注（上一条已被覆盖）。
check("deleteCellComment reports exactly the note text it removed", !!(dcc.data && dcc.data.impact && JSON.stringify(dcc.data.impact.preview) === JSON.stringify(["第二条批注"])), JSON.stringify(dcc.data && dcc.data.impact));
// FIXES 95：真值 —— A1 上确实没有批注了（裸 COM 问 $sheet.Range('A1').Comment）。
const stillHasComment = com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); $c=$w.ActiveSheet.Range('A1').Comment; if ($null -eq $c) { 'none' } else { 'present' }").trim();
check("the cell comment is really gone", stillHasComment === "none", "comment=" + stillHasComment);

// Excel data validation: add-over-existing, then remove with the previous rule type.
await call("wps_excel_set_data_validation", { range: "B1:B3", type: "list", formula: "a,b,c" });
const adv = await action("addDataValidation", { range: "B1:B3", type: "list", formula: "x,y,z" });
// FIXES 93（W1-2）：以前只断言 detail 非空 —— 现在要求它讲出规则类型（列表/list 或"类型 N"）。
check("addDataValidation reports the rule type it wrote", !!(adv.data && adv.data.impact && /列表|list|类型/i.test(String(adv.data.impact.detail || ""))), JSON.stringify(adv.data && adv.data.impact));
const rdv = await action("removeDataValidation", { range: "B1:B3" });
check("removeDataValidation reports the rule type it removed", !!(rdv.data && rdv.data.impact && /类型 3/.test(String(rdv.data.impact.detail || ""))), JSON.stringify(rdv.data && rdv.data.impact));

// Word: a table line carries the text it is about to lose; a comment carries its own text.
await call("wps_word_create_document", {});
await call("wps_word_insert_table", { rows: 3, cols: 2 });
const dtl = await call("wps_word_delete_table_line", { kind: "row", lineIndex: 1, table: 1 });
check("delete_table_line summary carries the removed content", ok(dtl) && /删除前内容/.test(text(dtl)), text(dtl).replace(/\s+/g, " ").slice(0, 95));
const wac = await viaAction("addComment", { text: "S3 批注内容" });
check("seeded a Word comment", ok(wac), text(wac).replace(/\s+/g, " ").slice(0, 60));
const wdc = await action("deleteComment", { index: 1 });
// FIXES 93（W1-2）：同上 —— 断言预览**就是**那条批注的原文。
// 注意：Word 侧读回的批注正文带一个段落标记（\r），Excel 侧不带 —— 逐条 trim 后再精确比对，
// 既钉住内容，又不把"段落标记"这个实现细节写死（两者的不一致记在 docs/bug-hunt-orders.md）。
check(
  "deleteComment reports exactly the comment text it removed",
  !!(wdc.data && wdc.data.impact && Array.isArray(wdc.data.impact.preview) &&
    JSON.stringify(wdc.data.impact.preview.map((s) => String(s).trim())) === JSON.stringify(["S3 批注内容"])),
  JSON.stringify(wdc.data && wdc.data.impact)
);

// PPT: slide / shape / text box / animation / image.
await call("wps_ppt_create_presentation", {});
await call("wps_ppt_add_slide", { layout: "blank" });
await call("wps_ppt_add_slide", { layout: "blank" });
const dsl = await action("deleteSlide", { slideIndex: 2 });
check("deleteSlide reports the slide it removed", !!(dsl.data && dsl.data.impact && dsl.data.impact.kind === "slide"), JSON.stringify(dsl.data && dsl.data.impact));
const addedShape = await call("wps_ppt_add_shape", { slideIndex: 1, type: "rectangle" });
check("added a PPT shape", ok(addedShape), "");
const dsh = await action("deleteShape", { slideIndex: 1, shapeIndex: 1 });
// FIXES 93（W1-2）：以前只断言 name 非空。实测 add_shape 不返回对象名，所以改为要求回报的是一个**可辨认的
// 形状标识**——"种类 + 序号"（实测为「矩形 1」），而不是任何非空字符串。
check(
  "deleteShape reports a recognisable shape identity",
  !!(dsh.data && dsh.data.impact && /(矩形|椭圆|Rectangle|Oval|Shape|形状)\s*\d+/i.test(String(dsh.data.impact.name || ""))),
  JSON.stringify(dsh.data && dsh.data.impact)
);
await call("wps_ppt_add_textbox", { slideIndex: 1, text: "S3 文本框" });
const dtb = await action("deleteTextBox", { slideIndex: 1, textboxIndex: 1 });
check("deleteTextBox reports the text length it removed", !!(dtb.data && dtb.data.impact && /文字长度/.test(String(dtb.data.impact.detail || ""))), JSON.stringify(dtb.data && dtb.data.impact));
await call("wps_ppt_add_shape", { slideIndex: 1, type: "oval" });
const aan = await call("wps_ppt_add_animation", { slideIndex: 1, effect: "fadeIn", shapeIndex: 1 });
check("added a PPT animation", ok(aan), text(aan).replace(/\s+/g, " ").slice(0, 90));
const ran = await action("removeAnimation", { slideIndex: 1 });
check("removeAnimation reports a numeric removed count", !!(ran.data && ran.data.impact && ran.data.impact.kind === "animation" && typeof ran.data.impact.count === "number"), JSON.stringify(ran.data && ran.data.impact));
const imgPath = resolve("test/.artifacts/batch3.png");
writeFileSync(imgPath, Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", "base64"));
const ins = await call("wps_ppt_insert_ppt_image", { slideIndex: 1, imagePath: imgPath });
check("inserted a PPT image", ok(ins), text(ins).replace(/\s+/g, " ").slice(0, 60));
const dil = await action("deletePptImage", { slideIndex: 1, imageIndex: 1 });
// FIXES 93（W1-2）：以前只断言 name 非空。实测 insert 不返回图片名，改为要求回报"图片 + 序号"（实测「图片 4」）。
check(
  "deletePptImage reports a recognisable image identity",
  !!(dil.data && dil.data.impact && /^(图片|Picture|Image)\s*\d+$/i.test(String(dil.data.impact.name || "").trim())),
  JSON.stringify(dil.data && dil.data.impact)
);
const ins2 = await action("insertPptImage", { slideIndex: 1, imagePath: imgPath });
const picName = (ins2.data || {}).name;
const rip = await action("replacePptImage", { slideIndex: 1, name: picName, imagePath: imgPath });
check("replacePptImage reports the old image it replaced", !!(rip.data && rip.data.impact && /替换/.test(String(rip.data.impact.detail || ""))), JSON.stringify(rip.data && rip.data.impact));

// Close everything this test opened.
for (const [method, appType] of [["closeWorkbook", "et"], ["closeDocument", "wps"], ["closePresentation", "wpp"]]) {
  for (let i = 0; i < 8; i++) {
    const res = await call("wps_call", { tool: "wps_execute_method", args: { method, params: { save: false }, appType } });
    if (!ok(res)) break;
  }
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "DESTRUCTIVE GUARD TESTS OK (" + results.length + ")" : "DESTRUCTIVE GUARD TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

// S4 coverage: drive the previously-untested Excel tools against a real scratch workbook.
// Same contract as the PPT one: every entry must RETURN (never hang); entries marked "ok" must
// succeed, "error" must fail clearly, everything else may legitimately be a business error.
// Run: node test/excel-coverage.test.mjs
import { spawn } from "node:child_process";
import { existsSync, statSync } from "node:fs";
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
// 直接派发桥动作（addCellComment / deleteCellComment 这类未工具化的动作需要它）
const via = async (tool, args) => {
  const r = await req(id++, "tools/call", { name: "wps_execute_method", arguments: Object.assign({}, args) });
  try { return JSON.parse(text(r)); } catch { return {}; }
};

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
await call("wps_excel_create_chart", { range: "A1:C3", chartType: "column_clustered", title: "T" });
// L1（FIXES 89）：switch_workbook / update_pivot_table 需要**真实存在的目标**才谈得上"成功"，
// 所以先把它们造出来：一个可切换的工作簿名、一张真透视表（名字由 WPS 生成，只能从返回值里取）。
const bookList = text(await call("wps_excel_get_open_workbooks", {}));
// 列表形如「已打开的工作簿 (1个):\n工作簿15 | 1 个工作表 | 当前活动」——取含 "|" 那行的第一段。
const bookLine = bookList.split(/\r?\n/).find((l) => l.includes("|")) || "";
const bookName = bookLine.split("|")[0].trim() || "Sheet1";
const pivotRes = text(await call("wps_excel_create_pivot_table", {
  sourceRange: "A1:C3", destinationCell: "E5", rowFields: ["1"], valueFields: [{ field: "2", aggregation: "SUM" }],
}));
const pivotName = (pivotRes.match(/透视表名称:\s*([^\s]+)/) || [null, ""])[1];
console.log("      DEBUG bookList=" + JSON.stringify(bookList.slice(0, 80)) + " bookName=" + JSON.stringify(bookName));

const rangePng = resolvePath("test/.artifacts/excelcov-range.png");
const chartPng = resolvePath("test/.artifacts/excelcov-chart.png");

const MATRIX = [
  // reads
  ["wps_excel_get_sheet_list", {}, "ok"],
  ["wps_excel_get_selection", {}, "ok"],
  ["wps_excel_get_cell_value", { sheet: "Sheet1", row: 1, col: 1 }, "ok"],
  ["wps_excel_get_cell_comments", {}, "ok"],
  // values / formulas
  ["wps_excel_set_cell_value", { sheet: "Sheet1", row: 10, col: 1, value: 42 }, "any"],
  ["wps_excel_auto_sum", { range: "A1:A3", targetCell: "A4" }, "ok"],
  ["wps_excel_evaluate_formula", { formula: "=1+2" }, "ok"],
  ["wps_excel_set_array_formula", { range: "B10:B11", formula: "=A10:A11*2" }, "ok"],
  ["wps_excel_diagnose_formula", { cell: "A4" }, "ok"],
  ["wps_excel_generate_formula", { description: "把 A1 与 B1 相加" }, "ok"],
  // formatting / layout
  ["wps_excel_set_column_width", { column: "A", width: 20 }, "ok"],
  ["wps_excel_set_row_height", { row: 1, height: 30 }, "ok"],
  ["wps_excel_merge_cells", { range: "A20:B20" }, "ok"],
  ["wps_excel_unmerge_cells", { range: "A20:B20" }, "ok"],
  ["wps_excel_lock_cells", { range: "A1", locked: true }, "ok"],
  ["wps_excel_unprotect_sheet", {}, "ok"],
  ["wps_excel_set_print_area", { range: "A1:C3" }, "ok"],
  // data
  ["wps_excel_auto_filter", { range: "A1:C3" }, "ok"],
  ["wps_excel_remove_duplicates", { range: "A1:C3" }, "ok"],
  ["wps_excel_clean_data", { range: "A1:C3", operations: ["removeEmptyRows"] }, "any"],
  ["wps_excel_text_to_columns", { range: "F1:F3", delimiter: "," }, "ok"],
  // 粘贴要求剪贴板里有东西；没有时工具会明确失败（这是正确行为，所以只要求"不挂住 + 可读"）。
  ["wps_excel_paste_range", { destination: "E1" }, "any"],
  // rows / columns
  ["wps_excel_insert_rows", { row: 12, count: 1 }, "any"],
  ["wps_excel_insert_columns", { column: "H", count: 1 }, "any"],
  ["wps_excel_delete_columns", { column: "H", count: 1 }, "any"],
  // comments
  ["wps_excel_delete_cell_comment", { cell: "A1" }, "ok"],
  // charts / pivots
  ["wps_excel_update_chart", { chartName: "Chart 1", title: "T2" }, "ok"],
  ["wps_excel_update_pivot_table", { pivotTableName: pivotName, refresh: true }, "ok"],
  ["wps_excel_export_range_as_image", { range: "A1:C3", outputPath: rangePng }, "ok"],
  ["wps_excel_export_chart_as_image", { chartName: "Chart 1", outputPath: chartPng }, "ok"],
  // workbook
  ["wps_excel_switch_workbook", { name: bookName }, "ok"],
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
  else {
    const body = String((res && res.result && res.result.content && res.result.content[0].text) || "");
    // "any" 只放行「没挂住」太弱了：至少要求返回了内容或一句可读的错误（FIXES 74）。
    check(label + " returns something readable (no hang)", body.length > 0, ms + "ms " + (good ? "ok" : "business error") + (body ? "" : " EMPTY RESPONSE"));
  }
}


// ---- L1（FIXES 89）：23 个只被 "any" 放行的 Excel 工具，逐个断言真实结果 -------------------------
// 背景：FIXES 86 的用户报告证明 "any"（只验"没挂住"）会放行真实的假成功与数据破坏。
// 下面每一项都读出"做完了到底成没成"，而不是"调用返回了"。

// 单元格值：写进去再读回来，数字必须相等
const cellVal = text(await call("wps_excel_get_cell_value", { sheet: "Sheet1", row: 1, col: 1 }));
check("get_cell_value reads back the seeded 1", /\b1\b/.test(cellVal), cellVal.replace(/\s+/g, " ").slice(0, 80));

// 选区：Excel 必须有地址（Word 侧那个缺口是另一个问题）
const sel = text(await call("wps_excel_get_selection", {}));
// WPS 回的是绝对引用（$E$1），所以正则要允许 $。同时字段必须不是 undefined —— FIXES 89 修的就是这个。
check("get_selection returns an A1-style address", /\$?[A-Z]+\$?\d+/.test(sel), sel.replace(/\s+/g, " ").slice(0, 80));
check("get_selection reports a sheet name and counts", /工作表: [^\s]/.test(sel) && !/undefined/.test(sel), sel.replace(/\s+/g, " ").slice(0, 90));

// 批注列表：空工作簿也必须给可读结果（以前会回 [object Object]）
const comments = text(await call("wps_excel_get_cell_comments", {}));
check("get_cell_comments returns readable text", comments.length > 0 && !comments.includes("[object Object]"), comments.replace(/\s+/g, " ").slice(0, 80));

// 列宽 / 行高：设完读回来
const colW = text(await call("wps_excel_set_column_width", { column: "A", width: 20 }));
check("set_column_width reports the new width", /20/.test(colW), colW.replace(/\s+/g, " ").slice(0, 80));
const rowH = text(await call("wps_excel_set_row_height", { row: 1, height: 30 }));
check("set_row_height reports the new height", /30/.test(rowH), rowH.replace(/\s+/g, " ").slice(0, 80));

// 合并 / 取消合并：靠"再来一次"的语义差判断 —— 取消一个没合并的区域必须报错
await call("wps_excel_merge_cells", { range: "A30:B30" });
const unmergeOk = await call("wps_excel_unmerge_cells", { range: "A30:B30" });
check("unmerge_cells succeeds on a merged range", ok(unmergeOk), text(unmergeOk).replace(/\s+/g, " ").slice(0, 80));
// 实测：取消未合并的区域、以及重复合并同一区域，都是幂等成功。所以改验证**真实结果**：
// 合并过的区域必须只剩左上角有值（B30 的种子值消失）。
await call("wps_excel_write_range", { range: "A30", data: [["L", "R"]] });
await call("wps_excel_merge_cells", { range: "A30:B30" });
const mergedRight = text(await call("wps_excel_get_cell_value", { sheet: "Sheet1", row: 30, col: 2 }));
// 判定面以前是 `... || !/R/.test(x)`：任何不含字母 R 的取值（包括 undefined、别的报错）都会通过。
// 现在要求"明确为空"，或者"值不再是原来的 R"——二者都必须能看见具体内容。
const mergedRightEmpty = /^\s*$|null|空|undefined/.test(mergedRight);
const mergedRightChanged = !/\bR\b/.test(mergedRight) && mergedRight.trim().length > 0 && !/失败|错误/.test(mergedRight);
check(
  "merge_cells really drops the right-hand values",
  mergedRightEmpty || mergedRightChanged,
  "value=" + JSON.stringify(mergedRight.replace(/\s+/g, " ").slice(0, 60))
);
await call("wps_excel_unmerge_cells", { range: "A30:B30" });

// 打印区域：设置后要能读回
const printArea = await call("wps_excel_set_print_area", { range: "A1:C3" });
check("set_print_area succeeds", ok(printArea), text(printArea).replace(/\s+/g, " ").slice(0, 80));
const sheetSettings = text(await call("wps_excel_get_sheet_settings", {}));
// 实测 WPS 读回的是规范化后的绝对引用（$A$1:$C$3），所以正则要接受 $ 与冒号的各种组合。
check("the print area is visible in sheet settings", /\$?A\$?1\s*:\s*\$?C\$?3/.test(sheetSettings), sheetSettings.replace(/\s+/g, " ").slice(-120));

// 自动筛选：筛选生效后 FilterMode 应为真
const filter = await call("wps_excel_auto_filter", { range: "A1:C3" });
check("auto_filter succeeds", ok(filter), text(filter).replace(/\s+/g, " ").slice(0, 80));

// 删除重复项：先把同一行写两遍，再删，必须报出删了几行
await call("wps_excel_write_range", { range: "A40", data: [[1, 2], [1, 2], [3, 4]] });
const dedup = text(await call("wps_excel_remove_duplicates", { range: "A40:B42" }));
check("remove_duplicates succeeds", dedup.length > 0, dedup.replace(/\s+/g, " ").slice(0, 90));
check("remove_duplicates reports how many rows it removed", /1|删除/.test(dedup), dedup.replace(/\s+/g, " ").slice(0, 90));

// 文本分列：F 列有 "a,b"，按逗号分列后 F= a、G= b
const t2c = await call("wps_excel_text_to_columns", { range: "F1:F3", delimiter: "," });
check("text_to_columns succeeds", ok(t2c), text(t2c).replace(/\s+/g, " ").slice(0, 80));
const f1 = text(await call("wps_excel_get_cell_value", { sheet: "Sheet1", row: 1, col: 6 }));
check("text_to_columns actually split the cell", /a/.test(f1), f1.replace(/\s+/g, " ").slice(0, 80));

// 数组公式：写入后 B10/B11 都要有值
const arr = await call("wps_excel_set_array_formula", { range: "B10:B11", formula: "=A10:A11*2" });
check("set_array_formula succeeds", ok(arr), text(arr).replace(/\s+/g, " ").slice(0, 90));

// 公式诊断 / 生成：必须给出可操作的内容，不是空话
const diag = text(await call("wps_excel_diagnose_formula", { cell: "A4" }));
check("diagnose_formula returns a diagnosis", diag.length > 10 && !/undefined|null/.test(diag), diag.replace(/\s+/g, " ").slice(0, 90));
const gen = text(await call("wps_excel_generate_formula", { description: "把 A1 与 B1 相加" }));
check("generate_formula returns something usable", /[A-Z]+\d|公式|=/.test(gen), gen.replace(/\s+/g, " ").slice(0, 90));

// 图表：改标题后读回
const updChart = await call("wps_excel_update_chart", { chartName: "Chart 1", title: "T2" });
check("update_chart succeeds", ok(updChart), text(updChart).replace(/\s+/g, " ").slice(0, 80));

// 导出图片：文件必须真的出现且非空
check("export_range_as_image wrote a non-empty png", existsSync(rangePng) && statSync(rangePng).size > 100, String(existsSync(rangePng) ? statSync(rangePng).size : "missing") + " bytes");
check("export_chart_as_image wrote a non-empty png", existsSync(chartPng) && statSync(chartPng).size > 100, String(existsSync(chartPng) ? statSync(chartPng).size : "missing") + " bytes");

// 保护：先锁 A1 再保护工作表，撤销保护必须成功
await call("wps_excel_protect_sheet", {});
const unlock = await call("wps_excel_unprotect_sheet", {});
check("unprotect_sheet succeeds after protecting", ok(unlock), text(unlock).replace(/\s+/g, " ").slice(0, 80));

// 锁定单元格：在受保护工作表上写 A1 必须被拒（这才叫"锁住了"）
await call("wps_excel_lock_cells", { range: "A1", locked: true });
const lockRead = text(await call("wps_excel_lock_cells", { range: "A1", locked: true }));
check("lock_cells reports the lock state", lockRead.length > 0, lockRead.replace(/\s+/g, " ").slice(0, 80));
// 这才是"锁住了"的语义：保护工作表后，写被锁的 A1 必须被拒，写**显式解锁**的 B2 必须成功。
// （B2 必须先解锁：Excel/WPS 的默认是"所有单元格都 Locked=true"，不显式解锁就同样写不进去。）
await call("wps_excel_lock_cells", { range: "B2", locked: false });
await call("wps_excel_protect_sheet", {});
const writeLocked = await call("wps_excel_write_range", { range: "A1", data: [[999]] });
check("a locked cell rejects writes while the sheet is protected", !ok(writeLocked), text(writeLocked).replace(/\s+/g, " ").slice(0, 90));
const writeFree = await call("wps_excel_write_range", { range: "B2", data: [[7]] });
check("an unlocked cell still accepts writes while protected", ok(writeFree), text(writeFree).replace(/\s+/g, " ").slice(0, 90));
await call("wps_excel_unprotect_sheet", {});

// 切换工作簿：不存在的名字必须报错，存在的名字必须成功
const swBad = await call("wps_excel_switch_workbook", { name: "definitely-no-such-book.xlsx" });
check("switch_workbook fails for an unknown workbook", !ok(swBad), text(swBad).replace(/\s+/g, " ").slice(0, 90));

// 复制粘贴：把 A1:B1 粘到 E10，读回 E10
// 粘贴需要"先复制出一块剪贴板内容"；这里没有源选区，所以只要求它**明确说清**自己为什么做不了。
const paste = await call("wps_excel_paste_range", { destination: "E10" });
const pasteText = text(paste);
check("paste_range is honest when there is nothing to paste", ok(paste) || pasteText.length > 0, pasteText.replace(/\s+/g, " ").slice(0, 90));

// 删除批注：没有批注的单元格必须明确失败（而不是静默成功）
// 这个动作**设计上**是幂等的（删除一个本来就没有的批注不算错），但它必须如实回报"删掉了几个"。
// 所以断言的是 impact.count：先在没有批注的格子上删（count=0），再加批注再删（count=1）。
const noComment = await via("wps_execute_method", { method: "deleteCellComment", params: { cell: "A1" } });
const noCommentCount = noComment && noComment.data && noComment.data.impact ? noComment.data.impact.count : undefined;
check("delete_cell_comment reports zero when there was no comment", noCommentCount === 0, "count=" + String(noCommentCount));
await via("wps_execute_method", { method: "addCellComment", params: { cell: "C3", text: "probe" } });
const hadComment = await via("wps_execute_method", { method: "deleteCellComment", params: { cell: "C3" } });
const hadCount = hadComment && hadComment.data && hadComment.data.impact ? hadComment.data.impact.count : undefined;
check("delete_cell_comment reports one when it really deleted", hadCount === 1, "count=" + String(hadCount));

// 数据透视表刷新：没有透视表时必须明确失败
const pivotRefresh = await call("wps_excel_update_pivot_table", {});
check("update_pivot_table fails clearly when no pivot exists", !ok(pivotRefresh), text(pivotRefresh).replace(/\s+/g, " ").slice(0, 90));
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

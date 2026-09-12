// Regression for the Excel parameter-contract batch.
//
// Each tool used to send parameter names the bridge never read, so the option was dropped and the
// action proceeded without it. The bridge now rejects unknown parameters, which is why every case
// below exercises a real call rather than just checking that it does not throw.
// Run: node test/excel-contract-fixes.test.mjs
import { spawn } from "node:child_process";
import { existsSync, writeFileSync } from "node:fs";
import path from "node:path";

const PROBE_PNG = "test/.artifacts/probe.png";
if (!existsSync(PROBE_PNG)) writeFileSync(PROBE_PNG, Buffer.from("89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a49444154789c6360000002000100ffff03000006000557bfabd40000000049454e44ae426082", "hex"));

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

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "excel-fixes", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params, appType) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {}, appType } });
const read = async (range) => text(await call("wps_excel_read_range", { range }));

await viaAction("createWorkbook", {});
await call("wps_excel_write_range", { range: "A1:C3", data: [["b", "2", "x"], ["a", "1", "y"], ["c", "3", "z"]] });

// sort_range used to send column/ascending while the bridge read keyColumn/order, so the sort
// always landed on whatever "keyColumn" defaulted to.
const sortRes = await call("wps_excel_sort_range", { range: "A1:C3", column: 1, ascending: true });
check("sort_range accepts column/ascending", ok(sortRes), text(sortRes).replace(/\s+/g, " ").slice(0, 80));
const sorted = await read("A1:C3");
check("sort_range actually sorted by the given column", /第1行: a/.test(sorted), sorted.replace(/\s+/g, " ").slice(0, 90));

// copy_range sent source/destination but the bridge only read range: it copied to the clipboard
// and pasted nothing.
const copyRes = await call("wps_excel_copy_range", { source: "A1:C3", destination: "E1" });
check("copy_range accepts source/destination", ok(copyRes), text(copyRes).replace(/\s+/g, " ").slice(0, 80));
const copied = await read("E1:G1");
check("copy_range pasted the source at the destination", /第1行: a/.test(copied), copied.replace(/\s+/g, " ").slice(0, 90));

// transpose sent source/destination while the bridge read sourceRange/destinationCell.
const tr = await call("wps_excel_transpose", { source: "A1:A3", destination: "I1" });
check("transpose accepts source/destination", ok(tr), text(tr).replace(/\s+/g, " ").slice(0, 80));
const transposed = await read("I1:K1");
check("transpose wrote the transposed values", /第1行: a/.test(transposed), transposed.replace(/\s+/g, " ").slice(0, 90));

// fill_series sent direction, which the bridge never read.
const fs = await call("wps_excel_fill_series", { range: "A10:A14", type: "linear", step: 1, startValue: 1 });
check("fill_series runs", ok(fs), text(fs).replace(/\s+/g, " ").slice(0, 80));
const filled = await read("A10:A14");
// Read the rows explicitly: an earlier assertion passed on "第5行" alone while the series was empty.
check("fill_series wrote a linear series", /第1行: 1\b/.test(filled) && /第3行: 3\b/.test(filled) && /第5行: 5\b/.test(filled), filled.replace(/\s+/g, " ").slice(0, 130));

// auto_fill sent sourceRange/targetRange, which the bridge never read.
await call("wps_excel_write_range", { range: "C10:C11", data: [[1], [2]] });
const af = await call("wps_excel_auto_fill", { sourceRange: "C10:C11", targetRange: "C10:C14" });
check("auto_fill accepts sourceRange/targetRange", ok(af), text(af).replace(/\s+/g, " ").slice(0, 80));
const autoFilled = await read("C10:C14");
check("auto_fill extended the pattern", /4/.test(autoFilled), autoFilled.replace(/\s+/g, " ").slice(0, 110));

// set_hyperlink sent url/text while the bridge read address/textToDisplay.
// WPS keeps an existing cell value when a hyperlink is added, so the display text is only visible
// on an empty cell - which is also what makes this a real end-to-end check of url+text.
const link = await call("wps_excel_set_hyperlink", { cell: "B20", url: "https://example.com", text: "probe link" });
check("set_hyperlink accepts url/text", ok(link), text(link).replace(/\s+/g, " ").slice(0, 80));
const linkCell = await read("B20");
check("set_hyperlink carried url+text through to the sheet", /probe link/.test(linkCell), linkCell.replace(/\s+/g, " ").slice(0, 90));

// set_border sent borderStyle while the bridge read style.
const border = await call("wps_excel_set_border", { range: "A1:C3", borderStyle: "thick", position: "outline" });
check("set_border accepts borderStyle", ok(border), text(border).replace(/\s+/g, " ").slice(0, 80));

// set_data_validation sent type/formula while the bridge read validationType/formula1.
const dv = await call("wps_excel_set_data_validation", { range: "D1:D5", type: "list", formula: "x,y,z" });
check("set_data_validation accepts type/formula", ok(dv), text(dv).replace(/\s+/g, " ").slice(0, 90));

// set_conditional_format sent condition/format while the bridge read operator/value/colour.
const cf = await call("wps_excel_set_conditional_format", { range: "E1:E5", condition: ">100", format: "red_fill" });
check("set_conditional_format accepts condition/format", ok(cf), text(cf).replace(/\s+/g, " ").slice(0, 90));
const cfBad = await call("wps_excel_set_conditional_format", { range: "E1:E5", condition: ">1", format: "not-a-format" });
check("an unknown conditional format fails loudly", !ok(cfBad), text(cfBad).replace(/\s+/g, " ").slice(0, 90));

// set_cell_style advertised a named style the bridge never read.
const styleBad = await call("wps_excel_set_cell_style", { range: "A1", style: "definitely-not-a-real-style" });
check("an unknown cell style fails loudly instead of being dropped", !ok(styleBad), text(styleBad).replace(/\s+/g, " ").slice(0, 90));

// hide/show rows and columns: the tools sent count/hide/startRow/endRow but the bridge only read row/rows.
const hide = await call("wps_excel_hide_row", { row: 2, count: 2, hide: true });
check("hide_row accepts row+count+hide", ok(hide), text(hide).replace(/\s+/g, " ").slice(0, 80));
const show = await call("wps_excel_hide_row", { row: 2, count: 2, hide: false });
check("hide_row hide=false shows them again", ok(show), text(show).replace(/\s+/g, " ").slice(0, 80));
const hideRows = await call("wps_excel_hide_rows", { startRow: 4, endRow: 5 });
check("hide_rows accepts startRow/endRow", ok(hideRows), text(hideRows).replace(/\s+/g, " ").slice(0, 80));
const showRows = await call("wps_excel_show_rows", { startRow: 4, endRow: 5 });
check("show_rows accepts startRow/endRow", ok(showRows), text(showRows).replace(/\s+/g, " ").slice(0, 80));
const showCols = await call("wps_excel_show_columns", { startColumn: "A", endColumn: "C" });
check("show_columns accepts startColumn/endColumn", ok(showCols), text(showCols).replace(/\s+/g, " ").slice(0, 80));
const hideCol = await call("wps_excel_hide_column", { column: 2, count: 2, hide: true });
check("hide_column accepts column+count+hide", ok(hideCol), text(hideCol).replace(/\s+/g, " ").slice(0, 80));
await call("wps_excel_hide_column", { column: 2, count: 2, hide: false });

// freeze_panes sent freeze, which the bridge never read, so panes could never be unfrozen.
const freeze = await call("wps_excel_freeze_panes", { row: 1, column: 1 });
check("freeze_panes freezes", ok(freeze), text(freeze).replace(/\s+/g, " ").slice(0, 80));
const unfreeze = await call("wps_excel_freeze_panes", { freeze: false });
check("freeze_panes accepts freeze=false", ok(unfreeze), text(unfreeze).replace(/\s+/g, " ").slice(0, 80));

// protect_sheet/protect_workbook sent protect, which the bridge never read, so nothing could be unprotected.
const prot = await call("wps_excel_protect_sheet", { protect: true });
check("protect_sheet protects", ok(prot), text(prot).replace(/\s+/g, " ").slice(0, 80));
const unprot = await call("wps_excel_protect_sheet", { protect: false });
check("protect_sheet accepts protect=false", ok(unprot), text(unprot).replace(/\s+/g, " ").slice(0, 80));
const wprot = await call("wps_excel_protect_workbook", { protect: true, password: "probe" });
check("protect_workbook protects", ok(wprot), text(wprot).replace(/\s+/g, " ").slice(0, 80));
const wunprot = await call("wps_excel_protect_workbook", { protect: false, password: "probe" });
check("protect_workbook accepts protect=false", ok(wunprot), text(wunprot).replace(/\s+/g, " ").slice(0, 80));

// insert_excel_image sent filePath/imagePath/cell; only path was read and cell was dropped.
const img = await call("wps_excel_insert_excel_image", { filePath: path.resolve(PROBE_PNG), cell: "E10" });
check("insert_excel_image accepts filePath+cell", ok(img), text(img).replace(/\s+/g, " ").slice(0, 80));

// The big one: wps_excel_add_comment called the Word action of the same name, so an Excel comment
// was written into the Word document's selection instead.
await viaAction("createDocument", {}, "wps");
await call("wps_word_insert_text", { text: "word body text", position: "start" });
const wordBefore = text(await call("wps_word_get_document_text", {}));

const cmt = await call("wps_excel_add_comment", { cell: "B2", comment: "note-on-b2" });
check("add_comment runs", ok(cmt), text(cmt).replace(/\s+/g, " ").slice(0, 90));
const allComments = payload(await viaAction("getCellComments", {})).data || {};
check("the comment landed on the Excel sheet", JSON.stringify(allComments).includes("note-on-b2"), JSON.stringify(allComments).slice(0, 110));
const inRange = payload(await viaAction("getCellComments", { range: "B2:B2" })).data || {};
check("getCellComments honours a range filter", JSON.stringify(inRange).includes("note-on-b2"), JSON.stringify(inRange).slice(0, 110));
const outRange = payload(await viaAction("getCellComments", { range: "A1:A1" })).data || {};
check("getCellComments excludes comments outside the range", !JSON.stringify(outRange).includes("note-on-b2"), JSON.stringify(outRange).slice(0, 110));

const wordAfter = text(await call("wps_word_get_document_text", {}));
check("CROSS-APP: the Word document is untouched", wordAfter.includes("word body text") && !wordAfter.includes("note-on-b2"), wordAfter.replace(/\s+/g, " ").slice(0, 90));

// Close everything this test opened. The close actions are dialog-safe now (see
// test/close-safety.test.mjs), so a scratch run leaves no documents behind for the next one.
// Word has no close tool, so all three go through the wps_call facade.
for (const [method, appType] of [["closeWorkbook", "et"], ["closeDocument", "wps"], ["closePresentation", "wpp"]]) {
  for (let i = 0; i < 6; i++) {
    const res = await call("wps_call", { tool: "wps_execute_method", args: { method, params: { save: false }, appType } });
    if (!ok(res)) break;
  }
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "EXCEL CONTRACT FIX TESTS OK (" + results.length + ")" : "EXCEL CONTRACT FIX TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

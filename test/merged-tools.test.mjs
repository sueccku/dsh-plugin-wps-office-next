import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import { existsSync, writeFileSync } from "node:fs";
import path from "node:path";

const require = createRequire(import.meta.url);
const { DEPRECATED_TOOLS } = require(path.resolve("mcp/dist/tools/deprecated.js"));

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
await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "merged", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params, appType) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {}, appType } });
const read = async (range) => text(await call("wps_excel_read_range", { range }));
const slideCount = async () => payload(await viaAction("getSlideCount", {})).data?.count ?? -1;

check("merge table has 18 entries", Object.keys(DEPRECATED_TOOLS).length === 18, Object.keys(DEPRECATED_TOOLS).length + " entries");

// ---------- Excel ----------
await viaAction("createWorkbook", {});
await call("wps_excel_write_range", { range: "A1:B4", data: [["a", 1], ["b", 2], ["c", 3], ["d", 4]] });
check("insert_row forwards to insert_rows", ok(await call("wps_excel_insert_row", { row: 2, count: 1 })), "");
const afterInsert = await read("A1:A5");
check("insert_row really inserted a row", /绗?琛? a/.test(afterInsert), afterInsert.replace(/\s+/g, " ").slice(0, 90));
check("delete_row forwards to delete_rows", ok(await call("wps_excel_delete_row", { row: 2, count: 1 })), "");
const afterDelete = await read("A1:A4");
check("delete_row really removed the row", /绗?琛? b/.test(afterDelete), afterDelete.replace(/\s+/g, " ").slice(0, 90));
check("insert_column forwards to insert_columns", ok(await call("wps_excel_insert_column", { column: "B", count: 1 })), "");
check("delete_column forwards to delete_columns", ok(await call("wps_excel_delete_column", { column: "B", count: 1 })), "");
check("hide_row hides", ok(await call("wps_excel_hide_row", { row: 2, count: 2, hide: true })), "");
check("hide_row hide=false shows again", ok(await call("wps_excel_hide_row", { row: 2, count: 2, hide: false })), "");
check("canonical hide_rows still works", ok(await call("wps_excel_hide_rows", { startRow: 3, endRow: 4 })), "");
check("canonical hide_rows hide=false works", ok(await call("wps_excel_hide_rows", { startRow: 3, endRow: 4, hide: false })), "");
await call("wps_excel_write_range", { range: "D1:D2", data: [[1], [2]] });
check("auto_fill forwards to fill_series", ok(await call("wps_excel_auto_fill", { sourceRange: "D1:D2", targetRange: "D1:D5" })), "");
const filled = await read("D1:D5");
check("auto_fill really extended the pattern", /绗?琛? 3/.test(filled), filled.replace(/\s+/g, " ").slice(0, 110));
check("canonical fill_series still works", ok(await call("wps_excel_fill_series", { range: "F1:F4", type: "linear", step: 1, startValue: 10 })), "");
const series = await read("F1:F4");
check("canonical fill_series still writes a series", /绗?琛? 10/.test(series) && /绗?琛? 13/.test(series), series.replace(/\s+/g, " ").slice(0, 110));

// ---------- Presentation ----------
await call("wps_ppt_create_presentation", {});
for (let i = 0; i < 2; i++) await call("wps_ppt_add_slide", { layout: "blank" });
await call("wps_ppt_add_textbox", { slideIndex: 1, text: "t1", left: 50, top: 40, width: 200, height: 40 });
await call("wps_ppt_add_textbox", { slideIndex: 1, text: "t2", left: 300, top: 40, width: 200, height: 40 });
check("insert_image forwards path to filePath", ok(await call("wps_ppt_insert_image", { slideIndex: 2, path: path.resolve(PROBE_PNG), left: 40, top: 40 })), "");
check("insert_slide_image forwards imagePath to filePath", ok(await call("wps_ppt_insert_slide_image", { slideIndex: 2, imagePath: path.resolve(PROBE_PNG), left: 60, top: 60 })), "");
check("canonical insert_ppt_image still works", ok(await call("wps_ppt_insert_ppt_image", { slideIndex: 2, filePath: path.resolve(PROBE_PNG), left: 80, top: 80 })), "");
const shapes = payload(await viaAction("getShapes", { slideIndex: 2 })).data?.shapes ?? [];
const pictures = shapes.filter((s) => s.type === 13).length;
check("three pictures landed on slide 2", pictures >= 3, "pictures=" + pictures);
check("set_animation forwards animationType to effect", ok(await call("wps_ppt_set_animation", { slideIndex: 1, shapeIndex: 1, animationType: "fadeIn" })), "");
check("set_transition forwards transition to effect", ok(await call("wps_ppt_set_transition", { slideIndex: 1, transition: "fade" })), "");
check("set_background forwards flat color", ok(await call("wps_ppt_set_background", { slideIndex: 1, color: "#FF0000" })), "");
check("set_background forwards flat imagePath", ok(await call("wps_ppt_set_background", { slideIndex: 1, imagePath: path.resolve(PROBE_PNG) })), "");
check("canonical set_slide_background still takes the object", ok(await call("wps_ppt_set_slide_background", { slideIndex: 1, background: { type: "solid", color: "#00FF00" } })), "");
check("add_chart forwards to insert_ppt_chart", ok(await call("wps_ppt_add_chart", { slideIndex: 2, chartType: "column_clustered", title: "probe" })), "");
check("align_objects aligns without an index list", ok(await call("wps_ppt_align_objects", { slideIndex: 1, alignment: "left" })), "");
check("canonical align_shapes still takes indices", ok(await call("wps_ppt_align_shapes", { slideIndex: 1, shapeIndices: [1, 2], alignment: "left" })), "");
const before = await slideCount();
check("duplicate_slide forwards to copy_slide", ok(await call("wps_ppt_duplicate_slide", { slideIndex: 1 })), "");
const afterCount = await slideCount();
check("duplicate_slide added a slide", afterCount === before + 1, "before=" + before + " after=" + afterCount);

// ---------- Word ----------
await viaAction("createDocument", {}, "wps");
await call("wps_word_insert_text", { text: "merge probe", position: "start" });
check("set_font_style forwards fontName/fontSize", ok(await call("wps_word_set_font_style", { fontName: "Arial", fontSize: 12, bold: true })), "");

// ---------- visibility ----------
const help = payload(await call("wps_help", {}));
const hiddenProbe = ["wps_excel_hide_row", "wps_ppt_add_chart", "wps_word_set_font_style"];
const leaked = hiddenProbe.filter((n) => JSON.stringify(help).includes(n));
check("merged names are hidden from wps_help", leaked.length === 0, "leaked=" + leaked.join(",") + " total=" + help.total);

// ---------- teardown ----------
for (const pair of [["closeWorkbook", "et"], ["closeDocument", "wps"], ["closePresentation", "wpp"]]) {
  for (let i = 0; i < 6; i++) {
    const res = await call("wps_call", { tool: "wps_execute_method", args: { method: pair[0], params: { save: false }, appType: pair[1] } });
    if (!ok(res)) break;
  }
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "MERGED TOOLS TESTS OK (" + results.length + ")" : "MERGED TOOLS TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

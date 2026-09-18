// S4 coverage: drive the previously-untested Word and common tools against a real scratch document.
// Same contract: every entry must RETURN (never hang); "ok" must succeed, "any" may be a business error.
// Run: node test/word-common-coverage.test.mjs
import { spawn } from "node:child_process";
import { writeFileSync } from "node:fs";
import { resolve as resolvePath } from "node:path";

const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const imgPath = resolvePath("test/.artifacts/wordcov.png");
const docPath = resolvePath("test/.artifacts/wordcov.docx");
const pdfPath = resolvePath("test/.artifacts/wordcov-convert.pdf");
writeFileSync(imgPath, Buffer.from(PNG, "base64"));

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

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "word-cov", version: "1" } });
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

// ---- scratch document --------------------------------------------------------------------
check("create_document", ok(await call("wps_word_create_document", {})), "");
await call("wps_word_insert_text", { text: "第一段测试文本。\n第二段测试文本。", position: "start" });
await call("wps_word_insert_bookmark", { name: "BM1" });
const saved = await call("wps_common_save_as", { filePath: docPath });

const MATRIX = [
  // Word
  ["wps_word_get_paragraphs", {}, "any"],
  ["wps_word_get_track_changes_status", {}, "ok"],
  ["wps_word_find_in_document", { find_text: "测试" }, "any"],
  ["wps_word_apply_style", { style_name: "Heading 1", range: "all" }, "any"],
  ["wps_word_set_font", { font_name: "微软雅黑", font_size: 12, range: "all" }, "any"],
  ["wps_word_set_text_color", { color: "#FF0000" }, "any"],
  ["wps_word_set_line_spacing", { lineSpacing: 1.5 }, "ok"],
  ["wps_word_set_paragraph", { alignment: "center" }, "any"],
  ["wps_word_set_page_setup", { orientation: "landscape" }, "any"],
  ["wps_word_insert_page_break", {}, "any"],
  ["wps_word_insert_section_break", { breakType: "nextPage" }, "any"],
  ["wps_word_insert_image", { imagePath: imgPath }, "any"],
  ["wps_word_generate_toc", {}, "any"],
  ["wps_word_replace_bookmark_content", { name: "BM1", text: "书签内容" }, "any"],
  ["wps_word_replace_range", { start_pos: 0, end_pos: 2, text: "替换" }, "any"],
  ["wps_word_smart_fill_field", { keyword: "甲方", value: "某公司" }, "any"],
  ["wps_word_switch_document", { name: "wordcov.docx" }, "any"],
  ["wps_word_proofread_basic", { text: "这是一段需要校对的中文文本。" }, "any"],
  // common
  ["wps_common_get_app_info", {}, "any"],
  ["wps_common_get_selected_text", {}, "any"],
  ["wps_common_set_selected_text", { text: "选中替换" }, "any"],
  ["wps_common_wire_check", {}, "any"],
  ["wps_convert_format", { targetFormat: "pdf", outputPath: pdfPath, app_type: "wps" }, "any"],
];

let succeeded = 0;
for (const [name, args, expect] of MATRIX) {
  const started = Date.now();
  const res = await call(name, args, name.includes("proofread") ? 90000 : 30000);
  const ms = Date.now() - started;
  const hung = !!(res && res.__timeout);
  const good = ok(res);
  if (good) succeeded++;
  const label = name.replace("wps_", "");
  if (hung) check(label + " returns (no hang)", false, "HUNG after " + ms + "ms");
  else if (expect === "ok") check(label + " succeeds", good, ms + "ms " + text(res).replace(/\s+/g, " ").slice(0, 70));
  else check(label + " returns (no hang)", true, ms + "ms " + (good ? "ok" : "business error"));
}

// save only once the document has a real path, otherwise a Save As dialog would wedge WPS.
if (ok(saved)) check("common_save (document already has a path)", ok(await call("wps_common_save", {})), "");
else check("common_save skipped (save_as failed, avoiding a modal Save As)", true, "save_as did not succeed");

for (let i = 0; i < 6; i++) {
  const res = await call("wps_call", { tool: "wps_execute_method", args: { method: "closeDocument", params: { save: false }, appType: "wps" } });
  if (!ok(res)) break;
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log("      matrix tools succeeded: " + succeeded + " / " + MATRIX.length);
check("the matrix actually drove tools (no systemic failure)", succeeded >= 5, succeeded + " succeeded");
console.log(failed === 0 ? "WORD/COMMON COVERAGE TESTS OK (" + results.length + ")" : "WORD/COMMON COVERAGE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

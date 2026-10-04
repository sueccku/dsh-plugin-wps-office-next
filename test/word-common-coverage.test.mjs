// S4 coverage: drive the previously-untested Word and common tools against a real scratch document.
// Same contract: every entry must RETURN (never hang); "ok" must succeed, "any" may be a business error.
// Run: node test/word-common-coverage.test.mjs
import { spawn, spawnSync } from "node:child_process";
import { existsSync, writeFileSync } from "node:fs";
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
// 裸 COM 读真值（只读）。输出走 base64：中文经 PowerShell → pipe 的编码链会被改坏。
function com(script) {
  const b64 = Buffer.from(script, "utf8").toString("base64");
  const r = spawnSync("powershell", ["-NoProfile", "-STA", "-Command",
    "$s = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + b64 + "')); $out = Invoke-Expression $s; [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$out))"],
    { encoding: "utf8", windowsHide: true });
  try { return Buffer.from((r.stdout || "").trim(), "base64").toString("utf8"); } catch { return ""; }
}
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
const saved = await call("wps_common_save_as", { outputPath: docPath });

const MATRIX = [
  // Word
  ["wps_word_get_paragraphs", {}, "any"],
  ["wps_word_get_track_changes_status", {}, "ok"],
  ["wps_word_find_in_document", { findText: "测试" }, "ok"],
  // range 现在只接受 {start,end}（0 基、end 不含）或 "all"。这里用显式坐标而不是 "all"：
  // 顺带验证"按字符范围施加格式"这条路径真的работает（FIXES 86 之前它根本不存在）。
  // 样式名必须用中文 WPS 里真实存在的：实测 "Heading 1" 会 E_FAIL（模板里没有这个英文名），
  // 而工具描述里恰恰把 "Heading 1" 当作示例 —— 这一条按真实可用的名字写。
  ["wps_word_apply_style", { styleName: "标题 1", range: { start: 0, end: 4 } }, "ok"],
  ["wps_word_set_font", { fontName: "微软雅黑", fontSize: 12, range: { start: 0, end: 4 } }, "ok"],
  ["wps_word_set_text_color", { color: "#FF0000" }, "ok"],
  ["wps_word_set_line_spacing", { lineSpacing: 1.5 }, "ok"],
  ["wps_word_set_paragraph", { alignment: "center" }, "ok"],
  ["wps_word_set_page_setup", { orientation: "landscape" }, "ok"],
  ["wps_word_insert_page_break", {}, "ok"],
  ["wps_word_insert_section_break", { breakType: "nextPage" }, "ok"],
  ["wps_word_insert_image", { imagePath: imgPath }, "ok"],
  ["wps_word_generate_toc", {}, "ok"],
  ["wps_word_replace_bookmark_content", { name: "BM1", text: "书签内容" }, "ok"],
  // 破坏性动作：单段内的替换必须真的成功（FIXES 86 起不校验边界的那版会把越界范围静默钳制到文末）。
  ["wps_word_replace_range", { startPos: 0, endPos: 2, text: "替换" }, "ok"],
  // 关键字"甲方"不在本文档里；工具明确报 not found（有专门断言覆盖这条诚实行为）。
  ["wps_word_smart_fill_field", { keyword: "甲方", value: "某公司" }, "any"],
  ["wps_word_switch_document", { name: "wordcov.docx" }, "ok"],
  ["wps_word_proofread_basic", { text: "这是一段需要校对的中文文本。" }, "any"],
  // common
  ["wps_common_get_app_info", {}, "ok"],
  ["wps_common_get_selected_text", {}, "any"],
  ["wps_common_set_selected_text", { text: "选中替换" }, "any"],
  ["wps_common_wire_check", {}, "ok"],
  ["wps_convert_format", { targetFormat: "pdf", outputPath: pdfPath, appType: "wps" }, "any"],
];

// ---- L1（FIXES 89）：11 个只被 "any" 放行的 Word 工具，逐个断言真实结果 --------------------------
// 每个都问一句「做完了到底成没成」，而不是「调用返回了」。真值用裸 COM 读。

const found = text(await call("wps_word_find_in_document", { findText: "测试" }));
check("find_in_document reports a match count", /出现了\s*\d+\s*次|共\s*\d+/.test(found), found.replace(/\s+/g, " ").slice(0, 90));

const colored = await call("wps_word_set_text_color", { color: "#FF0000" });
check("set_text_color succeeds", ok(colored), text(colored).replace(/\s+/g, " ").slice(0, 80));

const aligned = await call("wps_word_set_paragraph", { alignment: "center" });
check("set_paragraph succeeds", ok(aligned), text(aligned).replace(/\s+/g, " ").slice(0, 80));
const alignRead = com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); [string]$w.ActiveDocument.Paragraphs.Item(1).Alignment").trim();
check("set_paragraph really changed the alignment", alignRead === "1", "Alignment=" + alignRead);

const pageSetup = await call("wps_word_set_page_setup", { orientation: "landscape" });
check("set_page_setup succeeds", ok(pageSetup), text(pageSetup).replace(/\s+/g, " ").slice(0, 80));
const orientRead = com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); [string]$w.ActiveDocument.PageSetup.Orientation").trim();
check("set_page_setup actually oriented the page", orientRead === "1", "Orientation=" + orientRead);

const sectionsBefore = Number(com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); [string]$w.ActiveDocument.Sections.Count").trim());
const sectionBreak = await call("wps_word_insert_section_break", { breakType: "nextPage" });
check("insert_section_break succeeds", ok(sectionBreak), text(sectionBreak).replace(/\s+/g, " ").slice(0, 80));
const sectionsAfter = Number(com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); [string]$w.ActiveDocument.Sections.Count").trim());
check("insert_section_break really added a section", sectionsAfter > sectionsBefore, sectionsBefore + " -> " + sectionsAfter);

const shapesBefore = Number(com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); [string]$w.ActiveDocument.InlineShapes.Count").trim());
const img = await call("wps_word_insert_image", { imagePath: imgPath });
check("insert_image succeeds", ok(img), text(img).replace(/\s+/g, " ").slice(0, 80));
const shapesAfter = Number(com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); [string]$w.ActiveDocument.InlineShapes.Count").trim());
check("insert_image really added an inline shape", shapesAfter > shapesBefore, shapesBefore + " -> " + shapesAfter);

// 目录要求文档里有标题样式；直接调 COM 给第一段套上「标题 1」，再生成。
const styled = com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); $d = $w.ActiveDocument; $d.Paragraphs.Item(1).Range.Style = '标题 1'; [string]$d.Paragraphs.Item(1).Range.Style.NameLocal");
check("scratch doc has a heading paragraph", styled.includes("标题 1"), styled.slice(0, 30));
const toc = await call("wps_word_generate_toc", {});
check("generate_toc succeeds", ok(toc), text(toc).replace(/\s+/g, " ").slice(0, 80));
const tocCount = Number(com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); [string]$w.ActiveDocument.TablesOfContents.Count").trim());
check("generate_toc really created a table of contents", tocCount >= 1, "TablesOfContents=" + tocCount);

const bm = await call("wps_word_replace_bookmark_content", { name: "BM1", text: "书签内容" });
check("replace_bookmark_content succeeds", ok(bm), text(bm).replace(/\s+/g, " ").slice(0, 80));
const bmText = com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); [string]$w.ActiveDocument.Bookmarks.Item('BM1').Range.Text");
check("the bookmark now holds the new text", bmText.includes("书签内容"), bmText.slice(0, 40));

const fillBad = await call("wps_word_smart_fill_field", { keyword: "绝不存在的字段名XYZ", value: "v" });
check("smart_fill_field fails clearly for an unknown keyword", !ok(fillBad) || /未找到|没有找到|not found/i.test(text(fillBad)), text(fillBad).replace(/\s+/g, " ").slice(0, 90));

const swDoc = await call("wps_word_switch_document", { name: "wordcov.docx" });
check("switch_document succeeds for the saved scratch doc", ok(swDoc), text(swDoc).replace(/\s+/g, " ").slice(0, 80));
const swBadDoc = await call("wps_word_switch_document", { name: "definitely-not-open.docx" });
check("switch_document fails clearly for an unknown name", !ok(swBadDoc), text(swBadDoc).replace(/\s+/g, " ").slice(0, 90));

const proof = text(await call("wps_word_proofread_basic", { text: "这是一段需要校对的中文文本。" }));
check("proofread_basic returns a readable verdict", proof.length > 10 && !/undefined/.test(proof), proof.replace(/\s+/g, " ").slice(0, 90));
// ---- L1（FIXES 89）：通用工具的弱断言，也补上真实结果检查 -------------------------------------
const appInfo = text(await call("wps_common_get_app_info", {}));
check("common_get_app_info reports a real app or version", appInfo.length > 5 && !/undefined/.test(appInfo), appInfo.replace(/\s+/g, " ").slice(0, 90));
const wire = text(await call("wps_common_wire_check", {}));
check("common_wire_check reports the wiring state", wire.length > 5 && !/undefined/.test(wire), wire.replace(/\s+/g, " ").slice(0, 90));
// 选中文本读取：没有选中内容时也必须给可读结果（不能是 undefined）
const selText = text(await call("wps_common_get_selected_text", {}));
check("common_get_selected_text is readable even with no selection", selText.length > 0 && !/undefined/.test(selText), selText.replace(/\s+/g, " ").slice(0, 90));
const setSel = await call("wps_common_set_selected_text", { text: "选中替换" });
// 以前这里是 `ok(x) || text(x).length > 0` —— 裸报错也算过，等于没断言（FIXES 91 发版审计抓到）。
// 要么真的写进去（回读能查到），要么明确说清为什么不行。
const selRoundTrip = text(await call("wps_common_get_selected_text", {}));
check(
  "common_set_selected_text either writes the text or says why not",
  ok(setSel) ? selRoundTrip.includes("选中替换") : /失败|错误|没有|无法|not found/i.test(text(setSel)),
  "claimed=" + ok(setSel) + " roundTrip=" + selRoundTrip.replace(/\s+/g, " ").slice(0, 60)
);
const conv = await call("wps_convert_format", { targetFormat: "pdf", outputPath: pdfPath, appType: "wps" });
// 同上：成功就必须真的落盘，失败就必须有可读原因。
check(
  "convert_format either writes the file or says why not",
  ok(conv) ? existsSync(pdfPath) : /失败|错误|无法|不支持|not |missing/i.test(text(conv)),
  "claimed=" + ok(conv) + " pdfExists=" + existsSync(pdfPath) + " | " + text(conv).replace(/\s+/g, " ").slice(0, 70)
);
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
  else {
    const body = String((res && res.result && res.result.content && res.result.content[0].text) || "");
    // "any" 只放行「没挂住」太弱了：至少要求返回了内容或一句可读的错误（FIXES 74）。
    check(label + " returns something readable (no hang)", body.length > 0, ms + "ms " + (good ? "ok" : "business error") + (body ? "" : " EMPTY RESPONSE"));
  }
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

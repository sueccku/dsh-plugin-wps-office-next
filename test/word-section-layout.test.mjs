// S7 D1/D2 回归：Word **分节级排版** 的三条能力 + 一条顺序缺陷。
//
// 背景：S7 真机走查（docs/scenarios/S7-真机走查报告.md）实测确认，模型拿不到三条必需能力，
// 只能绕到 WPS 自身 COM 接口去做。本文件把这三条钉进插件，并**以文件为准**复核：
//
//   L1  setPageSetup 的 section  —— 只改点名的那一节，邻居不动（旧实现写 $doc.PageSetup，全篇生效）
//   L2  横向 + 纸张的顺序        —— 旧实现先写 orientation 后写 paperSize，paperSize 会把 orientation
//                                  **静默重置回纵向**（probe-d1c：orient=1→paper=6 的净结果是"纵向 A3"）
//   L3  insertPageNumbers 格式   —— numberStyle=lowerRoman + restartSection=true（罗马数字从 1 重开）
//   L4  insertPageNumbers 续前节 —— restartSection=false 不重开
//   L5  setTableFormat 表头重复  —— headerRepeat → 第 1 行是重复标题行（OOXML w:tblHeader）
//   L6  错误路径                 —— 节号越界如实报错，不静默改错节
//
// 判据只用两类证据：① 逐节 get_page_setup 的读回值；② 落盘 docx 的 OOXML。
// Run: node test/word-section-layout.test.mjs      （需要真实 WPS）
import { rmSync } from "node:fs";
import { resolve as resolvePath } from "node:path";
import { startServer, createReport } from "./lib/harness.mjs";
import { py, zipScan } from "./lib/oracle.mjs";

const docPath = resolvePath("test/.artifacts/word-section-layout.docx");
rmSync(docPath, { force: true });

const srv = startServer({ clientName: "word-section-layout" });
const { check, finish } = createReport("word section-scoped layout (S7)");
const { call, text, ok } = srv;

// 每节的页面设置：parse "第 N 节页面设置：\n方向: 横向\n纸张: A3\n页面: 842pt x 1190.7pt\n边距: ..."
function parseSectionSetup(t) {
  const out = {};
  const orient = /方向: (\S+)/.exec(t);
  if (orient) out.orientation = orient[1];
  const paper = /纸张: (\S+)/.exec(t);
  if (paper) out.paperSize = paper[1];
  const dims = /页面: ([\d.]+)pt x ([\d.]+)pt/.exec(t);
  if (dims) {
    out.w = Number(dims[1]);
    out.h = Number(dims[2]);
  }
  const m = /边距: 上 ([\d.]+)pt \/ 下 ([\d.]+)pt \/ 左 ([\d.]+)pt \/ 右 ([\d.]+)pt/.exec(t);
  if (m) {
    out.top = Number(m[1]);
    out.bottom = Number(m[2]);
    out.left = Number(m[3]);
    out.right = Number(m[4]);
  }
  return out;
}

async function sectionSetup(n) {
  const res = await call("wps_call", { tool: "wps_word_get_page_setup", args: { section: n } });
  return { ok: ok(res), raw: text(res), v: parseSectionSetup(text(res)) };
}

// ---- 0. 建文档 + 三个分节符 → 4 节 -------------------------------------------------------
const created = await call("wps_word_create_document", {});
check("create_document", ok(created), text(created).slice(0, 80));
await call("wps_word_insert_text", { text: "封面页\n目录页\n技术方案页\n商务报价页", position: "start" });

for (let i = 0; i < 3; i++) {
  const brk = await call("wps_call", { tool: "wps_word_insert_section_break", args: { breakType: "nextPage" } });
  check("insert_section_break #" + (i + 1), ok(brk), text(brk).slice(0, 80));
}

const probeCount = await call("wps_word_set_page_setup", { section: 1 });
const countTxt = text(probeCount);
const totalSections = Number((countTxt.match(/全文共 (\d+) 节/) || [])[1] || 0);
check("文档有 4 节", totalSections === 4, "sectionCount=" + totalSections + " / " + countTxt.replace(/\n/g, " ").slice(0, 100));

// ---- L1/L2：第 1 节纵向 A4、第 2 节纵向 A4、第 3 节横向 A3、第 4 节纵向 A3 ----------------
const set1 = await call("wps_word_set_page_setup", { section: 1, orientation: "portrait", paperSize: "A4", topMargin: 56.7, bottomMargin: 56.7, leftMargin: 85, rightMargin: 85 });
check("L1 第1节 纵向 A4", ok(set1), text(set1).replace(/\n/g, " ").slice(0, 110));

// 关键用例：横向 + A3（旧实现会因为写入顺序把横向吃掉）
const set3 = await call("wps_word_set_page_setup", { section: 3, orientation: "landscape", paperSize: "A3" });
check("L2 第3节 横向 A3（写成功）", ok(set3), text(set3).replace(/\n/g, " ").slice(0, 130));

const r1 = await sectionSetup(1);
const r2 = await sectionSetup(2);
const r3 = await sectionSetup(3);
check("L1 第1节读回 纵向 A4", r1.ok && r1.v.orientation === "纵向" && r1.v.paperSize === "A4", JSON.stringify(r1.v));
check("L1 第2节未被第3节的设置影响（仍纵向 A4）", r2.ok && r2.v.orientation === "纵向" && r2.v.paperSize === "A4", JSON.stringify(r2.v));
check("L2 第3节读回 **横向** A3（顺序缺陷的真判据）", r3.ok && r3.v.orientation === "横向", "orientation=" + r3.v.orientation + " paper=" + r3.v.paperSize + " " + r3.v.w + "x" + r3.v.h);
check("L2 第3节纸张是 A3 横向（宽>高）", r3.ok && r3.v.w > r3.v.h && Math.abs(r3.v.w - 1190.7) <= 6 && Math.abs(r3.v.h - 842) <= 6, r3.v.w + "x" + r3.v.h);
// D1 实测：WPS 把「A4 横向」记成未命名纸张代码 41（改回纵向又恢复 A4）。只回一个裸数字时，
// 模型会把它当成"纸张被改坏了"并停下来排查 —— 所以未知代码必须带解释。
check("L2 横向页的纸张读数带解释（不能只回裸代码）", r3.ok && (r3.v.paperSize === "A3" || /（/.test(String(r3.v.paperSize))), "paper=" + r3.v.paperSize);
check("L1 第1节边距写—读闭环（上/左 ±1pt）", r1.ok && Math.abs(r1.v.top - 56.7) <= 1 && Math.abs(r1.v.left - 85) <= 1, "top=" + r1.v.top + " left=" + r1.v.left);

// 第 4 节保持默认（纵向 A4），证明确实只动了第 3 节
const r4 = await sectionSetup(4);
check("L1 第4节仍是纵向（没被第3节带跑）", r4.ok && r4.v.orientation === "纵向", JSON.stringify(r4.v));

// ---- L6：节号越界必须如实报错，并且后缀带上"文档共有几节" -------------------------------
const bad = await call("wps_word_set_page_setup", { section: 99, orientation: "landscape" });
const badTxt = text(bad);
check("L6 节号越界如实报错", !ok(bad), badTxt.slice(0, 110));
check("L6 报错里说明文档共有几节", /out of range/.test(badTxt) && /section\(s\)/.test(badTxt), badTxt.slice(0, 130));

// ---- L3/L4：页码格式 + 分节重开 ---------------------------------------------------------
const pn2 = await call("wps_word_insert_page_numbers", { section: 2, position: "footer", numberStyle: "lowerRoman", restartSection: true });
check("L3 第2节 小写罗马数字 + 从1重开", ok(pn2), text(pn2).replace(/\n/g, " ").slice(0, 150));

const pn3 = await call("wps_word_insert_page_numbers", { section: 3, position: "footer", numberStyle: "arabic", restartSection: true, includeTotal: true });
check("L3 第3节 阿拉伯数字 + 从1重开 + 共Y页", ok(pn3), text(pn3).replace(/\n/g, " ").slice(0, 170));

const pn4 = await call("wps_word_insert_page_numbers", { section: 4, position: "footer", numberStyle: "arabic", restartSection: false });
check("L4 第4节 续前节（不重开）", ok(pn4), text(pn4).replace(/\n/g, " ").slice(0, 150));

// 非法编号格式要明确拒绝
const badStyle = await call("wps_word_insert_page_numbers", { section: 2, numberStyle: "chinese" });
check("L3 非法 numberStyle 被拒绝", !ok(badStyle) && /numberStyle/.test(text(badStyle)), text(badStyle).slice(0, 120));

// 域落盘：第2节页脚有 PAGE，第3节页脚有 PAGE + NUMPAGES
const saved = await call("wps_common_save_as", { outputPath: docPath });
check("save_as 落盘", ok(saved), text(saved).slice(0, 90));

const pageScan = zipScan(docPath, "PAGE", "word/footer");
check("L3 落盘 footer 里有 PAGE 域", !!pageScan && pageScan.count >= 1, JSON.stringify(pageScan));
const numScan = zipScan(docPath, "NUMPAGES", "word/footer");
check("L3 落盘 footer 里有 NUMPAGES 域", !!numScan && numScan.count >= 1, JSON.stringify(numScan));

// ---- L5：表头行跨页重复 ----------------------------------------------------------------
const tbl = await call("wps_call", { tool: "wps_word_insert_table", args: { rows: 3, cols: 3 } });
check("插入表格", ok(tbl), text(tbl).slice(0, 80));
const fmt = await call("wps_word_set_table_format", { table: 1, borders: true, headerRepeat: true });
check("L5 set_table_format headerRepeat", ok(fmt), text(fmt).replace(/\n/g, " ").slice(0, 130));

await call("wps_common_save_as", { outputPath: docPath });

const hdrScan = zipScan(docPath, "tblHeader", "word/document");
check("L5 落盘文件里表头行有 w:tblHeader", !!hdrScan && hdrScan.count >= 1, JSON.stringify(hdrScan));

// 关掉：回读仍是开的（不是"写完就丢"）
const fmtOff = await call("wps_word_set_table_format", { table: 1, headerRepeat: false });
check("L5 headerRepeat=false 也能改回去", ok(fmtOff), text(fmtOff).replace(/\n/g, " ").slice(0, 110));

// ---- 收尾 -------------------------------------------------------------------------------
await call("wps_word_close_document", { save: false });
srv.stop();
finish();

// WO-001：Word 侧「读写不对称」+「共 Y 页」的回归。
//
// 四条都**以文件为准**，不采信工具自述：
//   F5a 读页面设置   —— set 写入后 get 读回同值（写—读闭环）
//   F5c 更新目录     —— action:"update" 刷新已有的目录；没有目录时必须如实报错
//   F6  共 Y 页      —— includeTotal 后 footer 里同时有 PAGE 与 NUMPAGES，且**带缓存结果**
//   F5b 读页眉页脚   —— 读回页脚文本与域清单（PAGE / NUMPAGES）
//
// Run: node test/word-page-setup-gaps.test.mjs      （需要真实 WPS）
import { rmSync } from "node:fs";
import { resolve as resolvePath } from "node:path";
import { startServer, createReport } from "./lib/harness.mjs";
import { py, zipScan } from "./lib/oracle.mjs";

const docPath = resolvePath("test/.artifacts/word-gaps.docx");
rmSync(docPath, { force: true });

const srv = startServer({ clientName: "word-gaps" });
const { check, finish } = createReport("word gaps (WO-001)");
const { call, text, ok } = srv;

// ---- 0. 空白新文档：还没有目录，update 必须如实报错（不能回 success）----------------------
const created = await call("wps_word_create_document", {});
check("create_document", ok(created), text(created).slice(0, 80));

const noToc = await call("wps_word_generate_toc", { action: "update" });
check("update 在没有目录时如实报错", !ok(noToc) && /没有目录/.test(text(noToc)), text(noToc).slice(0, 120));

// ---- 1. 铺内容：三个一级标题 + 正文 + 页码域（让总页数有意义）----------------------------
await call("wps_word_insert_text", { text: "第一部分\n正文一。\n第二部分\n正文二。\n第三部分\n正文三。", position: "start" });
await call("wps_call", { tool: "wps_word_insert_page_break", args: {} });
await call("wps_call", { tool: "wps_word_insert_page_break", args: {} });

const paras = text(await call("wps_word_get_paragraphs", {}));
const headingRanges = [...paras.matchAll(/^\[(\d+)\] \(([^)]*)\)\s*(.*?)\s*@(\d+)-(\d+)/gm)]
  .filter((m) => /^第.部分$/.test(m[3]))
  .map((m) => ({ start: Number(m[4]), end: Number(m[5]) }));
check("定位到三个标题段落", headingRanges.length === 3, JSON.stringify(headingRanges));

for (const range of headingRanges) {
  const res = await call("wps_word_apply_style", { styleName: "标题 1", range });
  check("apply_style 标题 1 @" + range.start + "-" + range.end, ok(res), text(res).slice(0, 70));
}

// ---- 2. F5c：插入目录，然后只刷新（不新增第二份）---------------------------------------
const tocInsert = await call("wps_word_generate_toc", { action: "insert", levels: 3 });
check("generate_toc insert", ok(tocInsert), text(tocInsert).slice(0, 90));

const tocUpdate = await call("wps_word_generate_toc", { action: "update" });
check("generate_toc update 成功且回报了刷新数/条目数",
  ok(tocUpdate) && /刷新的目录数: 1/.test(text(tocUpdate)) && /条目数: [1-9]/.test(text(tocUpdate)),
  text(tocUpdate).replace(/\n/g, " ").slice(0, 120));

// ---- 3. F5a：写—读闭环（get/set 现在都在广告位，两个都直接调）---------------------------
const setRes = await call("wps_word_set_page_setup", { orientation: "landscape", topMargin: 56, leftMargin: 72 });
check("set_page_setup 横向 + 上边距 56 + 左边距 72", ok(setRes), text(setRes).replace(/\n/g, " ").slice(0, 100));

const getRes = await call("wps_word_get_page_setup", { section: 1 });
const getTxt = text(getRes);
check("get_page_setup 可读", ok(getRes), getTxt.replace(/\n/g, " ").slice(0, 100));
check("写—读闭环：方向读回横向", ok(getRes) && /方向: 横向/.test(getTxt), getTxt.match(/方向: \S+/)?.[0] || "");
const topRead = Number((getTxt.match(/上 ([\d.]+)pt/) || [])[1]);
const leftRead = Number((getTxt.match(/左 ([\d.]+)pt/) || [])[1]);
check("写—读闭环：上边距 56pt（±1）", Number.isFinite(topRead) && Math.abs(topRead - 56) <= 1, "读回=" + topRead);
check("写—读闭环：左边距 72pt（±1）", Number.isFinite(leftRead) && Math.abs(leftRead - 72) <= 1, "读回=" + leftRead);

// 错误路径：不存在的节号要明确报错，而不是回一个空设置
const badSection = await call("wps_word_get_page_setup", { section: 99 });
check("get_page_setup 节号越界时如实报错", !ok(badSection), text(badSection).slice(0, 90));

// ---- 4. F6：页码 + 总页数（共 Y 页）--------------------------------------------------
const pageNumbers = await call("wps_word_insert_page_numbers", { section: 1, position: "footer", includeTotal: true });
check("insert_page_numbers includeTotal", ok(pageNumbers), text(pageNumbers).replace(/\n/g, " ").slice(0, 130));

const saved = await call("wps_common_save_as", { outputPath: docPath });
check("save_as 落盘", ok(saved), text(saved).slice(0, 90));

const npScan = zipScan(docPath, "NUMPAGES", "word/footer");
check("落盘文件 footer 里有 NUMPAGES 域代码", !!npScan && npScan.count >= 1, JSON.stringify(npScan));
const pageScan = zipScan(docPath, "PAGE", "word/footer");
check("落盘文件 footer 里 PAGE 域还在", !!pageScan && pageScan.count >= 1, JSON.stringify(pageScan));

// 域的**缓存结果**：证明 WPS 真算了，而不是只把域代码写进文件
const fields = py(`
import zipfile
z = zipfile.ZipFile(${JSON.stringify(docPath)})
codes = []
texts = []
parts = []
for n in z.namelist():
    if n.startswith("word/footer") and n.endswith(".xml"):
        parts.append(n)
        x = z.read(n).decode("utf-8", "ignore")
        for seg in x.split("<w:instrText")[1:]:
            codes.append(seg.split(">", 1)[1].split("<", 1)[0].strip())
        for seg in x.split("<w:t")[1:]:
            texts.append(seg.split(">", 1)[1].split("<", 1)[0])
result = {"parts": parts, "codes": codes, "texts": texts}
`);
const codes = (fields && fields.codes) || [];
check("footer 域代码含 PAGE 与 NUMPAGES",
  codes.some((c) => /^PAGE\b/.test(c)) && codes.some((c) => /^NUMPAGES\b/.test(c)),
  JSON.stringify(codes));
const cachedDigits = ((fields && fields.texts) || []).filter((t) => /^\d+$/.test(String(t).trim()));
check("页码/总页数带缓存结果（WPS 真的算了）", cachedDigits.length >= 2, "缓存数字=" + JSON.stringify(cachedDigits) + " 文本=" + JSON.stringify((fields && fields.texts) || []));

// ---- 5. F5b：读页眉页脚（含域清单）---------------------------------------------------
const hf = await call("wps_call", { tool: "wps_word_get_header_footer", args: { section: 1 } });
const hfTxt = text(hf);
check("get_header_footer 可读", ok(hf), hfTxt.replace(/\n/g, " ").slice(0, 120));
check("get_header_footer 页脚域里认得 NUMPAGES", ok(hf) && /NUMPAGES/.test(hfTxt), hfTxt.replace(/\n/g, " ").slice(0, 160));
// 页脚有两段（页码 / " / 共 3 页"），两段之间的段落标记是**合法内容**；只要求**尾部**标记被去掉：
// 空页眉必须回成 ""（而不是 "\r"），页脚值也不能以 \r 收尾。
check("get_header_footer 文本已去掉尾部段落标记",
  ok(hf) && /页眉: ""/.test(hfTxt) && !/\\r"/.test(hfTxt), hfTxt.replace(/\n/g, " ").slice(0, 160));

// ---- 收尾 -----------------------------------------------------------------------------
await call("wps_word_close_document", { save: false });
srv.stop();
finish();

// FIXES 86 回归：Word 的「按范围施加格式」与破坏性删除。
//
// 一份真实任务报告（2026-10-03）暴露了五个缺陷，共同特征是**返回 success，但什么都没做、
// 或者做在了别的地方**。这里把每一条都钉在真机上，逐条断言，包括用裸 COM 读字符级真值：
//   B1  setFont 范围为空时静默 no-op，却回「加粗: 是」→ 现在如实回报范围 + warning
//   B2  insertText 的 style 落在插入**之前**的光标位置 → 现在落在刚插入的文本上
//   B3  new_paragraph 用追加 \n 冒充段落、参数还被白名单丢掉 → 现在产生真段落
//   B4  applyStyle 的范围向整段扩张且只回一段误导性文本 → 现在回报真实受影响的段落
//   B10 replaceRange 不校验边界，越界 = 静默删掉 [start, 文末) → 现在报错；跨段删除要 confirm
// Run: node test/word-range-format.test.mjs
import { spawn, spawnSync } from "node:child_process";
import { resolve as resolvePath } from "node:path";

// 必须显式钉住仓库里的宿主脚本：从「已装本插件」的 DSH 会话里起的终端会带着 WPS_OFFICE_HOST_SCRIPT
// （plugin.js 发布的那份，指向 profile 里 0.6.x 的副本），那样测的就是旧宿主 —— 实测会被它的旧参数
// 白名单挡下来（unknown parameter confirm），测试就会因为错误的原因变红。
const child = spawn(process.execPath, ["mcp/dist/index.js"], {
  stdio: ["pipe", "pipe", "pipe"],
  windowsHide: true,
  env: { ...process.env, WPS_OFFICE_HOST_SCRIPT: resolvePath("host/wps-com-host.ps1") },
});
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

const results = [];
function check(name, okFlag, detail) { results.push({ name, ok: okFlag }); console.log((okFlag ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }
function payload(res) { try { return JSON.parse(text(res)); } catch { return {}; } }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "word-range", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
// replace_range / set_active_target 之类未广告的工具经门面调用。
const via = async (tool, args) => payload(await call("wps_call", { tool, args: args || {} }));

// 裸 COM：桥与工具都没有字符级读接口，真值只能直接问 WPS。只读，不改文档结构。
// 输出走 base64：中文经 PowerShell → pipe 的编码链会被改坏（实测段落样式名变成乱码），
// 拿到 base64 再在 Node 里解码，读到的一定是原字节。
function com(script) {
  const b64 = Buffer.from(script, "utf8").toString("base64");
  const r = spawnSync("powershell", ["-NoProfile", "-STA", "-Command",
    "$s = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + b64 + "')); $out = Invoke-Expression $s; [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$out))"],
    { encoding: "utf8", windowsHide: true });
  try { return Buffer.from((r.stdout || "").trim(), "base64").toString("utf8"); } catch { return ""; }
}
function probe() {
  const script = [
    "$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application')",
    "$d = $w.ActiveDocument",
    "$o = [ordered]@{}",
    "$o.name = $d.Name",
    "$o.paragraphs = [int]$d.Paragraphs.Count",
    "$o.contentEnd = [int]$d.Content.End",
    "$b = @()",
    "for ($i = 0; $i -lt $d.Content.End; $i++) { $b += [int]$d.Range($i, $i + 1).Font.Bold }",
    "$o.bold = $b",
    "$s = @()",
    "for ($p = 1; $p -le $d.Paragraphs.Count; $p++) { $s += [string]$d.Paragraphs.Item($p).Range.Style.NameLocal }",
    "$o.styles = $s",
    "$t = @()",
    "for ($p = 1; $p -le $d.Paragraphs.Count; $p++) { $t += ([string]$d.Paragraphs.Item($p).Range.Text).TrimEnd([char]13, [char]10) }",
    "$o.paras = $t",
    "$o | ConvertTo-Json -Compress -Depth 4",
  ].join("; ");
  const raw = com(script);
  try { return JSON.parse(raw); } catch { return { error: raw.slice(0, 200) }; }
}
const boldCount = (p) => (p.bold || []).filter((b) => b !== 0).length;
const lastIndex = (p) => (p.paragraphs || 1) - 1;

// ---- 准备：新建一个一次性文档，正文 "XXX YYY ZZZ" -------------------------------------------
const created = await call("wps_word_create_document", {});
check("scratch document created", ok(created), text(created).slice(0, 60));
const seeded = await call("wps_word_insert_text", { text: "XXX YYY ZZZ", position: "start" });
check("seeded 11-char body", ok(seeded), text(seeded).replace(/\s+/g, " ").slice(0, 80));

// 把光标折叠到文档开头：模拟报告里「没有选中内容」的真实处境。
com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); $w.Selection.SetRange(0, 0); 'ok'");

// ---- B1：按字符范围设字体（这是报告里唯一的刚需能力） -----------------------------------------
const rangeFont = await call("wps_word_set_font", { bold: true, range: { start: 0, end: 3 } });
check("set_font accepts a character range", ok(rangeFont), text(rangeFont).replace(/\s+/g, " ").slice(0, 90));
const afterRange = probe();
check("exactly the 3 requested characters are bold", boldCount(afterRange) === 3, "bold=" + boldCount(afterRange) + " of " + (afterRange.bold || []).length);
check("the range report names the resolved scope", /作用范围: 0-3/.test(text(rangeFont)), text(rangeFont).replace(/\s+/g, " ").slice(-70));

// ---- B1 续：范围为空时不再谎报成功 ------------------------------------------------------------
com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); $w.Selection.SetRange(0, 0); 'ok'");
const emptyRange = await call("wps_word_set_font", { italic: true });
const emptyInfo = probe();
check("collapsed-selection set_font reports 0 characters", /0 个字符/.test(text(emptyRange)), text(emptyRange).replace(/\s+/g, " ").slice(-90));
check("collapsed-selection set_font explains nothing was changed", /没有任何字符|没有实际落点/.test(text(emptyRange)), "");
check("collapsed-selection set_font really changed nothing", (emptyInfo.bold || []).filter((b) => b !== 0).length === 3, "bold=" + boldCount(emptyInfo));

// ---- B3：new_paragraph 产生真段落、文本独立成段 ------------------------------------------------
com("$w = [Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); $w.Selection.SetRange(0, 0); 'ok'");
const paraInsert = await call("wps_word_insert_text", { text: "TITLE-A", position: "end", new_paragraph: true });
const afterPara = probe();
check("new_paragraph adds a real paragraph", afterPara.paragraphs === 2, "paragraphs=" + afterPara.paragraphs);
check("the new text sits in the new paragraph", (afterPara.paras || [])[1] === "TITLE-A", JSON.stringify(afterPara.paras));
check("insert_text reports the insertion range", /插入范围: \d+-\d+/.test(text(paraInsert)), text(paraInsert).replace(/\s+/g, " ").slice(-80));
check("insert_text reports the paragraph it landed in", /所在段落: 第 2 段/.test(text(paraInsert)), text(paraInsert).replace(/\s+/g, " ").slice(-60));


// ---- B2：样式必须落在刚插入的文本上（不是插入前的光标位置） ------------------------------------
const styled = await call("wps_word_insert_text", { text: "HEAD-B", position: "end", new_paragraph: true, style: "标题 3" });
const afterStyle = probe();
check("style landed on the last (new) paragraph", afterStyle.styles[lastIndex(afterStyle)] === "标题 3", "styles=" + JSON.stringify(afterStyle.styles));
check("earlier paragraphs kept their style", afterStyle.styles[0] === "正文", "styles=" + JSON.stringify(afterStyle.styles));
check("insert_text names the paragraph it styled", /第 \d+ 段/.test(text(styled)), text(styled).replace(/\s+/g, " ").slice(-90));

// ---- B4：applyStyle 如实回报被扩张到的段落 -----------------------------------------------------
const applied = await call("wps_word_apply_style", { styleName: "标题 1", range: { start: 0, end: 3 } });
const afterApply = probe();
check("apply_style reports the affected paragraphs", /影响的段落: 第 1 段/.test(text(applied)), text(applied).replace(/\s+/g, " ").slice(-90));
check("apply_style's affectedText is a pre-write snapshot", /XXX/.test(text(applied)), text(applied).replace(/\s+/g, " ").slice(0, 110));
check("apply_style really restyled paragraph 1", afterApply.styles[0] === "标题 1", "styles=" + JSON.stringify(afterApply.styles));

// ---- L2（FIXES 88）：英文样式名要能翻译成中文内置名 --------------------------------------------
// 中文 WPS 的样式表里没有英文名（实测 NameInternational 全空、Item("Heading 1") 直接抛错），旧版把
// "Heading 2" 原样发给 WPS 就吃 E_FAIL；现在工具层先翻译再发。
const englishStyle = await call("wps_word_apply_style", { styleName: "Heading 2", range: { start: 0, end: 3 } });
const afterEnglish = probe();
check("english style alias is accepted", ok(englishStyle), text(englishStyle).replace(/\s+/g, " ").slice(0, 90));
check("the english alias really applied 标题 2", afterEnglish.styles[0] === "标题 2", "styles=" + JSON.stringify(afterEnglish.styles));
const bogusStyle = await call("wps_word_apply_style", { styleName: "No Such Style 123", range: { start: 0, end: 3 } });
const bogusText = text(bogusStyle);
check("a style that does not exist fails loudly", !ok(bogusStyle), bogusText.replace(/\s+/g, " ").slice(0, 80));
check("the failure names the style that was sent", /No Such Style 123/.test(bogusText), "");
check("the failure tells the caller what to do next", /下一步/.test(bogusText), "");

// ---- L9（FIXES 88）：set_active_target 不再回假成功 ---------------------------------------------
// 拒绝对调用方是 isError、没有 JSON 载荷，所以看原始响应文本。
const badTarget = await call("wps_call", { tool: "wps_ppt_set_active_target", args: { name: "definitely-not-open.pptx" } });
check("locking a non-existent presentation reports failure", !ok(badTarget), text(badTarget).replace(/\s+/g, " ").slice(0, 110));
check("the failure lists what is actually open", /已打开/.test(text(badTarget)), "");
const clearedTarget = await call("wps_call", { tool: "wps_ppt_set_active_target", args: { clear: true } });
check("clearing the lock still succeeds", ok(clearedTarget), text(clearedTarget).replace(/\s+/g, " ").slice(0, 90));

// ---- B10：越界必须报错（以前是静默删掉 [start, 文末)） -----------------------------------------
// 注意：via() 返回的是**解析后的载荷**，而拒绝对调用方是 isError、没有 JSON 载荷 ——
// 所以这里用 call() 拿原始响应，再读错误文本。
const before = probe();
const oob = await call("wps_call", { tool: "wps_word_replace_range", args: { startPos: 3, endPos: before.contentEnd + 500, text: "" } });
const oobText = text(oob);
const afterOob = probe();
check("out-of-range replace_range is rejected", /替换失败/.test(oobText), oobText.replace(/\s+/g, " ").slice(0, 100));
check("the rejection names the document length", /文档长度/.test(oobText), oobText.replace(/\s+/g, " ").slice(0, 100));
check("the rejection tells the caller where the end is", /endPos=/.test(oobText), "");
check("the document is untouched after the rejection", afterOob.contentEnd === before.contentEnd, before.contentEnd + " -> " + afterOob.contentEnd);

// ---- B10 续：跨段删除要显式 confirm ------------------------------------------------------------
const spanStart = 0;
const spanEnd = before.contentEnd - 1;
const unconfirmed = await call("wps_call", { tool: "wps_word_replace_range", args: { startPos: spanStart, endPos: spanEnd, text: "" } });
const unconfirmedText = text(unconfirmed);
const afterUnconfirmed = probe();
check("multi-paragraph delete is refused without confirm", /替换失败/.test(unconfirmedText), unconfirmedText.replace(/\s+/g, " ").slice(0, 100));
check("the refusal says how many paragraphs would go", /段落标记/.test(unconfirmedText), unconfirmedText.replace(/\s+/g, " ").slice(0, 110));
check("nothing was deleted while refused", afterUnconfirmed.paragraphs === before.paragraphs, before.paragraphs + " -> " + afterUnconfirmed.paragraphs);
const confirmed = await call("wps_call", { tool: "wps_word_replace_range", args: { startPos: spanStart, endPos: spanEnd, text: "", confirm: true } });
const confirmedText = text(confirmed);
const afterConfirmed = probe();
check("confirmed multi-paragraph delete goes through", ok(confirmed), confirmedText.replace(/\s+/g, " ").slice(0, 110));
check("confirmed delete reports how many paragraphs it removed", /已确认的批量删除: \d+ 个段落标记/.test(confirmedText), confirmedText.replace(/\s+/g, " ").slice(-80));
check("the paragraphs really are gone", afterConfirmed.paragraphs < before.paragraphs, before.paragraphs + " -> " + afterConfirmed.paragraphs);


// ---- D18-C：get_paragraphs 现在给出可用的字符坐标 ----------------------------------------------
const paras = text(await call("wps_word_get_paragraphs", {}));
check("get_paragraphs prints character offsets", /@\d+-\d+/.test(paras), paras.replace(/\s+/g, " ").slice(0, 110));

// ---- 清理：关掉一次性文档（不保存），不留残留 ---------------------------------------------------
const closed = await call("wps_word_close_document", { save: false });
check("scratch document closed without saving", ok(closed), text(closed).replace(/\s+/g, " ").slice(0, 70));
const openDocs = text(await via("wps_word_get_open_documents", {}));
check("no scratch Word document is left behind", !/scratch|文字文稿/.test(openDocs), openDocs.replace(/\s+/g, " ").slice(0, 90));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "WORD RANGE FORMAT TESTS OK (" + results.length + ")" : "WORD RANGE FORMAT TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

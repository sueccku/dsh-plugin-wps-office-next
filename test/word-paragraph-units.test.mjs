// 回归：wps_word_set_paragraph 的**单位语义**必须与它自己的 schema 一致。
//
// 背景（S2 场景真机走查抓到的两个契约缺陷，都是"回 success 但落盘的不是要求的东西"）：
//   1. lineSpacing 的 schema 写的是"行间距倍数（正数），如 1.5、2"，桥却写
//      `LineSpacingRule = 4`（wdLineSpaceExactly/固定值）+ `1.5 * 12` 磅 —— 落盘成
//      `<w:lineRule="exact"/>`。用户要的是 1.5 倍，拿到的是"固定值 18 磅"：字号一变行距就错，
//      WPS 界面里显示成"固定值"。
//      同一份代码里的兄弟动作 setLineSpacing 用的是 `LineSpacingRule = 5`（multiple）→ 正确，
//      所以这是 setParagraph 单独走偏。
//   2. firstLineIndent / leftIndent / rightIndent 的 schema 写的是"字符数"，桥却 `* 28.35`
//      （厘米→磅）—— 要"2 字符"落盘成 2 厘米 = 1134 缇 ≈ 4.7 字符（小四）。
//      WPS 的 COM 里 `CharacterUnit*Indent` 是现成的字符单位属性，落盘为 `w:firstLineChars="200"`
//      这种随字号缩放的正确写法。
//
// 判据全部以**落盘后的 XML** 与**裸 COM 读回值**为准（不读插件自述）。
// Run: node test/word-paragraph-units.test.mjs   (needs WPS)
import { spawn, spawnSync } from "node:child_process";
import { rmSync } from "node:fs";
import { resolve as resolvePath } from "node:path";
import { zipScan } from "./lib/oracle.mjs";

const OUT = resolvePath("test/.artifacts/word-paragraph-units.docx");
try { rmSync(OUT, { force: true }); } catch { /* ignore */ }

const child = spawn(process.execPath, ["mcp/dist/index.js"], {
  stdio: ["pipe", "pipe", "pipe"],
  windowsHide: true,
  // 钉住仓库内的宿主：从"已装本插件"的会话里起的终端会带着 profile 副本的路径（FIXES 87）。
  env: { ...process.env, WPS_OFFICE_HOST_SCRIPT: resolvePath("host/wps-com-host.ps1") },
});
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

const results = [];
function check(name, okFlag, detail) { results.push({ name, ok: !!okFlag }); console.log((okFlag ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "word-paragraph-units", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const via = (tool, args) => call("wps_call", { tool, args: args || {} });

// 裸 COM 读回真值（只读）：输出走 base64，避免中文/编码链被改坏。
function com(script) {
  const b64 = Buffer.from(script, "utf8").toString("base64");
  const r = spawnSync("powershell", ["-NoProfile", "-STA", "-Command",
    "$s = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + b64 + "')); $out = Invoke-Expression $s; [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$out))"],
    { encoding: "utf8", windowsHide: true });
  try { return Buffer.from((r.stdout || "").trim(), "base64").toString("utf8"); } catch { return ""; }
}
const fmt = () => com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); $f=$w.ActiveDocument.Paragraphs.Item(1).Format; ($f.FirstLineIndent.ToString([Globalization.CultureInfo]::InvariantCulture) + '|' + $f.LeftIndent.ToString([Globalization.CultureInfo]::InvariantCulture) + '|' + $f.LineSpacingRule + '|' + $f.LineSpacing.ToString([Globalization.CultureInfo]::InvariantCulture))");

const created = await call("wps_word_create_document", {});
check("scratch document created", ok(created), text(created).slice(0, 60));
await call("wps_word_insert_text", { text: "单位回归测试段落。", position: "end" });
await call("wps_word_set_font", { fontName: "宋体", fontSize: 12, range: "all" });

// 一次调用里同时给"倍数行距"和"字符缩进"——这正是 S2 场景里的用法。
const applied = await call("wps_word_set_paragraph", {
  lineSpacing: 1.5, firstLineIndent: 2, leftIndent: 1, rightIndent: 1,
});
check("set_paragraph reports success", ok(applied), text(applied).replace(/\s+/g, " ").slice(0, 120));

const saved = await call("wps_common_save_as", { outputPath: OUT });
check("scratch document saved", ok(saved), text(saved).replace(/\s+/g, " ").slice(0, 80));
// 裸 COM 读回必须在**关闭文档之前**做（关掉之后 ActiveDocument 就没了 → 读到空串）。
const readback = fmt();
await call("wps_word_close_document", { save: false });
await new Promise((r) => setTimeout(r, 300));

// ---------- 1) 行距：倍数必须落成 lineRule="auto"，不能是"固定值" ----------
// needle 必须写成 XML 里**真实出现**的连续片段：`w:lineRule="exact"` 是属性，
// 前面跟的是同一元素里的别的属性（`<w:spacing w:line="360" …`），不是 `<`。
const exact = zipScan(OUT, 'w:lineRule="exact"', "word/document.xml");
const auto = zipScan(OUT, 'w:lineRule="auto"', "word/document.xml");
const line360 = zipScan(OUT, 'w:line="360" w:lineRule="auto"', "word/document.xml");
check("lineSpacing=1.5 does not land as a fixed ('exact') line height", exact && exact.count === 0, JSON.stringify(exact));
check("lineSpacing=1.5 lands as a multiple ('auto')", auto && auto.count >= 1, JSON.stringify(auto));
check("the multiple is 1.5 (w:line=360 with auto)", line360 && line360.count >= 1, JSON.stringify(line360));

// ---------- 2) 缩进：字符单位必须落成 *Chars ----------
const firstChars = zipScan(OUT, 'w:firstLineChars="200"', "word/document.xml");
const leftChars = zipScan(OUT, 'w:leftChars="100"', "word/document.xml");
const rightChars = zipScan(OUT, 'w:rightChars="100"', "word/document.xml");
check("firstLineIndent=2 lands in character units (w:firstLineChars=200)", firstChars && firstChars.count >= 1, JSON.stringify(firstChars));
check("leftIndent=1 lands in character units (w:leftChars=100)", leftChars && leftChars.count >= 1, JSON.stringify(leftChars));
check("rightIndent=1 lands in character units (w:rightChars=100)", rightChars && rightChars.count >= 1, JSON.stringify(rightChars));

// ---------- 3) 裸 COM 读回：量级必须对得上"字符"，不是"厘米" ----------
// 注意：FirstLineIndent 的**点值**是 WPS 按当前字宽现算的（这里 1 字符 ≈ 10.5pt → 2 字符 ≈ 21pt），
// 所以断言的是量级而不是硬编码 24pt：2 字符落在 12–36pt，1 字符落在 8–20pt。
// 修复前的值是 56.7pt / 28.35pt（= 2cm / 1cm），离"字符"差着 2 倍以上，会被这条挡住。
const parts = readback.split("|").map((s) => Number(s));
const [flPt, leftPt, rule, lsPt] = parts;
check("raw COM first-line indent is character-sized (about 2 chars), not 2 cm", Number.isFinite(flPt) && flPt >= 12 && flPt <= 36, readback);
check("raw COM left indent is character-sized (about 1 char), not 1 cm", Number.isFinite(leftPt) && leftPt >= 8 && leftPt <= 20, readback);
check("raw COM reports a multiple line-spacing rule, not 'exactly'", Number.isFinite(rule) && rule !== 4, readback);
check("raw COM line spacing is 1.5 lines worth of points", Number.isFinite(lsPt) && Math.abs(lsPt - 18) <= 0.5, readback);

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "WORD PARAGRAPH UNITS TESTS OK (" + results.length + ")" : "WORD PARAGRAPH UNITS TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

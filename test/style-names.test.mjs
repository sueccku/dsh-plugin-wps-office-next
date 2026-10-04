// FIXES 88（L2）：样式名解析的纯函数单测。不需要 WPS，已进 CI。
// 背景：中文版 WPS 的样式表里没有英文名 —— 实测 `Styles.Item("Heading 1")`、`Range.Style = "Heading 1"`
// 与 `Style.NameInternational` 全部失败/为空（478 个样式一律空串）。工具描述却把 "Heading 1" 当示例，
// 模型照着调用就吃 E_FAIL。这里验证「英文别名 / 少空格 / 全角空格」都能翻译成 WPS 认的名字。
// Run: node test/style-names.test.mjs
import { normalizeStyleName, resolveStyleName } from '../mcp/dist/tools/word/style-names.js';

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }

// --- 英文别名必须翻译成中文内置名（含大小写与空格变体） ---
const headingCases = [
  ["Heading 1", "标题 1"],
  ["heading 2", "标题 2"],
  ["HEADING 3", "标题 3"],
  ["Heading1", "标题 1"],
  ["Heading  4", "标题 4"],
  ["heading\u30005", "标题 5"],
  ["Heading 9", "标题 9"],
];
for (const [input, expected] of headingCases) {
  const r = resolveStyleName(input);
  check("heading alias: " + JSON.stringify(input) + " -> " + expected, r.name === expected && r.translated, JSON.stringify(r));
}

const aliasCases = [
  ["Normal", "正文"],
  ["normal", "正文"],
  ["Body Text", "正文文本"],
  ["Title", "标题"],
  ["Subtitle", "副标题"],
  ["Quote", "引用"],
  ["Intense Quote", "明显引用"],
  ["Emphasis", "强调"],
  ["Caption", "题注"],
  ["List Paragraph", "列表段落"],
];
for (const [input, expected] of aliasCases) {
  const r = resolveStyleName(input);
  check("builtin alias: " + JSON.stringify(input) + " -> " + expected, r.name === expected && r.translated, JSON.stringify(r));
}

// --- 中文名：原样保留，只在"少空格/多空格"时修正 ---
const cjkCases = [
  ["标题 1", "标题 1"],
  ["标题1", "标题 1"],
  ["标题  1", "标题 1"],
  ["副标题", "副标题"],
  ["正文", "正文"],
  ["  正文  ", "正文"],
];
for (const [input, expected] of cjkCases) {
  const r = resolveStyleName(input);
  check("cjk name: " + JSON.stringify(input) + " -> " + expected, r.name === expected, JSON.stringify(r));
}

// --- 自定义样式名不能被改写（否则会悄悄套错用户自己的样式） ---
check("custom name passes through unchanged", resolveStyleName("我的公司样式").name === "我的公司样式", "我的公司样式");
check("custom name with spaces is only whitespace-normalized", resolveStyleName("我的  公司 样式").name === "我的 公司 样式", "");
check("english alias is not reported as translated for custom names", resolveStyleName("My Style").translated === false, "");
check("unknown english heading-like name is left alone", resolveStyleName("Heading 10").name === "Heading 10", "Heading 10");
check("empty input stays empty", normalizeStyleName("   ") === "" && resolveStyleName("").name === "", "");

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "STYLE NAMES TESTS OK (" + results.length + ")" : "STYLE NAMES TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

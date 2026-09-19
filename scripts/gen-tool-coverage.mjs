// Generate docs/tool-coverage.md: the tool x action x advertised x tested matrix.
//
// Inputs: spec/tool-definitions.json (the catalog), spec/advertised.json (the advertised surface),
//         the static tool->action parse of mcp/src/tools (scripts/lib/tool-action-map.mjs), and the
//         test corpus that names tools.
// Output: docs/tool-coverage.md (deterministic; CI regenerates and diffs it).
//
// 一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的覆盖矩阵一节。
import { readFileSync, readdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { analyseToolSource } from "./lib/tool-action-map.mjs";

const B = String.fromCharCode(96); // markdown code span, kept out of the source so this file stays quote-safe
const code = (s) => B + s + B;
const root = new URL("..", import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, "$1");

const defs = JSON.parse(readFileSync(join(root, "spec", "tool-definitions.json"), "utf8"));
const advertised = new Set(JSON.parse(readFileSync(join(root, "spec", "advertised.json"), "utf8")));
const { map: toolAction } = analyseToolSource();

// Which test files (and the e2e run) name each tool.
const testFiles = readdirSync(join(root, "test")).filter((f) => f.endsWith(".test.mjs")).sort();
const corpus = testFiles.map((f) => [f, readFileSync(join(root, "test", f), "utf8")]);
const e2e = readFileSync(join(root, "scripts", "e2e.mjs"), "utf8");
function namedBy(tool) {
  const hits = [];
  for (const [f, text] of corpus) if (text.includes(tool)) hits.push(f);
  if (e2e.includes(tool)) hits.push("e2e");
  return hits;
}

const appOf = (name) => {
  const m = /^wps_([a-z]+)_/.exec(name);
  const key = m ? m[1] : "other";
  const labels = { excel: "Excel", word: "Word", ppt: "PPT", common: "通用", convert: "转换", other: "其他" };
  return labels[key] || key;
};

const rows = defs.map((d) => {
  const evidence = namedBy(d.name);
  const ta = toolAction.get(d.name);
  return { name: d.name, app: appOf(d.name), action: ta ? ta.action : "", advertised: advertised.has(d.name), tested: evidence.length > 0, evidence };
}).sort((a, b) => a.name.localeCompare(b.name));

const apps = ["Excel", "Word", "PPT", "通用", "转换", "其他"];
const summary = apps.map((app) => {
  const group = rows.filter((r) => r.app === app);
  if (group.length === 0) return null;
  return { app, total: group.length, tested: group.filter((r) => r.tested).length, advertised: group.filter((r) => r.advertised).length };
}).filter(Boolean);

const untested = rows.filter((r) => !r.tested);
const noAction = rows.filter((r) => !r.action);
const evidenceCell = (list) => {
  if (list.length === 0) return "—";
  const shown = list.slice(0, 3).join(", ");
  return list.length > 3 ? shown + ", +" + (list.length - 3) : shown;
};

const lines = [];
lines.push("# 工具覆盖矩阵（生成物）");
lines.push("");
lines.push("> 由 " + code("node scripts/gen-tool-coverage.mjs") + " 生成，**请勿手改**；CI 会重新生成并对账。");
lines.push("> 数据来源：" + code("spec/tool-definitions.json") + "（工具目录）、" + code("spec/advertised.json") + "（广告面）、");
lines.push("> " + code("mcp/src/tools/**") + " 的静态解析（工具 → 桥 action，见 " + code("scripts/lib/tool-action-map.mjs") + "）、");
lines.push("> " + code("test/*.test.mjs") + " 与 " + code("scripts/e2e.mjs") + "（工具名出现过 = 覆盖证据）。");
lines.push("");
lines.push("## 汇总");
lines.push("");
lines.push("| 应用 | 工具数 | 被测试点名 | 广告面 |");
lines.push("| --- | ---: | ---: | ---: |");
for (const s of summary) lines.push("| " + s.app + " | " + s.total + " | " + s.tested + " | " + s.advertised + " |");
lines.push("| **合计** | **" + rows.length + "** | **" + rows.filter((r) => r.tested).length + "** | **" + rows.filter((r) => r.advertised).length + "** |");
lines.push("");
lines.push("## 矩阵");
lines.push("");
lines.push("「广告」= 每次请求随 tools/list 下发的 69 个工具之一；其余经 " + code("wps_call") + " / " + code("wps_help") + " 触达。");
lines.push("");
lines.push("| 工具 | 应用 | 桥 action | 广告 | 被测试点名（证据） |");
lines.push("| --- | --- | --- | :---: | --- |");
for (const r of rows) {
  lines.push("| " + code(r.name) + " | " + r.app + " | " + (r.action ? code(r.action) : "—") + " | " + (r.advertised ? "✅" : "") + " | " + evidenceCell(r.evidence) + " |");
}
lines.push("");
lines.push("## 未被测试点名的工具（" + untested.length + "）");
lines.push("");
if (untested.length === 0) lines.push("无——" + rows.length + " 个工具全部至少被一个测试或 e2e 点名。");
else for (const r of untested) lines.push("- " + code(r.name));
lines.push("");
lines.push("## 解析不出桥 action 的工具（" + noAction.length + "）");
lines.push("");
if (noAction.length === 0) lines.push("无——每个工具都能从 handler 源码静态解析出它驱动的 action。");
else for (const r of noAction) lines.push("- " + code(r.name));
lines.push("");
lines.push("## 已知 WPS 不支持（实测，不再尝试）");
lines.push("");
lines.push("以下不是「没做」，是 COM 层面不成立（证据见 " + code("docs/FIXES.md") + " 与探针脚本），推广材料应写「不支持」：");
lines.push("");
lines.push("- Word 水印：页眉 " + code("Shapes") + " 拒绝一切添加，" + code("Count") + " 恒为 0。");
lines.push("- 文档属性：" + code("BuiltInDocumentProperties") + " / " + code("CustomDocumentProperties") + " 是坏壳。");
lines.push("- Excel 切片器：" + code("SlicerCaches.Add2") + " 可调用，但 " + code("Slicers.Count") + " 恒为 0。");
lines.push("- 方案管理器：" + code("Worksheet.Scenarios") + " 在 COM 里是方法，语义读不干净。");
lines.push("");
lines.push("## 说明");
lines.push("");
lines.push("- 「被测试点名」只统计工具名在测试 / e2e 源码里出现过，是覆盖的**必要条件**，不是充分条件；真正验证行为的是各场景测试本身。");
lines.push("- 覆盖率的 ratchet 断言在 " + code("test/spec-reproduction.test.mjs") + "；未覆盖清单也可用 " + code("node scripts/smoke-tools.mjs") + " 查看。");
lines.push("");

writeFileSync(join(root, "docs", "tool-coverage.md"), lines.join("\n"));
console.log("wrote docs/tool-coverage.md: " + rows.length + " tools, " + untested.length + " untested, " + noAction.length + " without an action");

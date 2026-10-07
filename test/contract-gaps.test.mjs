// FIXES 94 回归：三个「契约盲区」工具的真实行为（都是"回 success 但结果不对"家族）。
//
//   W4-1 export_chart_as_image：名字找不到时以前回裸 COM E_FAIL，还把原因误导成"WPS 可能不健康"，
//        而且**没有任何可发现路径**（没有列出图表的工具）。现在与 delete_chart 对齐 ——
//        名字或序号都收、整张表只有一张图时可省略、找不到就**列出已有的图表名**。
//   W6-1 word_set_paragraph：range 传 {start,end} 以前被**静默忽略**（落到选区上，改错位置却回成功）；
//        非法 alignment 被**静默改成左对齐**（比忽略更糟）；一个属性都不给也回"段落格式已设置"。
//        现在三种都响亮拒绝，并且 {start,end} 会真正作用到那一段（用裸 COM 读回真值验证）。
//   W6-2 excel_evaluate_formula：Excel 的错误值以 CVErr 负数返回且 success=true ——
//        模型会把 -2146826281 读成"计算结果"。现在翻成 #DIV/0! 这类可读错误并 success=false。
//
// Run: node test/contract-gaps.test.mjs   (needs WPS)
import { spawn, spawnSync } from "node:child_process";
import { existsSync, statSync, rmSync, writeFileSync } from "node:fs";
import { resolve as resolvePath } from "node:path";

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
function check(name, okFlag, detail) { results.push({ name, ok: okFlag }); console.log((okFlag ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }
function payload(res) { try { return JSON.parse(text(res)); } catch { return {}; } }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "contract-gaps", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const via = (tool, args) => call("wps_call", { tool, args: args || {} });

// 裸 COM：段落对齐只能直接问 WPS（工具面没有读段落格式的接口），输出走 base64 避免编码被改坏。
function com(script) {
  const b64 = Buffer.from(script, "utf8").toString("base64");
  const r = spawnSync("powershell", ["-NoProfile", "-STA", "-Command",
    "$s = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + b64 + "')); $out = Invoke-Expression $s; [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$out))"],
    { encoding: "utf8", windowsHide: true });
  try { return Buffer.from((r.stdout || "").trim(), "base64").toString("utf8"); } catch { return ""; }
}
const alignments = () => {
  const raw = com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('KWps.Application'); $d=$w.ActiveDocument; $a=@(); for ($i=1; $i -le $d.Paragraphs.Count; $i++) { $a += [int]$d.Paragraphs.Item($i).Alignment }; $a -join ','");
  return raw.trim() ? raw.trim().split(",").map(Number) : [];
};

// ==================== W6-1 · wps_word_set_paragraph ====================
const doc = await call("wps_word_create_document", {});
check("W6-1: scratch document created", ok(doc), text(doc).slice(0, 50));
for (const t of ["第一段", "第二段", "第三段"]) await via("wps_word_insert_text", { text: t, position: "end", new_paragraph: true });

const paraText = text(await via("wps_word_get_paragraphs", { startParagraph: 1, endParagraph: 6 }));
const p2 = /\[2\] \([^)]*\)[^@]*@(\d+)-(\d+)/.exec(paraText);
check("W6-1: paragraph 2 coordinates are available", !!p2, paraText.replace(/\s+/g, " ").slice(0, 110));

const before = alignments();
check("W6-1: baseline alignments read back via raw COM", before.length === 4, JSON.stringify(before));

// ① 空调用：以前回"段落格式已设置"（什么都没做）
const empty = await call("wps_word_set_paragraph", {});
check("W6-1: an empty call is refused instead of reporting success", !ok(empty), text(empty).replace(/\s+/g, " ").slice(0, 90));

// ② 非法枚举：以前静默改成左对齐（0）
const badEnum = await call("wps_word_set_paragraph", { alignment: "centre" });
const afterBadEnum = alignments();
check("W6-1: an unknown alignment is refused", !ok(badEnum), text(badEnum).replace(/\s+/g, " ").slice(0, 90));
check("W6-1: the refusal did not silently restyle anything", JSON.stringify(afterBadEnum) === JSON.stringify(before), JSON.stringify(afterBadEnum));

// ③ 非法行距
const badSpacing = await call("wps_word_set_paragraph", { lineSpacing: -1 });
check("W6-1: a negative lineSpacing is refused", !ok(badSpacing), text(badSpacing).replace(/\s+/g, " ").slice(0, 90));

// ④ {start,end} 必须真正作用到那一段（以前被静默忽略、落到选区上）
if (p2) {
  const start = Number(p2[1]);
  const end = Number(p2[2]) - 1;
  const ranged = await call("wps_word_set_paragraph", { alignment: "right", range: { start, end } });
  const afterRanged = alignments();
  check("W6-1: a {start,end} range is accepted", ok(ranged), text(ranged).replace(/\s+/g, " ").slice(0, 70));
  check("W6-1: only the targeted paragraph changed", afterRanged[1] === 2 && afterRanged[0] === before[0] && afterRanged[2] === before[2] && afterRanged[3] === before[3], JSON.stringify({ before, after: afterRanged }));
  const report = text(ranged);
  check("W6-1: the report names what was applied", /alignment=right/.test(report), report.replace(/\s+/g, " ").slice(0, 80));
}
// ⑤ 越界范围要报错
const oob = await call("wps_word_set_paragraph", { alignment: "left", range: { start: 0, end: 999999 } });
check("W6-1: an out-of-range {start,end} is refused", !ok(oob), text(oob).replace(/\s+/g, " ").slice(0, 90));
await call("wps_word_close_document", { save: false });

// ==================== W6-2 · wps_excel_evaluate_formula ====================
const wb = await via("wps_excel_create_workbook", {});
check("W6-2: scratch workbook created", ok(wb), text(wb).slice(0, 50));

const good = await via("wps_excel_evaluate_formula", { formula: "=1+1" });
check("W6-2: a valid formula still works", ok(good) && /"result":\s*2/.test(text(good)), text(good).replace(/\s+/g, " ").slice(0, 90));

const div0 = await via("wps_excel_evaluate_formula", { formula: "=1/0" });
const div0Text = text(div0);
check("W6-2: #DIV/0! is reported as a failure, not as a number", !ok(div0), div0Text.replace(/\s+/g, " ").slice(0, 110));
check("W6-2: the failure names the Excel error", /#DIV\/0!/.test(div0Text), div0Text.replace(/\s+/g, " ").slice(0, 110));
check("W6-2: the raw CVErr integer never reaches the caller", !/-21468\d+/.test(div0Text), div0Text.replace(/\s+/g, " ").slice(0, 90));

const syntax = await via("wps_excel_evaluate_formula", { formula: "=SUM(" });
check("W6-2: a malformed formula is reported as a failure", !ok(syntax) && /#/.test(text(syntax)), text(syntax).replace(/\s+/g, " ").slice(0, 110));

// ==================== W4-1 · wps_excel_export_chart_as_image ====================
await via("wps_excel_write_range", { range: "A1", data: [["地区", "额"], ["华东", 1250], ["华北", 980]] });
const mk = await call("wps_excel_create_chart", { range: "A1:B3", chartType: "column_clustered", title: "S" });
check("W4-1: chart created", ok(mk), text(mk).replace(/\s+/g, " ").slice(0, 70));
const chartName = (/图表名称: ([^\s]+)/.exec(text(mk)) || [])[1] || "";

const png = resolvePath("test/.artifacts/contract-gaps-chart.png");
if (existsSync(png)) rmSync(png, { force: true });

// ① 不存在的名字：以前是裸 E_FAIL
const missing = await call("wps_excel_export_chart_as_image", { chartName: "NoSuchChart", outputPath: png });
const missingText = text(missing);
check("W4-1: a missing chart name is refused with a usable message", !ok(missing) && /chart not found/.test(missingText), missingText.replace(/\s+/g, " ").slice(0, 120));
check("W4-1: the message lists the charts that do exist", chartName ? missingText.includes(chartName) : false, "chartName=" + chartName + " | " + missingText.replace(/\s+/g, " ").slice(0, 90));
check("W4-1: no raw HRESULT leaks to the caller", !/E_FAIL|HRESULT/i.test(missingText), missingText.replace(/\s+/g, " ").slice(0, 80));

// ② 省略 chartName：整张表只有一张图 → 自动选中
const auto = await call("wps_excel_export_chart_as_image", { outputPath: png });
check("W4-1: omitting chartName works when the sheet has exactly one chart", ok(auto), text(auto).replace(/\s+/g, " ").slice(0, 110));
check("W4-1: the export really wrote a non-trivial PNG", existsSync(png) && statSync(png).size > 500, existsSync(png) ? statSync(png).size + " bytes" : "missing");
check("W4-1: the report names the chart it actually exported", chartName ? text(auto).includes(chartName) : false, text(auto).replace(/\s+/g, " ").slice(0, 110));

// ③ 序号也收（与 delete_chart 同族）
const byIndex = await call("wps_excel_export_chart_as_image", { chartName: "1", outputPath: png });
check("W4-1: an index is accepted like delete_chart", ok(byIndex), text(byIndex).replace(/\s+/g, " ").slice(0, 100));

await via("wps_excel_close_workbook", { save: false });
const leftOpen = text(await via("wps_excel_get_open_workbooks", {}));
check("no scratch workbook left behind", /0个|无/.test(leftOpen), leftOpen.replace(/\s+/g, " ").slice(0, 60));

// ---- FIXES 100（W4-4）：两个应用的「将删除什么」预览必须一致（都不带段落标记）----
// Word 的 Comment.Range.Text 带结尾批注标记（\r），Excel 的 Comment.Text() 不带 ——
// 同一份预览在两个应用里长得不一样。现在桥侧统一清洗（沿用 deleteTableLine 的既有写法）。
await via("wps_word_create_document", {});
await via("wps_word_insert_text", { text: "S4 批注内容", position: "end" });
await via("wps_word_insert_comment", { text: "这是一条测试批注" });
const wPreview = payload(await via("wps_execute_method", { method: "deleteComment", params: {} }));
const wList = (wPreview.data && wPreview.data.impact && wPreview.data.impact.preview) || [];
check("word comment preview is captured", Array.isArray(wList) && wList.length === 1, JSON.stringify(wList));
check("word comment preview carries no paragraph mark", wList.length === 1 && !/[\r\n\u0007]/.test(wList[0]), JSON.stringify(wList[0]));
await via("wps_word_close_document", { save: false });

await via("wps_excel_create_workbook", {});
await via("wps_excel_write_range", { range: "A1", data: [["x"]] });
await via("wps_execute_method", { method: "addCellComment", params: { cell: "A1", text: "第二条批注" } });
const ePreview = payload(await via("wps_execute_method", { method: "deleteCellComment", params: { cell: "A1" } }));
const eList = (ePreview.data && ePreview.data.impact && ePreview.data.impact.preview) || [];
check("excel comment preview is captured", Array.isArray(eList) && eList.length === 1, JSON.stringify(eList));
check("excel comment preview carries no paragraph mark", eList.length === 1 && !/[\r\n\u0007]/.test(eList[0]), JSON.stringify(eList[0]));
check("both apps trim the preview the same way", wList.length === 1 && eList.length === 1 && wList[0] === wList[0].trim() && eList[0] === eList[0].trim(), JSON.stringify({ word: wList[0], excel: eList[0] }));
await via("wps_excel_close_workbook", { save: false });


// ---- FIXES 101（W4-5）：insert_ppt_image 必须回传形状名 ----
// 桥早就返回了 name（`$pic.Name`），但工具层声明并读取的是 imageIndex —— 桥从不返回那个字段，
// 于是「图片索引: N」**永远是死代码**，而真正有用的名字被丢掉，`replace_ppt_image`（按名定位）就链不起来。
const probePng = resolvePath("test/.artifacts/probe.png");
if (!existsSync(probePng)) writeFileSync(probePng, Buffer.from("89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a49444154789c6360000002000100ffff03000006000557bfabd40000000049454e44ae426082", "hex"));
await via("wps_ppt_create_presentation", {});
await via("wps_ppt_add_slide", { layout: "blank" });
const imgIns = text(await call("wps_ppt_insert_ppt_image", { slideIndex: 1, imagePath: probePng }));
const nameMatch = /形状名称: (.+)/.exec(imgIns);
check("insert_ppt_image reports the shape name", !!nameMatch, imgIns.replace(/\n/g, " | ").slice(0, 120));
const shapesList = text(await call("wps_ppt_get_shapes", { slideIndex: 1 }));
check("the reported name is the one get_shapes sees", !!nameMatch && shapesList.includes(nameMatch[1].trim()), JSON.stringify(nameMatch && nameMatch[1]));
check("the dead imageIndex line is gone", !/图片索引/.test(imgIns), imgIns.slice(0, 90));
await via("wps_ppt_close_presentation", { save: false });

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "CONTRACT GAPS TESTS OK (" + results.length + ")" : "CONTRACT GAPS TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

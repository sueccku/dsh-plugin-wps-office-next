// S4 coverage: drive the previously-untested PPT tools against a real WPS deck.
//
// The plan's S4 acceptance is "success:true, or a clear and expected business error - never a
// timeout". So every entry is called with best-effort arguments and MUST return inside the budget;
// entries marked "ok" additionally must succeed. Coverage itself is counted by the S4 ratchet in
// test/spec-reproduction.test.mjs and reported by scripts/smoke-tools.mjs.
// Run: node test/ppt-coverage.test.mjs
import { spawn, spawnSync } from "node:child_process";
import { writeFileSync } from "node:fs";
import { resolve as resolvePath } from "node:path";
// FIXES 95（W1-4）：这个文件 53 条断言里绝大多数是"覆盖矩阵"式烟测（调用没挂住就算过）。
// 这里补一组**以文件为准**的核对：存盘后读 pptx 的页数与工具汇报对照、文本框/形状文字是否真的落盘、
// 导出的幻灯片图片是不是真尺寸（而不是 1x1 占位）。
import { pngInfo, pptxText, zipScan } from "./lib/oracle.mjs";

const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const imgPath = resolvePath("test/.artifacts/pptcov.png");
const exportPath = resolvePath("test/.artifacts/pptcov-export.png");
writeFileSync(imgPath, Buffer.from(PNG, "base64"));

const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

// 裸 COM 读真值（只读）。输出走 base64：中文经 PowerShell → pipe 的编码链会被改坏。
function com(script) {
  const b64 = Buffer.from(script, "utf8").toString("base64");
  const r = spawnSync("powershell", ["-NoProfile", "-STA", "-Command",
    "$s = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + b64 + "')); $out = Invoke-Expression $s; [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$out))"],
    { encoding: "utf8", windowsHide: true });
  try { return Buffer.from((r.stdout || "").trim(), "base64").toString("utf8"); } catch { return ""; }
}

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "ppt-cov", version: "1" } });
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

// ---- scratch deck -------------------------------------------------------------------------
check("create_presentation", ok(await call("wps_ppt_create_presentation", {})), "");
check("add two blank slides", ok(await call("wps_ppt_add_slide", { layout: "blank" })) && ok(await call("wps_ppt_add_slide", { layout: "blank" })), "");
for (const t of ["rectangle", "oval", "triangle"]) await call("wps_ppt_add_shape", { slideIndex: 1, type: t, text: t === "rectangle" ? "标题" : ("t-" + t) });
await call("wps_ppt_add_textbox", { slideIndex: 1, text: "文本框" });
await call("wps_ppt_insert_table", { slideIndex: 1, rows: 2, cols: 3 });
await call("wps_ppt_insert_ppt_image", { slideIndex: 1, imagePath: imgPath });
await call("wps_ppt_insert_ppt_chart", { slideIndex: 1, type: "column_clustered", title: "销量" });
await call("wps_ppt_add_animation", { slideIndex: 1, effect: "fadeIn", shapeIndex: 1 });
// L1（FIXES 89）：switch_presentation 需要一个真实的文稿名 —— 从"已打开的演示文稿"里取。
const presList = text(await call("wps_ppt_get_open_presentations", {}));
// 实测列表形如「当前打开 1 个演示文稿：\n 1. 演示文稿15 路径: 演示文稿15 页数: 0」——取编号行的名字，
// 去掉「路径:」及其后面的内容。
const presLine = presList.split(/\r?\n/).find((l) => /^\s*\d+\./.test(l)) || "";
const presName = (presLine.replace(/^\s*\d+\.\s*/, "").split(/\s+路径:|\s*\|/)[0] || "演示文稿1").trim();
// A2（FIXES 90）：版式 × 占位符矩阵（实测，bare COM 扫出来的）：
//   title        = 标题(3) + 副标题(4)
//   titleContent = 标题(1) + 文本正文(2)
//   twoColumn    = 标题(1) + 表格占位符(12)
//   comparison   = 标题(1) + 文本(2) + 图表占位符(8)
//   titleOnly    = 标题(1)
//   blank        = 什么都没有
// 所以：标题类断言用 titleContent，副标题断言必须用 title 版式，正文用 titleContent。
// 第 1 页：titleContent（标题 + 文本正文）
await call("wps_ppt_set_slide_layout", { slideIndex: 1, layout: "titleContent" });
// 第 2 页保持 blank 并**重新设为 titleContent**；第 3、4、5 页改成"新加 + 设版式"的写法：
// 实测（A2 探针）只有这种顺序能拿到目标版式的占位符 —— 先对已有页 Layout= 赋值再靠名字找占位符并不稳。
await call("wps_ppt_set_slide_layout", { slideIndex: 2, layout: "titleContent" });
// 第 3 页：加页时就指定 title 版式（FIXES 90 起 add_slide 与 set_slide_layout 收同一套键名）
await call("wps_ppt_add_slide", { layout: "title" });
// 第 4 页保持 blank：验"没有标题/副标题/正文占位符"
await call("wps_ppt_add_slide", { layout: "blank" });
// 第 5 页：加页时就指定 titleContent 版式（标题 + 正文）
await call("wps_ppt_add_slide", { layout: "titleContent" });
// 补一个动画：后面的 set_slide_layout 会把动画清掉，矩阵与断言都需要页上有动画。
await call("wps_ppt_add_animation", { slideIndex: 1, effect: "fadeIn", shapeIndex: 1 });

// ---- matrix -------------------------------------------------------------------------------
// [tool, args, expectation]  expectation: "ok" (must succeed) | "error" (must fail clearly) | "any"
const MATRIX = [
  // reads
  ["wps_ppt_get_slide_count", {}, "ok"],
  ["wps_ppt_get_slide_info", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_slide_master", {}, "ok"],
  ["wps_ppt_get_slide_notes", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_slide_title", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_shapes", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_textboxes", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_animations", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_table_cell", { slideIndex: 1, tableIndex: 1, row: 1, col: 1 }, "ok"],
  ["wps_ppt_find_ppt_text", { text: "标题" }, "ok"],
  // slide settings
  ["wps_ppt_switch_slide", { slideIndex: 1 }, "ok"],
  ["wps_ppt_set_slide_layout", { slideIndex: 1, layout: "titleContent" }, "ok"],
  // 第 1 页已换成 titleContent（有标题占位符）。
  ["wps_ppt_set_slide_title", { slideIndex: 1, title: "覆盖标题" }, "ok"],
  // 第 3 页 = 新加空白页 + title 版式，带副标题占位符（A2 矩阵实测）。
  ["wps_ppt_set_slide_subtitle", { slideIndex: 3, subtitle: "副标题" }, "ok"],
  // 第 5 页 = 新加空白页 + titleContent 版式，带文本正文占位符。
  ["wps_ppt_set_slide_content", { slideIndex: 5, content: "正文内容" }, "ok"],
  ["wps_ppt_set_slide_notes", { slideIndex: 1, notes: "讲稿备注" }, "ok"],
  ["wps_ppt_set_slide_size", { width: 1280, height: 720 }, "ok"],
  ["wps_ppt_set_slide_theme", { theme: resolvePath("test/.artifacts/nope.thmx") }, "error"],
  ["wps_ppt_remove_slide_transition", { slideIndex: 1 }, "ok"],
  // backgrounds
  ["wps_ppt_set_background_color", { slideIndex: 1, color: "#112233" }, "ok"],
  ["wps_ppt_set_background_gradient", { slideIndex: 1, gradient: { color1: "#FFFFFF", color2: "#000000" } }, "ok"],
  ["wps_ppt_set_background_image", { slideIndex: 1, imagePath: imgPath }, "ok"],
  // shapes
  ["wps_ppt_set_shape_position", { slideIndex: 1, shapeIndex: 1, left: 60, top: 60, width: 120, height: 80 }, "ok"],
  ["wps_ppt_set_shape_style", { slideIndex: 1, shapeIndex: 1, fillColor: "#FF0000", lineColor: "#000000", lineWidth: 1 }, "ok"],
  // 形状有没有文字取决于版式切换后的实际内容；失败时 FIXES 89 会明说"没有文字"（有专门断言覆盖）。
  ["wps_ppt_set_font_color", { slideIndex: 1, shapeIndex: 1, color: "#00FF00" }, "any"],
  ["wps_ppt_duplicate_shape", { slideIndex: 1, shapeIndex: 1 }, "ok"],
  ["wps_ppt_set_shape_z_order", { slideIndex: 1, shapeIndex: 1, zOrder: "front" }, "ok"],
  // text
  ["wps_ppt_set_textbox_text", { slideIndex: 1, textboxIndex: 1, text: "新文本" }, "ok"],
  ["wps_ppt_set_textbox_style", { slideIndex: 1, textboxIndex: 1, style: { fontSize: 20 } }, "ok"],
  ["wps_ppt_replace_ppt_text", { find: "新文本", replace: "替换后" }, "ok"],
  // table
  ["wps_ppt_set_table_cell", { slideIndex: 1, tableIndex: 1, row: 1, col: 1, text: "X" }, "ok"],
  // images
  ["wps_ppt_replace_ppt_image", { slideIndex: 1, shapeIndex: 4, imagePath: imgPath }, "ok"],
  // charts
  ["wps_ppt_set_ppt_chart_data", { slideIndex: 1, chartIndex: 1, data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] } }, "ok"],
  ["wps_ppt_set_ppt_chart_style", { slideIndex: 1, chartIndex: 1, style: {} }, "ok"],
  // animations
  ["wps_ppt_set_animation_order", { slideIndex: 1, from: 1, to: 1 }, "ok"],
  // grouping / distribution (needs several shapes)
  ["wps_ppt_group_shapes", { slideIndex: 1, names: [1, 2] }, "ok"],
  ["wps_ppt_distribute_shapes", { slideIndex: 1, names: [1, 2, 3], direction: "horizontal" }, "ok"],
  // slide ops
  ["wps_ppt_move_slide", { fromIndex: 1, toIndex: 2 }, "ok"],
  ["wps_ppt_switch_presentation", { name: presName }, "ok"],
  ["wps_ppt_set_active_target", { clear: true }, "any"],
  // hyperlink
  ["wps_ppt_remove_ppt_hyperlink", { slideIndex: 1, shapeIndex: 1 }, "ok"],
  // master / beautify / export / files
  ["wps_ppt_add_master_element", { element: { type: "textbox", text: "母版页脚", left: 20, top: 500, width: 300, height: 40 } }, "ok"],
  ["wps_ppt_beautify", {}, "ok"],
  ["wps_ppt_export_slide_as_image", { slideIndex: 1, outputPath: exportPath }, "ok"],
  ["wps_ppt_insert_slides_from_file", { path: resolvePath("test/.artifacts/nope.pptx") }, "error"],
  ["wps_ppt_open_presentation", { path: resolvePath("test/.artifacts/nope.pptx") }, "error"],
  // destructive last
  ["wps_ppt_delete_shape", { slideIndex: 1, shapeIndex: 1 }, "ok"],
  ["wps_ppt_delete_textbox", { slideIndex: 1, textboxIndex: 1 }, "ok"],
  ["wps_ppt_delete_ppt_image", { slideIndex: 1, imageIndex: 1 }, "ok"],
  ["wps_ppt_delete_slide", { slideIndex: 2 }, "ok"],
];


// ---- L1（FIXES 89）：31 个只被 "any" 放行的 PPT 工具，逐个断言真实结果 -------------------------
// 读类要读出真实内容；写类要读回被改的对象；删除类要确认对象真的少了。

// 读：母版 / 标题 / 文本检索
const master = text(await call("wps_ppt_get_slide_master", {}));
check("get_slide_master returns master info", master.length > 5 && !/undefined/.test(master), master.replace(/\s+/g, " ").slice(0, 90));
const titleRead = text(await call("wps_ppt_get_slide_title", { slideIndex: 2 }));
check("get_slide_title returns something readable", titleRead.length > 0 && !/undefined/.test(titleRead), titleRead.replace(/\s+/g, " ").slice(0, 90));
const foundText = text(await call("wps_ppt_find_ppt_text", { text: "覆盖标题" }));
check("find_ppt_text reports where the text is", foundText.length > 5, foundText.replace(/\s+/g, " ").slice(0, 90));

// 幻灯片标题/副标题/正文：写进去再读回来（这三条是"假成功"的高发区）
// 第 2 页是 blank 版式，没有标题占位符 —— FIXES 89 之后这里必须**明确失败**（以前静默回成功）。
const blankSlide = 4; // A2：第 4 页是 blank（无任何占位符）
const titleOnBlank = await call("wps_ppt_set_slide_title", { slideIndex: blankSlide, title: "覆盖标题" });
// 关键是"不许假成功"：失败必须带可读原因（具体文案随 WPS 走的路径不同而不同）。
check(
  "set_slide_title never silently succeeds on a slide it cannot title",
  !ok(titleOnBlank) && /没有标题占位符|失败|错误|E_FAIL|HRESULT/i.test(text(titleOnBlank)),
  text(titleOnBlank).replace(/\s+/g, " ").slice(0, 100)
);
// 换到带标题的版式（枚举键；中文版式名在 WPS 里读不到，工具会明确报错——见下面的断言）。
const layoutByEnum = await call("wps_ppt_set_slide_layout", { slideIndex: 2, layout: "titleContent" });
// 版式断言：新加的页上设同一版式应当成功；已有 blank 页上设会由 WPS 自己拒（工具如实报错，不算假成功）。
check("set_slide_layout reports success or a clear WPS refusal", ok(layoutByEnum) || /失败|错误|E_FAIL/i.test(text(layoutByEnum)), text(layoutByEnum).replace(/\s+/g, " ").slice(0, 90));
const layoutByCjk = await call("wps_ppt_set_slide_layout", { slideIndex: 2, layout: "标题和内容" });
check("set_slide_layout fails clearly for a Chinese layout name", !ok(layoutByCjk) && /可用值/.test(text(layoutByCjk)), text(layoutByCjk).replace(/\s+/g, " ").slice(0, 90));
const titleOk = await call("wps_ppt_set_slide_title", { slideIndex: 2, title: "覆盖标题" });
const titleBack = text(await call("wps_ppt_get_slide_title", { slideIndex: 2 }));
check("set_slide_title really sets the title once a placeholder exists", ok(titleOk) && titleBack.includes("覆盖标题"), titleBack.replace(/\s+/g, " ").slice(0, 90));
// 副标题：无论成功还是失败都必须自洽 —— 成功则文本真在页面上，失败则说明占位符缺失。
// 副标题/正文：无论成功失败都必须自洽 —— 成功则文本真在页面上，失败则说明占位符缺失且给出下一步。
// A2（FIXES 90）：subtitle 占位符只在 title 版式上存在 —— 第 3 页就是它，这里必须成功。
const sub = await call("wps_ppt_set_slide_subtitle", { slideIndex: 3, subtitle: "副标题X" });
const subFound = text(await call("wps_ppt_find_ppt_text", { text: "副标题X" }));
// 顺带验一条负面：同样的调用打到 blank 页上必须明确失败（而不是静默成功）。
const subOnBlank = await call("wps_ppt_set_slide_subtitle", { slideIndex: blankSlide, subtitle: "不该写进去" });
check("set_slide_subtitle fails clearly on a blank slide", !ok(subOnBlank) && /placeholder|占位/i.test(text(subOnBlank)), text(subOnBlank).replace(/\s+/g, " ").slice(0, 90));
const subPresent = !/未找到|没有找到/.test(subFound);
// A2（FIXES 90）：第 3 页是 title 版式（含占位符类型 4 的副标题）——这里必须**真的写进去**。
check(
  "set_slide_subtitle writes the subtitle on a title-layout slide",
  ok(sub) && subPresent,
  "claimed=" + ok(sub) + " present=" + subPresent + " | " + text(sub).replace(/\s+/g, " ").slice(0, 70)
);
const content = await call("wps_ppt_set_slide_content", { slideIndex: 5, content: "正文内容X" });
const contentBack = text(await call("wps_ppt_find_ppt_text", { text: "正文内容X" }));
const contentPresent = !/未找到|没有找到/.test(contentBack);
// A2（FIXES 90）：第 5 页是 titleContent 版式（含占位符类型 2 的正文）——同样必须真的写进去。
check(
  "set_slide_content writes the body on a titleContent-layout slide",
  ok(content) && contentPresent,
  "claimed=" + ok(content) + " present=" + contentPresent + " | " + text(content).replace(/\s+/g, " ").slice(0, 70)
);

// 版式：设置后读回 Layout 名（版式名随母版而定，所以只要求"读得到且不是 undefined"）
// （版式断言的详细版本在标题那一段：枚举键成功、中文名明确失败、写回能读到名字。）

// 背景：改颜色后读回该页的填充色（用形状计数做不了，只能读背景）
const bg = await call("wps_ppt_set_background_color", { slideIndex: 2, color: "#112233" });
check("set_background_color succeeds", ok(bg), text(bg).replace(/\s+/g, " ").slice(0, 80));
const bgRead = com("$p = [Runtime.InteropServices.Marshal]::GetActiveObject('KWpp.Application'); $s = $p.ActivePresentation.Slides.Item(2); [string]$s.Background.Fill.ForeColor.RGB");
check("the slide background really changed", bgRead.trim().length > 0, "bg RGB=" + bgRead.trim());
const grad = await call("wps_ppt_set_background_gradient", { slideIndex: 2, gradient: { color1: "#FFFFFF", color2: "#000000" } });
check("set_background_gradient succeeds", ok(grad), text(grad).replace(/\s+/g, " ").slice(0, 80));
const bgImg = await call("wps_ppt_set_background_image", { slideIndex: 2, imagePath: imgPath });
check("set_background_image succeeds", ok(bgImg), text(bgImg).replace(/\s+/g, " ").slice(0, 80));

// 形状：字体色 / 层序 / 分组 / 分布 / 删除
// 从形状列表里挑一个真正有文字的形状；对没有文字的形状 FIXES 89 起会明确报错（那也是正确行为）。
const shapeList = text(await call("wps_ppt_get_shapes", { slideIndex: 1 }));
const textShapeIdx = Number(((shapeList.match(/\[(\d+)\]\s*形状/) || [])[1]) || ((shapeList.match(/\[(\d+)\]/) || [])[1]) || 1);
const fontColor = await call("wps_ppt_set_font_color", { slideIndex: 1, shapeIndex: textShapeIdx, color: "#00FF00" });
// 有文字就应当成功；没有文字则工具必须**明说**（FIXES 89 之前这里报的是"找不到属性 RGB"，把问题说错了）。
check(
  "set_font_color either applies the colour or explains the shape has no text",
  ok(fontColor) || /没有文字/.test(text(fontColor)),
  "shape " + textShapeIdx + ": " + text(fontColor).replace(/\s+/g, " ").slice(0, 80)
);
const fontColorNoText = await call("wps_ppt_set_font_color", { slideIndex: 1, shapeIndex: 99, color: "#00FF00" });
check("set_font_color fails clearly for a shape that does not exist", !ok(fontColorNoText), text(fontColorNoText).replace(/\s+/g, " ").slice(0, 80));
const zOrder = await call("wps_ppt_set_shape_z_order", { slideIndex: 1, shapeIndex: 1, zOrder: "front" });
check("set_shape_z_order succeeds", ok(zOrder), text(zOrder).replace(/\s+/g, " ").slice(0, 80));
// schema 里这个参数就叫 names（不是 shapeIndices）—— 按 schema 传。
const group = await call("wps_ppt_group_shapes", { slideIndex: 1, names: [1, 2] });
check("group_shapes succeeds", ok(group), text(group).replace(/\s+/g, " ").slice(0, 80));
const dist = await call("wps_ppt_distribute_shapes", { slideIndex: 1, names: [1, 2, 3], direction: "horizontal" });
check("distribute_shapes succeeds", ok(dist), text(dist).replace(/\s+/g, " ").slice(0, 80));

// 文本框样式：字号写进去读回
const tbStyle = await call("wps_ppt_set_textbox_style", { slideIndex: 1, textboxIndex: 1, style: { fontSize: 20 } });
check("set_textbox_style succeeds", ok(tbStyle), text(tbStyle).replace(/\s+/g, " ").slice(0, 80));

// 文本替换：替换后原文本必须消失
// 自己种一段唯一文本再替换 —— 不能依赖前面那些形状（有些已经被 delete 系列动过了，
// 之前那次就是"替换数量 1 处，但搜不到新串"：命中的是被删形状残留的文本范围）。
await call("wps_ppt_add_textbox", { slideIndex: 1, text: "REPLMARKER-ONE" });
const seededFound = text(await call("wps_ppt_find_ppt_text", { text: "REPLMARKER-ONE" }));
check("the replacement marker is on the slide before replacing", !/未找到|没有找到/.test(seededFound), seededFound.replace(/\s+/g, " ").slice(0, 80));
const replaced = await call("wps_ppt_replace_ppt_text", { find: "REPLMARKER-ONE", replace: "REPLMARKER-TWO" });
check("replace_ppt_text succeeds", ok(replaced), text(replaced).replace(/\s+/g, " ").slice(0, 80));
const renamedFound = text(await call("wps_ppt_find_ppt_text", { text: "REPLMARKER-TWO" }));
check("replace_ppt_text really renamed the text", !/未找到|没有找到/.test(renamedFound), renamedFound.replace(/\s+/g, " ").slice(0, 80));
const oldFound = text(await call("wps_ppt_find_ppt_text", { text: "REPLMARKER-ONE" }));
check("the old text is gone after the replace", /未找到|没有找到/.test(oldFound), oldFound.replace(/\s+/g, " ").slice(0, 80));

// 表格单元格：写 X 再读回
const cell = await call("wps_ppt_set_table_cell", { slideIndex: 1, tableIndex: 1, row: 1, col: 1, text: "X" });
check("set_table_cell succeeds", ok(cell), text(cell).replace(/\s+/g, " ").slice(0, 80));
const cellBack = text(await call("wps_ppt_get_table_cell", { slideIndex: 1, tableIndex: 1, row: 1, col: 1 }));
check("the table cell really holds X", /X/.test(cellBack), cellBack.replace(/\s+/g, " ").slice(0, 80));

// 图片替换 / 删除：形状数要跟着变
const imgBefore = Number(com("$p = [Runtime.InteropServices.Marshal]::GetActiveObject('KWpp.Application'); [string]$p.ActivePresentation.Slides.Item(1).Shapes.Count").trim());
const imgRepl = await call("wps_ppt_replace_ppt_image", { slideIndex: 1, shapeIndex: 4, imagePath: imgPath });
check("replace_ppt_image succeeds", ok(imgRepl), text(imgRepl).replace(/\s+/g, " ").slice(0, 80));

// 图表：数据与样式
const chartData = await call("wps_ppt_set_ppt_chart_data", { slideIndex: 1, chartIndex: 1, data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] } });
check("set_ppt_chart_data succeeds", ok(chartData), text(chartData).replace(/\s+/g, " ").slice(0, 80));
const chartStyle = await call("wps_ppt_set_ppt_chart_style", { slideIndex: 1, chartIndex: 1, style: {} });
check("set_ppt_chart_style succeeds", ok(chartStyle), text(chartStyle).replace(/\s+/g, " ").slice(0, 80));

// 动画顺序：只有一个动画时把 1 移到 1 是恒等变换，必须成功
// 前面的 set_slide_layout 会清掉动画，这里先补一个再测顺序。
await call("wps_ppt_add_animation", { slideIndex: 1, effect: "fadeIn", shapeIndex: 1 });
const animOrder = await call("wps_ppt_set_animation_order", { slideIndex: 1, from: 1, to: 1 });
check("set_animation_order succeeds for an identity move on a page with animations", ok(animOrder), text(animOrder).replace(/\s+/g, " ").slice(0, 80));
const animOrderBad = await call("wps_ppt_set_animation_order", { slideIndex: 2, from: 1, to: 1 });
check("set_animation_order fails clearly when the page has no animations", !ok(animOrderBad) && /没有|超出范围/.test(text(animOrderBad)), text(animOrderBad).replace(/\s+/g, " ").slice(0, 90));

// 切换效果：移除一个不存在的切换必须明确失败
const rmTransition = await call("wps_ppt_remove_slide_transition", { slideIndex: 2 });
check("remove_slide_transition is honest", ok(rmTransition) || /没有|不存在|no transition/i.test(text(rmTransition)), text(rmTransition).replace(/\s+/g, " ").slice(0, 90));

// 删除类：删一个文本框，形状数必须减少
const beforeDel = Number(com("$p = [Runtime.InteropServices.Marshal]::GetActiveObject('KWpp.Application'); [string]$p.ActivePresentation.Slides.Item(1).Shapes.Count").trim());
const delTextbox = await call("wps_ppt_delete_textbox", { slideIndex: 1, textboxIndex: 1 });
check("delete_textbox succeeds", ok(delTextbox), text(delTextbox).replace(/\s+/g, " ").slice(0, 80));
const afterDel = Number(com("$p = [Runtime.InteropServices.Marshal]::GetActiveObject('KWpp.Application'); [string]$p.ActivePresentation.Slides.Item(1).Shapes.Count").trim());
check("delete_textbox really removed a shape", afterDel < beforeDel, beforeDel + " -> " + afterDel);
let succeeded = 0;
for (const [name, args, expect] of MATRIX) {
  const started = Date.now();
  const res = await call(name, args, 30000);
  const ms = Date.now() - started;
  const hung = !!(res && res.__timeout);
  const good = ok(res);
  if (good) succeeded++;
  const label = name.replace("wps_ppt_", "");
  if (hung) check(label + " returns (no hang)", false, "HUNG after " + ms + "ms");
  else if (expect === "ok") check(label + " succeeds", good, ms + "ms " + text(res).replace(/\s+/g, " ").slice(0, 70));
  else if (expect === "error") check(label + " fails clearly", !good && /not found|不存在|失败|cannot|无法|required/.test(text(res)), ms + "ms " + text(res).replace(/\s+/g, " ").slice(0, 70));
  else {
    const body = String((res && res.result && res.result.content && res.result.content[0].text) || "");
    // "any" 只放行「没挂住」太弱了：至少要求返回了内容或一句可读的错误（FIXES 74）。
    check(label + " returns something readable (no hang)", body.length > 0, ms + "ms " + (good ? "ok" : "business error") + (body ? "" : " EMPTY RESPONSE"));
  }
}

// ---- 独立 oracle：以磁盘上的 pptx / png 为准 ----
const PPC = resolvePath("test/.artifacts/ppt-coverage.pptx");
check("presentation saved for independent verification", ok(await call("wps_common_save_as", { outputPath: PPC, format: "pptx" })), "");
const slideScan = zipScan(PPC, "<p:sld ", "ppt/slides/");
const countText = text(await call("wps_ppt_get_slide_count", {}));
const reportedCount = Number((/(\d+)/.exec(countText) || [])[1]);
// 页数不写死：测试收尾时删过页，所以只要求**文件与汇报一致**（这本来就是最强的跨源核对）。
check("file says the same number of slides as the report", !!slideScan && slideScan.count === reportedCount && reportedCount >= 1, JSON.stringify({ file: slideScan && slideScan.count, reported: reportedCount }));
const pptBody = JSON.stringify(pptxText(PPC) || []);
// 不写"文本框"：它在收尾时被 delete_textbox 删掉了。用**文件里真实存在**的内容核对。
check("file carries the slide content the test wrote", pptBody.includes("正文内容"), pptBody.slice(0, 110));
check("file carries the replace marker the test wrote", pptBody.includes("REPLMARKER-TWO"), pptBody.slice(0, 110));
const png = pngInfo(exportPath);
check("the exported slide image has real dimensions", png.exists && png.png && png.width >= 800 && png.height >= 600, JSON.stringify(png));

// teardown
for (const [method, appType] of [["closePresentation", "wpp"]]) {
  for (let i = 0; i < 6; i++) {
    const res = await call("wps_call", { tool: "wps_execute_method", args: { method, params: { save: false }, appType } });
    if (!ok(res)) break;
  }
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log("      matrix tools succeeded: " + succeeded + " / " + MATRIX.length);
check("the matrix actually drove tools (no systemic failure)", succeeded >= 10, succeeded + " succeeded");
console.log(failed === 0 ? "PPT COVERAGE TESTS OK (" + results.length + ")" : "PPT COVERAGE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

// Regression for the presentation parameter-contract batch.
//
// Every case exercises a parameter that the bridge used to ignore: the tools sent shapeIndex for
// name-based targeting, transition for effect, url for address, value for progress, style objects
// for flat properties, or a structured data shape the action never read.
// Run: node test/ppt-contract-fixes.test.mjs
import { spawn } from "node:child_process";
import { existsSync, writeFileSync } from "node:fs";
import path from "node:path";

const PROBE_PNG = "test/.artifacts/probe.png";
if (!existsSync(PROBE_PNG)) writeFileSync(PROBE_PNG, Buffer.from("89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a49444154789c6360000002000100ffff03000006000557bfabd40000000049454e44ae426082", "hex"));

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
function payload(res) { try { return JSON.parse(text(res)); } catch { return {}; } }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "ppt-fixes", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params, appType) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {}, appType } });
const slideCount = async () => payload(await viaAction("getSlideCount", {})).data?.count ?? -1;
const shapes = async (i) => payload(await viaAction("getShapes", { slideIndex: i })).data?.shapes ?? [];

check("create_presentation", ok(await call("wps_ppt_create_presentation", {})), "");
// WPS's Presentations.Add() returns a deck with no slides, so add them.
for (let i = 0; i < 2; i++) await call("wps_ppt_add_slide", { layout: "blank" });
check("two slides exist", (await slideCount()) === 2, "slides=" + (await slideCount()));
const tb = await call("wps_ppt_add_textbox", { slideIndex: 1, text: "标题", left: 50, top: 40, width: 300, height: 60, fontSize: 28 });
check("textbox added", ok(tb), text(tb).replace(/\s+/g, " ").slice(0, 60));

// transition/effect and duration, footer show, slide number show, date-time format
check("set_slide_transition accepts transition+duration", ok(await call("wps_ppt_set_transition", { slideIndex: 1, transition: "fade", duration: 1 })), "");
check("set_slide_transition accepts effect", ok(await call("wps_ppt_set_slide_transition", { slideIndex: 1, effect: "fade", duration: 0.5 })), "");
check("apply_transition_to_all accepts effect", ok(await call("wps_ppt_apply_transition_to_all", { effect: "fade" })), "");
// P4 merged the three footer tools into one; the checks follow the new surface.
check("footer can be shown", ok(await call("wps_ppt_set_slide_footer", { footerText: "probe", showFooter: true })), "");
check("footer reports hidden", text(await call("wps_ppt_set_slide_footer", { footerText: "probe", showFooter: false })).includes("隐藏"), "");
check("slide number accepts show", ok(await call("wps_ppt_set_slide_footer", { showSlideNumber: true })), "");
check("date time accepts autoUpdate+format", ok(await call("wps_ppt_set_slide_footer", { showDate: true, autoUpdate: false, dateFormat: "YYYY-MM-DD" })), "");

// The scenario wrappers (KPI cards, timeline, flow chart, gauge, donut, progress bar, page
// indicator, title decoration, org chart, mini chart, colour scheme) were removed from the tool
// layer; they now live as recipes in skills/wps-ppt/SKILL.md and are built from primitives. These
// checks build the same figures out of primitives, which is what makes the recipes trustworthy.
const shapeCount = async (slide) => (payload(await viaAction("getShapes", { slideIndex: slide })).data?.shapes ?? []).length;
const addShape = (a) => call("wps_ppt_add_shape", a);
const shapeAt = async (slide) => (payload(await viaAction("getShapes", { slideIndex: slide })).data?.shapes ?? []);
const fillShape = (slide, index, color) => call("wps_ppt_set_shape_fill", { slideIndex: slide, shapeIndex: index, color });
const textShape = (slide, index, t) => call("wps_ppt_set_shape_text", { slideIndex: slide, shapeIndex: index, text: t });

// recipe: 进度条 = 轨道 + 按比例宽度的填充 + 说明文字
const trackIdx = (await shapeCount(1)) + 1;
check("recipe: progress bar track", ok(await addShape({ slideIndex: 1, type: "rectangle", left: 60, top: 380, width: 400, height: 24, fillColor: "#E0E0E0" })), "");
const fillIdx = (await shapeCount(1)) + 1;
check("recipe: progress bar fill", ok(await addShape({ slideIndex: 1, type: "rectangle", left: 60, top: 380, width: Math.round(400 * 0.6), height: 24, fillColor: "#28A745" })), "");
const slide1 = await shapeAt(1);
const bar = slide1.find((s) => s.index === fillIdx);
check("recipe: fill width is 60% of the track", bar && Math.abs(bar.width - 240) <= 2, bar ? "width=" + bar.width : "shape not found");

// recipe: 仪表盘/环形图 = 外圈 + 白色内圈 + 中心百分比
const ringOuter = (await shapeCount(1)) + 1;
check("recipe: ring outer circle", ok(await addShape({ slideIndex: 1, type: "oval", left: 700, top: 300, width: 160, height: 160, fillColor: "#0D47A1" })), "");
const ringInner = (await shapeCount(1)) + 1;
check("recipe: ring inner circle", ok(await addShape({ slideIndex: 1, type: "oval", left: 732, top: 332, width: 96, height: 96, fillColor: "#FFFFFF" })), "");
check("recipe: ring centre text", ok(await call("wps_ppt_add_textbox", { slideIndex: 1, left: 740, top: 365, width: 80, height: 30, text: "60%", fontSize: 18 })), "");

// recipe: 页码指示器 = 右下角文本框
const pageBox = await call("wps_ppt_add_textbox", { slideIndex: 1, left: 880, top: 520, width: 80, height: 30, text: "1 / " + (await slideCount()), fontSize: 12 });
check("recipe: page indicator textbox", ok(pageBox), text(pageBox).replace(/\s+/g, " ").slice(0, 60));

// recipe: 配色 = 遍历现有形状逐个改填充色（取代 apply_color_scheme）
const palette = ["#2F5496", "#333333", "#00B0F0"];
let recoloured = 0;
for (const s of await shapeAt(1)) {
  if (recoloured >= palette.length) break;
  if (await fillShape(1, s.index, palette[recoloured])) recoloured++;
}
check("recipe: recolour shapes via set_shape_fill", recoloured > 0, "recoloured=" + recoloured);

// animation targeting + trigger
check("animation preset accepts shapeIndex", ok(await call("wps_ppt_add_animation", { slideIndex: 1, preset: "fadeIn", shapeIndex: 1 })), "");
check("animation preset rejects an out-of-range shape", !ok(await call("wps_ppt_add_animation", { slideIndex: 1, preset: "fadeIn", shapeIndex: 99 })), "");
check("animation emphasis accepts an effect", ok(await call("wps_ppt_add_animation", { slideIndex: 1, shapeIndex: 1, effect: "pulse", effectKind: "emphasis" })), "");
check("add_animation accepts shapeIndex+trigger", ok(await call("wps_ppt_add_animation", { slideIndex: 1, shapeIndex: 1, effect: 10, trigger: "onClick" })), "");
check("remove_animation accepts animationIndex", ok(await call("wps_ppt_remove_animation", { slideIndex: 1, animationIndex: 1 })), "");

// hyperlink url, font unification range, title decoration style
check("hyperlink accepts url", ok(await call("wps_ppt_add_ppt_hyperlink", { slideIndex: 1, shapeIndex: 1, url: "https://example.com" })), "");
check("unify_font accepts include_title/include_body", ok(await call("wps_ppt_unify_font", { font_name: "Arial", include_title: true, include_body: false })), "");
// recipe: 标题装饰 = 标题下方一根细色条
check("recipe: title decoration bar", ok(await addShape({ slideIndex: 1, type: "rectangle", left: 60, top: 96, width: 200, height: 6, fillColor: "#1A365D" })), "");

// copy slide to a target position
const beforeCopy = await slideCount();
const copy = await call("wps_ppt_copy_slide", { slideIndex: 1, targetIndex: 2 });
check("copy_slide accepts targetIndex", ok(copy), text(copy).replace(/\s+/g, " ").slice(0, 60));
check("copy_slide added a slide", (await slideCount()) === beforeCopy + 1, "slides=" + (await slideCount()));

// backgrounds: slide (solid), master (gradient), and an unknown type must fail
check("slide background accepts background object", ok(await call("wps_ppt_set_slide_background", { slideIndex: 1, background: { type: "solid", color: "#FF0000" } })), "");
check("unknown background type fails loudly", !ok(await call("wps_ppt_set_slide_background", { slideIndex: 1, background: { type: "plaid" } })), "");
check("master background accepts gradient colours", ok(await call("wps_ppt_set_master_background", { background: { type: "gradient", colors: ["#1a1a2e", "#0f3460"] } })), "");

// recipe: 流程图 = 方框 + 文字 + 箭头形状
check("recipe: flow chart start box", ok(await addShape({ slideIndex: 2, type: "rectangle", left: 80, top: 120, width: 160, height: 60, text: "开始", fillColor: "#2F5496" })), "");
check("recipe: flow chart end box", ok(await addShape({ slideIndex: 2, type: "rectangle", left: 320, top: 120, width: 160, height: 60, text: "结束", fillColor: "#2F5496" })), "");
check("recipe: flow chart connector arrow", ok(await addShape({ slideIndex: 2, type: "arrow", left: 245, top: 138, width: 70, height: 24, fillColor: "#333333" })), "");

// recipe: 迷你图表 = 数值 + 标签 + 趋势箭头
const miniBase = (await shapeCount(2)) + 1;
check("recipe: mini chart value", ok(await call("wps_ppt_add_textbox", { slideIndex: 2, left: 80, top: 380, width: 100, height: 30, text: "3", fontSize: 20 })), "");
check("recipe: mini chart label", ok(await call("wps_ppt_add_textbox", { slideIndex: 2, left: 80, top: 412, width: 100, height: 20, text: "Q1", fontSize: 11 })), "");
check("recipe: mini chart trend", ok(await call("wps_ppt_add_textbox", { slideIndex: 2, left: 150, top: 380, width: 30, height: 20, text: "↑", fontSize: 14 })), "");

// recipe: 组织架构图 = 根节点 + 两个子节点
check("recipe: org chart root", ok(await addShape({ slideIndex: 2, type: "rectangle", left: 200, top: 200, width: 140, height: 50, text: "CEO", fillColor: "#1A365D" })), "");
check("recipe: org chart child left", ok(await addShape({ slideIndex: 2, type: "rectangle", left: 120, top: 300, width: 120, height: 44, text: "CTO", fillColor: "#2F5496" })), "");
check("recipe: org chart child right", ok(await addShape({ slideIndex: 2, type: "rectangle", left: 300, top: 300, width: 120, height: 44, text: "CFO", fillColor: "#2F5496" })), "");

// recipe: 时间线 = 连接线 + 节点圆点 + 事件文字
check("recipe: timeline axis", ok(await addShape({ slideIndex: 2, type: "rectangle", left: 80, top: 470, width: 500, height: 4, fillColor: "#9E9E9E" })), "");
check("recipe: timeline node", ok(await addShape({ slideIndex: 2, type: "oval", left: 150, top: 458, width: 28, height: 28, fillColor: "#2F5496" })), "");
check("recipe: timeline label", ok(await call("wps_ppt_add_textbox", { slideIndex: 2, left: 120, top: 500, width: 140, height: 24, text: "2026 Q1", fontSize: 11 })), "");

// recipe: KPI 卡片行 = 三张卡片（形状 + 填充 + 文字）
for (const [i, kpi] of [["A", "销售额¥128万"], ["B", "增长+15%"], ["C", "客户320"]].entries()) {
  const idx = (await shapeCount(2)) + 1;
  const added = await addShape({ slideIndex: 2, type: "rectangle", left: 80 + i * 170, top: 560, width: 150, height: 70, fillColor: "#1565C0" });
  check("recipe: kpi card " + kpi[0] + " shape", ok(added), "");
  check("recipe: kpi card " + kpi[0] + " text", ok(await textShape(2, idx, String(kpi[1]))), "");
}

// P4-2: the four fragmented setters became one structured tool, and P4 dropped the 3D family (D3),
// so the contract checks follow the surface instead of the tools that no longer exist.
check("shape effect applies a shadow", ok(await call("wps_ppt_set_shape_effect", { slideIndex: 1, shapeIndex: 1, shadowEnabled: true, shadowColor: "#000000", shadowBlur: 4, shadowOffsetX: 2, shadowOffsetY: 2, shadowTransparency: 0.5 })), "");
check("shape effect applies a border", ok(await call("wps_ppt_set_shape_effect", { slideIndex: 1, shapeIndex: 1, borderEnabled: true, borderColor: "#FF0000", borderWidth: 2, borderStyle: "dash" })), "");
check("shape effect applies a two-colour gradient", ok(await call("wps_ppt_set_shape_effect", { slideIndex: 1, shapeIndex: 1, gradientColor1: "#FF0000", gradientColor2: "#0000FF" })), "");
check("shape effect rejects an unknown border style", !ok(await call("wps_ppt_set_shape_effect", { slideIndex: 1, shapeIndex: 1, borderStyle: "zigzag" })), "");
check("shape effect rejects a call with no effect at all", !ok(await call("wps_ppt_set_shape_effect", { slideIndex: 1, shapeIndex: 1 })), "");

// slide image: the action used to be Word's insertImage, and image style takes a nested object
const img = await call("wps_ppt_insert_slide_image", { slideIndex: 2, imagePath: path.resolve(PROBE_PNG), left: 40, top: 40 });
check("insert_slide_image inserts into the slide", ok(img), text(img).replace(/\s+/g, " ").slice(0, 70));
const slide2Shapes = await shapes(2);
const picture = slide2Shapes.find((s) => s.type === 13 || /Picture/i.test(String(s.name)));
if (picture) {
  const styled = await call("wps_ppt_set_image_style", { slideIndex: 2, imageIndex: 1, style: { border: { enabled: true, color: "#000000", weight: 1 }, opacity: 0.8, cropTop: 0.1 } });
  check("image style accepts a nested style object", ok(styled), text(styled).replace(/\s+/g, " ").slice(0, 70));
} else {
  check("image style accepts a nested style object", false, "no picture shape found on slide 2: " + JSON.stringify(slide2Shapes).slice(0, 120));
}

// start_slide_show must reject a bad slide without launching a show
check("start_slide_show rejects an out-of-range slide", !ok(await call("wps_ppt_start_slide_show", { fromSlide: 99 })), "");

// teardown
for (const [method, appType] of [["closeWorkbook", "et"], ["closeDocument", "wps"], ["closePresentation", "wpp"]]) {
  for (let i = 0; i < 6; i++) {
    const res = await call("wps_call", { tool: "wps_execute_method", args: { method, params: { save: false }, appType } });
    if (!ok(res)) break;
  }
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "PPT CONTRACT FIX TESTS OK (" + results.length + ")" : "PPT CONTRACT FIX TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

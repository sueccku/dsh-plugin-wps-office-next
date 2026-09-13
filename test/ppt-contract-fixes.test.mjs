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
check("footer can be shown", ok(await call("wps_ppt_set_ppt_footer", { text: "probe", show: true })), "");
check("footer reports hidden", text(await call("wps_ppt_set_ppt_footer", { text: "probe", show: false })).includes("隐藏"), "");
check("slide number accepts show", ok(await call("wps_ppt_set_slide_number", { show: true })), "");
check("date time accepts autoUpdate+format", ok(await call("wps_ppt_set_ppt_date_time", { show: true, autoUpdate: false, format: "YYYY-MM-DD", text: "2026-01-01" })), "");

// gauge max, progress value, page indicator position
check("gauge accepts value+max", ok(await call("wps_ppt_create_gauge", { slideIndex: 1, value: 75, max: 150 })), "");
check("progress bar accepts value", ok(await call("wps_ppt_create_progress_bar", { slideIndex: 1, value: 60, label: "进度" })), "");
check("page indicator accepts position", ok(await call("wps_ppt_add_page_indicator", { slideIndex: 1, position: "bottom-left" })), "");
check("unknown indicator position fails loudly", !ok(await call("wps_ppt_add_page_indicator", { slideIndex: 1, position: "sideways" })), "");

// colour scheme slide range
check("color scheme accepts slideIndex", ok(await call("wps_ppt_apply_color_scheme", { slideIndex: 1, scheme: "tech" })), "");
check("color scheme rejects an out-of-range slide", !ok(await call("wps_ppt_apply_color_scheme", { slideIndex: 99, scheme: "tech" })), "");

// animation targeting + trigger
check("animation preset accepts shapeIndex", ok(await call("wps_ppt_add_animation_preset", { slideIndex: 1, preset: "fadeIn", shapeIndex: 1 })), "");
check("animation preset rejects an out-of-range shape", !ok(await call("wps_ppt_add_animation_preset", { slideIndex: 1, preset: "fadeIn", shapeIndex: 99 })), "");
check("add_animation accepts shapeIndex+trigger", ok(await call("wps_ppt_add_animation", { slideIndex: 1, shapeIndex: 1, effect: 10, trigger: "onClick" })), "");
check("remove_animation accepts animationIndex", ok(await call("wps_ppt_remove_animation", { slideIndex: 1, animationIndex: 1 })), "");

// hyperlink url, font unification range, title decoration style
check("hyperlink accepts url", ok(await call("wps_ppt_add_ppt_hyperlink", { slideIndex: 1, shapeIndex: 1, url: "https://example.com" })), "");
check("unify_font accepts include_title/include_body", ok(await call("wps_ppt_unify_font", { font_name: "Arial", include_title: true, include_body: false })), "");
check("title decoration accepts style", ok(await call("wps_ppt_add_title_decoration", { slideIndex: 1, style: "sidebar" })), "");
check("unknown title decoration fails loudly", !ok(await call("wps_ppt_add_title_decoration", { slideIndex: 1, style: "nope" })), "");

// copy slide to a target position
const beforeCopy = await slideCount();
const copy = await call("wps_ppt_copy_slide", { slideIndex: 1, targetIndex: 2 });
check("copy_slide accepts targetIndex", ok(copy), text(copy).replace(/\s+/g, " ").slice(0, 60));
check("copy_slide added a slide", (await slideCount()) === beforeCopy + 1, "slides=" + (await slideCount()));

// backgrounds: slide (solid), master (gradient), and an unknown type must fail
check("slide background accepts background object", ok(await call("wps_ppt_set_slide_background", { slideIndex: 1, background: { type: "solid", color: "#FF0000" } })), "");
check("unknown background type fails loudly", !ok(await call("wps_ppt_set_slide_background", { slideIndex: 1, background: { type: "plaid" } })), "");
check("master background accepts gradient colours", ok(await call("wps_ppt_set_master_background", { background: { type: "gradient", colors: ["#1a1a2e", "#0f3460"] } })), "");

// structured scenario data the action used to ignore
check("flow chart accepts nodes+connections", ok(await call("wps_ppt_create_flow_chart", { slideIndex: 2, nodes: [{ id: "a", text: "开始", type: "start" }, { id: "b", text: "结束" }], connections: [{ from: "a", to: "b" }] })), "");
check("mini charts accept data", ok(await call("wps_ppt_create_mini_charts", { slideIndex: 2, data: [{ label: "Q1", values: [1, 2, 3] }, { label: "Q2", values: [5, 4, 2] }] })), "");
check("org chart accepts a tree", ok(await call("wps_ppt_create_org_chart", { slideIndex: 2, data: { name: "CEO", children: [{ name: "CTO" }, { name: "CFO" }] } })), "");
// A multi-slice donut is not supported (WPS's pie shape has one adjustment), so this exercises the
// single share it does draw, with the centre text the caller asked for.
const donut = await call("wps_ppt_create_donut_chart", { slideIndex: 1, value: 0.6, title: "占比", centerText: "60%" });
check("donut chart accepts value+title+centerText", ok(donut), text(donut).replace(/\s+/g, " ").slice(0, 80));

// nested style objects
check("shape shadow accepts a shadow object", ok(await call("wps_ppt_set_shape_shadow", { slideIndex: 1, shapeIndex: 1, shadow: { enabled: true, color: "#000000", blur: 4, offsetX: 2, offsetY: 2, opacity: 0.5 } })), "");
check("shape border accepts a border object", ok(await call("wps_ppt_set_shape_border", { slideIndex: 1, shapeIndex: 1, border: { enabled: true, color: "#FF0000", weight: 2, style: "dash" } })), "");
check("shape gradient accepts stops", ok(await call("wps_ppt_set_shape_gradient", { slideIndex: 1, shapeIndex: 1, gradient: { stops: [{ color: "#FF0000", position: 0 }, { color: "#0000FF", position: 1 }] } })), "");
check("shape gradient rejects the unsupported angle", !ok(await call("wps_ppt_set_shape_gradient", { slideIndex: 1, shapeIndex: 1, gradient: { angle: 45, stops: [{ color: "#FF0000" }, { color: "#0000FF" }] } })), "");
check("unknown border style fails loudly", !ok(await call("wps_ppt_set_shape_border", { slideIndex: 1, shapeIndex: 1, border: { style: "zigzag" } })), "");
check("3d rotation accepts a rotation object", ok(await call("wps_ppt_set_3d_rotation", { slideIndex: 1, shapeIndex: 1, rotation: { rotationX: 20, rotationY: 30 } })), "");

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

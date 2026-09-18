// S4 coverage: drive the previously-untested PPT tools against a real WPS deck.
//
// The plan's S4 acceptance is "success:true, or a clear and expected business error - never a
// timeout". So every entry is called with best-effort arguments and MUST return inside the budget;
// entries marked "ok" additionally must succeed. Coverage itself is counted by the S4 ratchet in
// test/spec-reproduction.test.mjs and reported by scripts/smoke-tools.mjs.
// Run: node test/ppt-coverage.test.mjs
import { spawn } from "node:child_process";
import { writeFileSync } from "node:fs";
import { resolve as resolvePath } from "node:path";

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
await call("wps_ppt_insert_ppt_image", { slideIndex: 1, filePath: imgPath });
await call("wps_ppt_insert_ppt_chart", { slideIndex: 1, chartType: "column_clustered", title: "销量" });
await call("wps_ppt_add_animation", { slideIndex: 1, effect: "fadeIn", shapeIndex: 1 });

// ---- matrix -------------------------------------------------------------------------------
// [tool, args, expectation]  expectation: "ok" (must succeed) | "error" (must fail clearly) | "any"
const MATRIX = [
  // reads
  ["wps_ppt_get_slide_count", {}, "ok"],
  ["wps_ppt_get_slide_info", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_slide_master", {}, "any"],
  ["wps_ppt_get_slide_notes", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_slide_title", { slideIndex: 1 }, "any"],
  ["wps_ppt_get_shapes", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_textboxes", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_animations", { slideIndex: 1 }, "ok"],
  ["wps_ppt_get_table_cell", { slideIndex: 1, tableIndex: 1, row: 1, col: 1 }, "ok"],
  ["wps_ppt_find_ppt_text", { text: "标题" }, "any"],
  // slide settings
  ["wps_ppt_switch_slide", { slideIndex: 1 }, "ok"],
  ["wps_ppt_set_slide_layout", { slideIndex: 1, layout: "title_content" }, "any"],
  ["wps_ppt_set_slide_title", { slideIndex: 1, title: "覆盖标题" }, "any"],
  ["wps_ppt_set_slide_subtitle", { slideIndex: 1, subtitle: "副标题" }, "any"],
  ["wps_ppt_set_slide_content", { slideIndex: 1, content: "正文内容" }, "any"],
  ["wps_ppt_set_slide_notes", { slideIndex: 1, notes: "讲稿备注" }, "ok"],
  ["wps_ppt_set_slide_size", { width: 1280, height: 720 }, "ok"],
  ["wps_ppt_set_slide_theme", { theme: resolvePath("test/.artifacts/nope.thmx") }, "error"],
  ["wps_ppt_remove_slide_transition", { slideIndex: 1 }, "any"],
  // backgrounds
  ["wps_ppt_set_background_color", { slideIndex: 1, color: "#112233" }, "any"],
  ["wps_ppt_set_background_gradient", { slideIndex: 1, gradient: { color1: "#FFFFFF", color2: "#000000" } }, "any"],
  ["wps_ppt_set_background_image", { slideIndex: 1, imagePath: imgPath }, "any"],
  // shapes
  ["wps_ppt_set_shape_position", { slideIndex: 1, shapeIndex: 1, left: 60, top: 60, width: 120, height: 80 }, "ok"],
  ["wps_ppt_set_shape_style", { slideIndex: 1, shapeIndex: 1, fillColor: "#FF0000", lineColor: "#000000", lineWidth: 1 }, "ok"],
  ["wps_ppt_set_font_color", { slideIndex: 1, shapeIndex: 1, color: "#00FF00" }, "any"],
  ["wps_ppt_duplicate_shape", { slideIndex: 1, shapeIndex: 1 }, "ok"],
  ["wps_ppt_set_shape_z_order", { slideIndex: 1, shapeIndex: 1, order: "front" }, "any"],
  // text
  ["wps_ppt_set_textbox_text", { slideIndex: 1, textboxIndex: 1, text: "新文本" }, "ok"],
  ["wps_ppt_set_textbox_style", { slideIndex: 1, textboxIndex: 1, style: { fontSize: 20 } }, "any"],
  ["wps_ppt_replace_ppt_text", { find: "新文本", replace: "替换后" }, "any"],
  // table
  ["wps_ppt_set_table_cell", { slideIndex: 1, tableIndex: 1, row: 1, col: 1, text: "X" }, "ok"],
  // images
  ["wps_ppt_replace_ppt_image", { slideIndex: 1, filePath: imgPath }, "any"],
  // charts
  ["wps_ppt_set_ppt_chart_data", { slideIndex: 1, chartIndex: 1, data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] } }, "any"],
  ["wps_ppt_set_ppt_chart_style", { slideIndex: 1, chartIndex: 1, style: "style1" }, "any"],
  // animations
  ["wps_ppt_set_animation_order", { slideIndex: 1, animationIndex: 1, newOrder: 1 }, "any"],
  // grouping / distribution (needs several shapes)
  ["wps_ppt_group_shapes", { slideIndex: 1, shapeIndices: [1, 2] }, "any"],
  ["wps_ppt_distribute_shapes", { slideIndex: 1, shapeIndices: [1, 2, 3], direction: "horizontal" }, "any"],
  // slide ops
  ["wps_ppt_move_slide", { fromIndex: 1, toIndex: 2 }, "any"],
  ["wps_ppt_switch_presentation", { name: "演示文稿1" }, "any"],
  ["wps_ppt_set_active_target", { clear: true }, "any"],
  // hyperlink
  ["wps_ppt_remove_ppt_hyperlink", { slideIndex: 1, shapeIndex: 1 }, "any"],
  // master / beautify / export / files
  ["wps_ppt_add_master_element", { element: "slideNumber" }, "any"],
  ["wps_ppt_beautify", {}, "any"],
  ["wps_ppt_export_slide_as_image", { slideIndex: 1, outputPath: exportPath }, "any"],
  ["wps_ppt_insert_slides_from_file", { filePath: resolvePath("test/.artifacts/nope.pptx") }, "error"],
  ["wps_ppt_open_presentation", { filePath: resolvePath("test/.artifacts/nope.pptx") }, "error"],
  // destructive last
  ["wps_ppt_delete_shape", { slideIndex: 1, shapeIndex: 1 }, "any"],
  ["wps_ppt_delete_textbox", { slideIndex: 1, textboxIndex: 1 }, "any"],
  ["wps_ppt_delete_ppt_image", { slideIndex: 1, imageIndex: 1 }, "any"],
  ["wps_ppt_delete_slide", { slideIndex: 2 }, "any"],
];

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
  else check(label + " returns (no hang)", true, ms + "ms " + (good ? "ok" : "business error"));
}

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

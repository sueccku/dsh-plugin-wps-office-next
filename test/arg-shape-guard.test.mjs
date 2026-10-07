// P2: the tool registry rejects arguments whose *shape* cannot work (an array or object where a scalar
// is declared, or the other way round) before the handler ever reaches WPS. Scalar-to-scalar stays
// permissive on purpose: the bridge and COM coerce 42 and "42" alike (spec declares value: string), so
// strict JSON typing would only reject calls that work today. Enum values stay the bridge's business
// (unknown paperSize / unknown borderStyle ... are asserted in the bridge tests).
// Run: node test/arg-shape-guard.test.mjs
import { spawn } from "node:child_process";

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
function shapeRejected(res) { return /的形状不对/.test(text(res)); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "shape", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });

// data is declared array: a JSON string is the classic LLM mistake and must fail clearly.
const strData = await call("wps_excel_write_range", { range: "A1", data: "[[1,2]]" });
check("array parameter rejects a JSON string", shapeRejected(strData), text(strData).replace(/\s+/g, " ").slice(0, 110));

// background is declared object: a bare string must not be passed to the bridge.
const strBg = await call("wps_ppt_set_slide_background", { slideIndex: 1, background: "solid" });
check("object parameter rejects a string", shapeRejected(strBg), text(strBg).replace(/\s+/g, " ").slice(0, 110));

// calls is declared array: an object must be rejected before wps_batch tries to iterate it.
const objCalls = await call("wps_batch", { calls: { tool: "wps_excel_read_range" } });
check("array parameter rejects an object", shapeRejected(objCalls), text(objCalls).replace(/\s+/g, " ").slice(0, 110));

// The guard must not fire on ordinary scalar calls: value: 42 is legal today.
const scalar = await call("wps_excel_set_cell_value", { sheet: "Sheet1", row: 10, col: 1, value: 42 });
check("scalar parameter is not shape-checked", !shapeRejected(scalar), text(scalar).replace(/\s+/g, " ").slice(0, 110));

// And a well-formed call is untouched by the guard (whether WPS answers or not).
const good = await call("wps_excel_read_range", { range: "A1" });
check("well-formed call passes the guard", !shapeRejected(good), text(good).replace(/\s+/g, " ").slice(0, 110));

// Missing required parameters keep their own wording from before the guard existed.
const missing = await call("wps_excel_read_range", {});
check("missing required parameter keeps its own error", /Missing required parameter/.test(text(missing)) && !shapeRejected(missing), text(missing).replace(/\s+/g, " ").slice(0, 110));

// ---- FIXES 99（W6-4）：未知参数**名**必须响亮拒绝，不能静默按默认值执行 ----
// 在这条路径上，工具注册层是**唯一还看得见原始实参**的地方：handler 用显式解构拼参数，
// 未知键在到达桥之前就被丢掉，桥的"未知键拒绝"永远看不见它。
// 实证（W6-4）：`wps_excel_text_to_columns { range, sep }` 曾经不报错、按默认逗号执行，
// 回报还写着「分隔符: ","」—— 调用方以为自己传的 sep 生效了。
// 这三个断言在注册层就被拒，**不会走到 WPS**（所以本文件仍然是纯本地测试）。
const unknownParam = await call("wps_excel_text_to_columns", { range: "A1:A2", sep: "," });
check("an unknown parameter name is rejected instead of silently dropped", /未知参数/.test(text(unknownParam)) && /sep/.test(text(unknownParam)), text(unknownParam).slice(0, 140));
check("the rejection lists the accepted parameter names", /delimiter/.test(text(unknownParam)), text(unknownParam).slice(0, 160));
const nearMissParam = await call("wps_excel_text_to_columns", { range: "A1:A2", delimiter1: "," });
check("a near-miss parameter name gets a suggestion", /是不是想传 delimiter/.test(text(nearMissParam)), text(nearMissParam).slice(0, 160));
const declaredParam = await call("wps_help", { tool: "wps_excel_text_to_columns" });
check("a declared parameter is not rejected by the guard", !/未知参数/.test(text(declaredParam)), text(declaredParam).slice(0, 110));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "ARG SHAPE TESTS OK (" + results.length + ")" : "ARG SHAPE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

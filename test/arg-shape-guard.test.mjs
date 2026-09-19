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

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "ARG SHAPE TESTS OK (" + results.length + ")" : "ARG SHAPE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

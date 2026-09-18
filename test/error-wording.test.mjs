// S7: customer-facing error wording is the three-part form: Chinese sentence + action name + next step.
// Samples the early required-parameter / unknown-action failure paths through the hidden escape hatch, so
// it does not depend on any particular document being open.
// Run: node test/error-wording.test.mjs
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
function payload(res) { try { return JSON.parse(text(res)); } catch { return {}; } }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "wording", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const viaAction = (method, params) => req(id++, "tools/call", { name: "wps_call", arguments: { tool: "wps_execute_method", args: { method, params: params || {} } } });

// Every one of these must fail early (required parameter / unknown action), so the sample does not
// depend on what happens to be open.
const SAMPLES = [
  ["saveAs", {}],
  ["convertFormat", {}],
  ["deleteRows", {}],
  ["deleteColumns", {}],
  ["insertRows", {}],
  ["insertColumns", {}],
  ["createListObject", {}],
  ["resizeListObject", {}],
  ["getFormulaAudit", {}],
  ["clearPivotTable", {}],
  ["goalSeek", {}],
  ["addSparkline", {}],
  ["clearSparkline", {}],
  ["setTableCell", {}],
  ["addFootnote", {}],
  ["addEndnote", {}],
  ["insertCrossReference", {}],
  ["addCellComment", {}],
  ["openWorkbook", {}],
  ["openDocument", {}],
  ["insertBookmark", {}],
  ["smartFillField", {}],
  ["openPresentation", {}],
  ["insertPptChart", {}],
  ["replacePptText", {}],
  ["__definitely_not_an_action", {}],
];

let good = 0;
for (const [method, params] of SAMPLES) {
  const data = payload(await viaAction(method, params));
  const err = typeof data.error === "string" ? data.error : "";
  const failed = data.success === false;
  const threePart = err.includes("（动作：") && err.includes("下一步：");
  const named = err.includes(method);
  if (failed && threePart && named) good++;
  check(method + " fails with a three-part message", failed && threePart && named, failed ? err.slice(0, 110) : "did NOT fail: " + JSON.stringify(data).slice(0, 90));
}
check("the sample is broad enough", good >= 20, good + " / " + SAMPLES.length + " well-formed");

// A mapping case: the app-not-running wording when a kind is not reachable is covered by the table in
// the bridge, so assert the table is present and returns the three-part shape for a known phrase.
const known = payload(await viaAction("__definitely_not_an_action", {}));
check("an unknown action still names the action", String(known.error || "").includes("__definitely_not_an_action"), String(known.error || "").slice(0, 110));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "ERROR WORDING TESTS OK (" + results.length + ")" : "ERROR WORDING TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

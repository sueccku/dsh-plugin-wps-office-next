// P2 / known defect C7: when the caller does not name the target, the bridge falls back to the "active"
// object. With more than one workbook/presentation open that target follows window focus, so a timeout
// retry or a second call can land on a different file than the caller meant. The bridge now says so in
// the result warnings instead of changing a file silently; naming the target keeps the result clean.
// The warning travels on the pass-through surface (wps_execute_method here), where bridge output is
// returned verbatim.
// This test closes every workbook/presentation it can reach, like the other WPS tests do.
// Run: node test/target-ambiguity.test.mjs
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
function ok(res) { return !!(res && res.result && !res.result.isError); }
function payload(res) { try { return JSON.parse(text(res)); } catch { return {}; } }
const hasWarn = (p, re) => Array.isArray(p.warnings) && p.warnings.some((w) => re.test(String(w)));

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "ambiguity", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {} } });
const raw = async (method, params) => payload(await viaAction(method, params));
async function closeAll(method, appType) {
  let closed = 0;
  for (let i = 0; i < 10; i++) {
    const res = await call("wps_call", { tool: "wps_execute_method", args: { method, params: { save: false }, appType } });
    if (!ok(res)) break;
    closed++;
  }
  return closed;
}

// --- Excel -----------------------------------------------------------------
await closeAll("closeWorkbook", "et");
const one = await raw("createWorkbook", {});
check("one workbook open", one.success === true, JSON.stringify(one).slice(0, 80));
const solo = await raw("getRangeData", { range: "A1" });
check("a single workbook raises no ambiguity warning", !hasWarn(solo, /个打开的工作簿/), JSON.stringify(solo.warnings || []).slice(0, 120));

const two = await raw("createWorkbook", {});
check("second workbook open", two.success === true, JSON.stringify(two).slice(0, 80));
const bare = await raw("getRangeData", { range: "A1" });
check("no sheet + two workbooks -> warning", hasWarn(bare, /个打开的工作簿/), JSON.stringify(bare.warnings || []).slice(0, 160));
check("the warning names the drift risk and the fix", hasWarn(bare, /活动工作表/), JSON.stringify(bare.warnings || []).slice(0, 160));
check("the warning also rides inside data", Array.isArray(bare.data && bare.data.warnings) && bare.data.warnings.length > 0, JSON.stringify((bare.data || {}).warnings || []).slice(0, 120));

const named = await raw("getRangeData", { range: "A1", sheet: "Sheet1" });
check("explicit sheet succeeds", named.success === true, JSON.stringify(named).slice(0, 90));
check("explicit sheet silences the ambiguity warning", !hasWarn(named, /个打开的工作簿/), JSON.stringify(named.warnings || []).slice(0, 120));
await closeAll("closeWorkbook", "et");

// --- PowerPoint ------------------------------------------------------------
await closeAll("closePresentation", "wpp");
const pOne = await raw("createPresentation", {});
check("one presentation open", pOne.success === true, JSON.stringify(pOne).slice(0, 80));
const pSolo = await raw("getActivePresentation", {});
check("a single presentation raises no ambiguity warning", !hasWarn(pSolo, /个打开的演示文稿/), JSON.stringify(pSolo.warnings || []).slice(0, 120));

const pTwo = await raw("createPresentation", {});
check("second presentation open", pTwo.success === true, JSON.stringify(pTwo).slice(0, 80));
const pBare = await raw("getActivePresentation", {});
check("no presentationName + two presentations -> warning", hasWarn(pBare, /个打开的演示文稿/), JSON.stringify(pBare.warnings || []).slice(0, 160));
const pName = "wps-plugin-target-probe";
const pNamed = await raw("getActivePresentation", { presentationName: pName });
check("an explicit presentationName raises no ambiguity warning", !hasWarn(pNamed, /个打开的演示文稿/) && pNamed.success !== true, JSON.stringify(pNamed).slice(0, 100));
await closeAll("closePresentation", "wpp");

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "TARGET AMBIGUITY TESTS OK (" + results.length + ")" : "TARGET AMBIGUITY TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

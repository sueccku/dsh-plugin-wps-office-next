// Closing documents must never raise a modal dialog.
//
// The original leak: the tools send "save" while the bridge read only "saveChanges", so save=false
// was ignored and Close(true) ran on a workbook that had never been saved. That opens a Save As
// dialog which nothing answers, the COM call blocks, and every later call fails with
// RPC_E_CALL_REJECTED while workbooks pile up.
//
// Run: node test/close-safety.test.mjs
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

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "close", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;

// A hung COM call is the failure mode under test, so every call is watched: if it does not come
// back, the MCP child is killed and the run fails loudly instead of blocking CI forever.
async function call(name, args, ms = 90000) {
  let timer;
  const race = new Promise((resolve) => { timer = setTimeout(() => resolve("__TIMEOUT__"), ms); });
  const res = await Promise.race([req(id++, "tools/call", { name, arguments: args }), race]);
  clearTimeout(timer);
  if (res === "__TIMEOUT__") {
    console.log("FAIL " + name + " HUNG for " + ms + "ms -> killing the MCP child");
    child.kill();
    process.exit(1);
  }
  return res;
}
const viaAction = (method, params) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {} } }, 60000);
const wbCount = async () => payload(await viaAction("getOpenWorkbooks")).data?.count ?? -1;
const presCount = async () => payload(await viaAction("getOpenPresentations")).data?.count ?? -1;

// ---------------- Excel ----------------
const before = await wbCount();
check("baseline workbook count is readable", before >= 0, "workbooks=" + before);

await call("wps_excel_create_workbook", {});
check("create_workbook adds one", (await wbCount()) === before + 1, "workbooks=" + (await wbCount()));

const c1 = await call("wps_excel_close_workbook", { save: false });
const t1 = text(c1);
check("close_workbook(save:false) succeeds", ok(c1), t1.replace(/\s+/g, " ").slice(0, 100));
check("close_workbook(save:false) reports 未保存", t1.includes("未保存"), t1.replace(/\s+/g, " ").slice(0, 100));
check("close_workbook(save:false) did not prompt and did not leak", (await wbCount()) === before, "workbooks=" + (await wbCount()));

// The regression proper: save defaults to true, and on a never-saved workbook that used to open
// a modal Save As dialog. It must instead close quietly and say so.
await call("wps_excel_create_workbook", {});
const c2 = await call("wps_excel_close_workbook", {});
const t2 = text(c2);
check("close_workbook(default save:true) on an unsaved workbook returns", ok(c2), t2.replace(/\s+/g, " ").slice(0, 110));
check("it warns that nothing was written to disk", t2.includes("never saved"), t2.replace(/\s+/g, " ").slice(0, 110));
check("no workbook leaked", (await wbCount()) === before, "workbooks=" + (await wbCount()));

// ---------------- PowerPoint ----------------
const pb = await presCount();
if (pb < 0) {
  console.log("SKIP presentation checks (no presentation host)");
} else {
  await call("wps_ppt_create_presentation", {});
  const p1 = await call("wps_ppt_close_presentation", { save: false });
  const pt1 = text(p1);
  check("close_presentation(save:false) succeeds", ok(p1), pt1.replace(/\s+/g, " ").slice(0, 100));
  check("close_presentation(save:false) reports 未保存", pt1.includes("未保存"), pt1.replace(/\s+/g, " ").slice(0, 100));
  check("close_presentation(save:false) did not leak", (await presCount()) === pb, "presentations=" + (await presCount()));

  await call("wps_ppt_create_presentation", {});
  const p2 = await call("wps_ppt_close_presentation", {});
  const pt2 = text(p2);
  check("close_presentation(default save:true) on an unsaved deck returns", ok(p2), pt2.replace(/\s+/g, " ").slice(0, 110));
  check("it warns that nothing was written to disk", pt2.includes("never saved"), pt2.replace(/\s+/g, " ").slice(0, 110));
  check("no presentation leaked", (await presCount()) === pb, "presentations=" + (await presCount()));
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "CLOSE SAFETY TESTS OK (" + results.length + ")" : "CLOSE SAFETY TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

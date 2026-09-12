// Worksheet operation group regression.
// Upstream tools spell the sheet parameter as name/oldName while the bridge only read "sheet",
// so delete/rename/copy/move acted on whatever sheet happened to be active, and switch always failed.
// position is 0-based, as the tool schemas state.
// Run: node test/sheet-ops.test.mjs
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

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "sheets", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {} } });
const state = async () => (payload(await viaAction("getSheetList", {})).data) || { sheets: [], activeSheet: "" };
const order = async () => (await state()).sheets.map((s) => (typeof s === "string" ? s : s.name));
const info = async (cell, sheet) => (payload(await viaAction("getCellInfo", { cell, sheet })).data) || {};

await viaAction("createWorkbook", {});
check("workbook created", (await order()).length >= 1, (await order()).join(", "));

const c0 = await call("wps_excel_create_sheet", { name: "Alpha", position: 0 });
check("create_sheet at position 0", ok(c0), text(c0).replace(/\s+/g, " ").slice(0, 90));
check("new sheet is first", (await order())[0] === "Alpha", (await order()).join(", "));

const c2 = await call("wps_excel_create_sheet", { name: "Beta", position: 2 });
check("create_sheet at position 2", ok(c2), text(c2).replace(/\s+/g, " ").slice(0, 90));
check("new sheet sits at index 2", (await order())[2] === "Beta", (await order()).join(", "));

// The schema promises "append at the end" when position is omitted.
const cDef = await call("wps_excel_create_sheet", { name: "Appended" });
check("omitted position appends at the end", ok(cDef) && (await order())[(await order()).length - 1] === "Appended", (await order()).join(", "));

const cNeg = await call("wps_excel_create_sheet", { name: "Nope", position: -3 });
check("negative position is rejected", !ok(cNeg), text(cNeg).slice(0, 70));

const r1 = await call("wps_excel_rename_sheet", { oldName: "Alpha", newName: "AlphaRenamed" });
check("rename_sheet renames the named sheet", ok(r1), text(r1).replace(/\s+/g, " ").slice(0, 80));
check("renamed in place, old name gone", (await order()).includes("AlphaRenamed") && !(await order()).includes("Alpha"), (await order()).join(", "));

const cp = await call("wps_excel_copy_sheet", { name: "AlphaRenamed", newName: "AlphaCopy", position: 0 });
check("copy_sheet at position 0", ok(cp), text(cp).replace(/\s+/g, " ").slice(0, 80));
check("copy is first and renamed", (await order())[0] === "AlphaCopy", (await order()).join(", "));
check("copy left the source intact", (await order()).includes("AlphaRenamed"), (await order()).join(", "));

const sw = await call("wps_excel_switch_sheet", { name: "Beta" });
check("switch_sheet activates the named sheet", ok(sw) && (await state()).activeSheet === "Beta", text(sw).replace(/\s+/g, " ").slice(0, 60));

const mv = await call("wps_excel_move_sheet", { name: "Beta", position: 0 });
check("move_sheet runs", ok(mv), text(mv).replace(/\s+/g, " ").slice(0, 80));
check("moved sheet is now first", (await order())[0] === "Beta", (await order()).join(", "));

// THE safety case: deleting one sheet must not touch a different (active) one
await call("wps_excel_create_sheet", { name: "Victim", position: 0 });
await call("wps_excel_switch_sheet", { name: "AlphaCopy" });
const del = await call("wps_excel_delete_sheet", { name: "Victim" });
check("delete_sheet runs", ok(del), text(del).replace(/\s+/g, " ").slice(0, 80));
check("DELETED THE NAMED SHEET ONLY", !(await order()).includes("Victim") && (await order()).includes("AlphaCopy") && (await state()).activeSheet === "AlphaCopy", (await order()).join(", "));

const delNone = await call("wps_excel_delete_sheet", {});
check("delete without a target fails loudly", !ok(delNone), text(delNone).slice(0, 70));

// set_number_format must honour the sheet parameter
await call("wps_excel_switch_sheet", { name: "AlphaCopy" });
const nf = await call("wps_excel_set_number_format", { range: "A1", format: "0.00%", sheet: "Beta" });
check("set_number_format with sheet", ok(nf), text(nf).replace(/\s+/g, " ").slice(0, 80));
check("format applied on the named sheet", String((await info("A1", "Beta")).numberFormat).includes("0.00%"), "beta=" + (await info("A1", "Beta")).numberFormat);
check("format did not leak to the active sheet", !String((await info("A1", "AlphaCopy")).numberFormat).includes("0.00%"), "active=" + (await info("A1", "AlphaCopy")).numberFormat);

// Close everything this test opened. The close actions are dialog-safe now (see
// test/close-safety.test.mjs), so a scratch run leaves no documents behind for the next one.
// Word has no close tool, so all three go through the wps_call facade.
for (const [method, appType] of [["closeWorkbook", "et"], ["closeDocument", "wps"], ["closePresentation", "wpp"]]) {
  for (let i = 0; i < 6; i++) {
    const res = await call("wps_call", { tool: "wps_execute_method", args: { method, params: { save: false }, appType } });
    if (!ok(res)) break;
  }
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "SHEET OPS TESTS OK (" + results.length + ")" : "SHEET OPS TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

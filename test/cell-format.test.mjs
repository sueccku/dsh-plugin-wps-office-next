// Excel cell formatting regression: the whole format object used to be discarded by the bridge.
// Run: node test/cell-format.test.mjs
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

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "fmt", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {} } });
const info = async (cell) => payload(await viaAction("getCellInfo", { cell })).data || {};

await viaAction("createWorkbook", {});
await call("wps_excel_write_range", { range: "A1:C2", data: [["a", "b", "c"], ["d", "e", "f"]] });

// nested format object, every supported property at once
const nested = {
  bold: true, italic: true, fontSize: 14, fontName: "Arial", fontColor: "#FF0000",
  bgColor: "#FFFF00", horizontalAlignment: "center", verticalAlignment: "center",
  wrapText: true, numberFormat: "0.00%"
};
const r1 = await call("wps_excel_set_cell_format", { range: "A1", format: nested });
check("nested format accepted", ok(r1), text(r1).slice(0, 120));
const a1 = await info("A1");
check("bold applied", a1.font && a1.font.bold === true, JSON.stringify(a1.font));
check("italic applied", a1.font && a1.font.italic === true, JSON.stringify(a1.font));
check("fontSize applied", a1.font && Number(a1.font.size) === 14, JSON.stringify(a1.font));
check("fontName applied", a1.font && a1.font.name === "Arial", JSON.stringify(a1.font));
check("fontColor applied", a1.font && Number(a1.font.color) === 255, "color=" + (a1.font && a1.font.color));
check("bgColor applied", Number(a1.backgroundColor) === 65535, "bg=" + a1.backgroundColor);
check("horizontalAlignment applied", Number(a1.horizontalAlignment) === -4108, "h=" + a1.horizontalAlignment);
check("verticalAlignment applied", Number(a1.verticalAlignment) === -4108, "v=" + a1.verticalAlignment);
check("wrapText applied", a1.wrapText === true, "wrap=" + a1.wrapText);
check("numberFormat applied", String(a1.numberFormat).includes("0.00%"), "nf=" + a1.numberFormat);

// flat properties must work too (they are declared in the schema)
const r2 = await call("wps_excel_set_cell_format", { range: "B2", format: {}, bold: true, fontSize: 20 });
check("flat properties accepted", ok(r2), text(r2).slice(0, 120));
const b2 = await info("B2");
check("flat bold applied", b2.font && b2.font.bold === true, JSON.stringify(b2.font));
check("flat fontSize applied", b2.font && Number(b2.font.size) === 20, JSON.stringify(b2.font));

// formatting must not leak to neighbouring cells
const a2 = await info("A2");
check("no bleed to A2", !(a2.font && a2.font.bold === true) && Number(a2.font && a2.font.size) !== 14, JSON.stringify(a2.font));

// an empty format must fail loudly, not report a fake success
const r3 = await call("wps_excel_set_cell_format", { range: "C1", format: {} });
check("empty format fails loudly", !ok(r3) && text(r3).includes("no supported format property"), text(r3).slice(0, 110));

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
console.log(failed === 0 ? "CELL FORMAT TESTS OK (" + results.length + ")" : "CELL FORMAT TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

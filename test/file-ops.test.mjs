// Regression for the file-path and conversion batch.
//
// The open/save/convert tools all sent "filePath" (and often two more aliases) while the bridge read
// "path"/"outputPath", so the target path was dropped and the operation either failed or acted on the
// wrong document. convertToPDF also picked the first running application, so with Excel open it
// exported the workbook no matter which document the caller meant.
// Run: node test/file-ops.test.mjs
import { spawn } from "node:child_process";
import { existsSync, mkdirSync, rmSync } from "node:fs";
import path from "node:path";

const OUT = path.resolve("test/.artifacts/file-ops");
mkdirSync(OUT, { recursive: true });
const XLSX = path.join(OUT, "probe-workbook.xlsx");
const PDF = path.join(OUT, "probe-export.pdf");
for (const f of [XLSX, PDF]) { if (existsSync(f)) rmSync(f); }

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

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "file-ops", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params, appType) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {}, appType } });

await viaAction("createWorkbook", {});
await call("wps_excel_write_range", { range: "A1:B2", data: [["k", "v"], ["1", "2"]] });

// save_as used to send filePath+outputPath and rely on one being read; the bridge reads path.
const saved = await call("wps_common_save_as", { filePath: XLSX });
check("save_as accepts filePath", ok(saved), text(saved).replace(/\s+/g, " ").slice(0, 90));
check("save_as actually wrote the file to the requested path", existsSync(XLSX), XLSX);

// open_workbook had the same mismatch.
const before = payload(await viaAction("getOpenWorkbooks", {})).data?.count ?? -1;
const opened = await call("wps_excel_open_workbook", { filePath: XLSX });
check("open_workbook accepts filePath", ok(opened), text(opened).replace(/\s+/g, " ").slice(0, 90));
const after = payload(await viaAction("getOpenWorkbooks", {})).data?.count ?? -1;
check("open_workbook really opened the requested file", after >= before, "before=" + before + " after=" + after);

// A missing file must be reported, not silently ignored.
const missing = await call("wps_excel_open_workbook", { filePath: path.join(OUT, "does-not-exist.xlsx") });
check("open_workbook reports a missing file instead of succeeding", !ok(missing), text(missing).replace(/\s+/g, " ").slice(0, 90));

// convert_to_pdf picked Excel whenever Excel was running, whatever the caller meant.
const pdf = await call("wps_convert_to_pdf", { outputPath: PDF, app_type: "excel" });
check("convert_to_pdf accepts app_type", ok(pdf), text(pdf).replace(/\s+/g, " ").slice(0, 90));
check("convert_to_pdf wrote the PDF to the requested path", existsSync(PDF), PDF);
const pdfExcelOnly = await call("wps_convert_to_pdf", { outputPath: PDF, app_type: "ppt" });
check("app_type=ppt does not silently export Excel instead", !ok(pdfExcelOnly), text(pdfExcelOnly).replace(/\s+/g, " ").slice(0, 90));

// Word header/footer now honour the section the tool passes.
await viaAction("createDocument", {}, "wps");
await call("wps_word_insert_text", { text: "body", position: "start" });
const hdr = await call("wps_word_insert_header", { text: "probe header", section: 1 });
check("insert_header accepts section", ok(hdr), text(hdr).replace(/\s+/g, " ").slice(0, 80));
const ftr = await call("wps_word_insert_footer", { text: "probe footer", section: 1 });
check("insert_footer accepts section", ok(ftr), text(ftr).replace(/\s+/g, " ").slice(0, 80));
const badSection = await call("wps_word_insert_header", { text: "nope", section: 99 });
check("a non-existent section fails loudly", !ok(badSection), text(badSection).replace(/\s+/g, " ").slice(0, 90));

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
console.log(failed === 0 ? "FILE OPS TESTS OK (" + results.length + ")" : "FILE OPS TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

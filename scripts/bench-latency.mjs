import { spawn } from "node:child_process";

const entry = process.argv[2] || "mcp/dist/index.js";
const child = spawn(process.execPath, [entry], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
let stderr = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((resolve) => { pending.set(id, resolve); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", (d) => { stderr += d.toString(); });

function stats(values) {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  const median = sorted.length % 2 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
  return { n: values.length, min: sorted[0], median, max: sorted[sorted.length - 1] };
}

const startup = Date.now();
await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "bench", version: "1.0.0" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
await req(2, "tools/list", {});
const readyMs = Date.now() - startup;

let id = 10;
async function timeCall(name, args) {
  const t = Date.now();
  const res = await req(id++, "tools/call", { name, arguments: args });
  const ms = Date.now() - t;
  const text = res && res.result && res.result.content && res.result.content[0] ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {});
  return { ms, ok: !!(res && res.result && !res.result.isError), text };
}

const report = { startupMs: readyMs, cold: {}, warm: {} };

const coldPing = await timeCall("wps_common_ping", {});
report.cold.ping = coldPing.ms;
const coldWire = await timeCall("wps_common_wire_check", {});
report.cold.wireCheck = coldWire.ms;
const coldApp = await timeCall("wps_common_get_app_info", {});
report.cold.getAppInfo = coldApp.ms;

const pingSeries = [];
for (let i = 0; i < 7; i++) { const r = await timeCall("wps_common_ping", {}); pingSeries.push(r.ms); }
report.warm.ping = stats(pingSeries);

const wireSeries = [];
for (let i = 0; i < 5; i++) { const r = await timeCall("wps_common_wire_check", {}); wireSeries.push(r.ms); }
report.warm.wireCheck = stats(wireSeries);

const appSeries = [];
for (let i = 0; i < 5; i++) { const r = await timeCall("wps_common_get_app_info", {}); appSeries.push(r.ms); }
report.warm.getAppInfo = stats(appSeries);

const excelSeries = [];
for (let i = 0; i < 5; i++) { const r = await timeCall("wps_excel_get_sheet_list", {}); excelSeries.push(r.ms); }
report.warm.getSheetList = stats(excelSeries);

report.coldPingText = coldPing.text.slice(0, 60);
report.excelText = (await timeCall("wps_excel_get_sheet_list", {})).text.replace(/\s+/g, " ").slice(0, 80);
if (stderr.trim()) report.stderrTail = stderr.trim().split("\n").slice(-3).join(" | ").slice(0, 300);

import { writeFileSync, mkdirSync } from "node:fs";
mkdirSync("test/.artifacts", { recursive: true });
writeFileSync("test/.artifacts/bench.json", JSON.stringify(report, null, 2));
console.log("BENCH_WRITTEN test/.artifacts/bench.json");
child.kill();
process.exit(0);

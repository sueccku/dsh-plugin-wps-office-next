// FIXES 66: a force-killed client takes the resident host down with it (FIXES 65), so the WPS
// instances that host started are left behind as headless processes. The host now records what it
// started (owner pids + app kinds); the NEXT host reclaims that record - but only when the recorded
// owner is provably dead and the instance holds no unsaved work. This test drives the reclaim and the
// safety gate (no provenance -> no action). Like the other WPS tests it closes what it can reach.
// Run: node test/orphan-reclaim.test.mjs
import { spawn, spawnSync } from "node:child_process";
import { existsSync, readFileSync, rmSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const RECORD = join(homedir(), ".wps-office-mcp", "owned-apps.json");

function ps(cmd) { return String(spawnSync("powershell", ["-NoProfile", "-Command", cmd], { encoding: "utf8" }).stdout || "").trim(); }
function appCount() { return parseInt(ps("Get-Process wps,et,wpp -ErrorAction SilentlyContinue | Measure-Object | ForEach-Object { $_.Count }") || "0", 10) || 0; }
function killAll() { ps("Get-Process wps,et,wpp -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue"); }
function recordText() { try { return readFileSync(RECORD, "utf8"); } catch { return ""; } }

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }

function startClient() {
  const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
  let buf = "";
  const pending = new Map();
  child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
  child.stderr.on("data", () => {});
  const send = (o) => child.stdin.write(JSON.stringify(o) + "\n");
  const req = (id, method, params) => new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); });
  const ready = (async () => {
    await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "orphan", version: "1" } });
    send({ jsonrpc: "2.0", method: "notifications/initialized" });
  })();
  return { child, req, ready };
}
const ping = (c) => c.req(2, "tools/call", { name: "wps_call", arguments: { tool: "wps_execute_method", args: { method: "ping", params: {} } } });

// --- 1) a force-killed session leaves both an instance and a record ----------------------------
killAll();
await sleep(2500);
rmSync(RECORD, { force: true });
const before = appCount();
const a = startClient();
await a.ready;
await a.req(2, "tools/call", { name: "wps_word_create_document", arguments: {} });
await sleep(1500);
const during = appCount();
check("the first session started a WPS application", during > before, "before=" + before + " during=" + during);
a.child.kill();
await sleep(5000);
check("the session recorded the application it started", /word/.test(recordText()), recordText().slice(0, 120) || "no record file");
const orphaned = appCount();
check("the force-killed session left the instance behind", orphaned > 0, "apps=" + orphaned);

// --- 2) the next session reclaims it -----------------------------------------------------------
const b = startClient();
await b.ready;
await ping(b);
let left = appCount();
for (let i = 0; i < 20 && left > 0; i++) { await sleep(1000); left = appCount(); }
check("the next session reclaims the orphan", left === 0, "apps=" + left);
check("the record is cleared after reclaiming", !existsSync(RECORD), recordText().slice(0, 120) || "cleared");
b.child.kill();
await sleep(2500);

// --- 3) safety: no provenance, no action -------------------------------------------------------
const c = startClient();
await c.ready;
await c.req(2, "tools/call", { name: "wps_word_create_document", arguments: {} });
await sleep(1500);
check("session three started an application", appCount() > 0, "apps=" + appCount());
c.child.kill();
await sleep(4000);
rmSync(RECORD, { force: true });   // pretend this instance was never started by us
const d = startClient();
await d.ready;
await ping(d);
await sleep(4000);
check("without a record the running application is left alone", appCount() > 0, "apps=" + appCount());
d.child.kill();
await sleep(2000);
killAll();
await sleep(1500);

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "ORPHAN RECLAIM TESTS OK (" + results.length + ")" : "ORPHAN RECLAIM TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

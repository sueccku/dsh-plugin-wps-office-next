// FIXES 66: a force-killed client takes the resident host down with it (FIXES 65), so the WPS
// instances that host started are left behind as headless processes. The host now records what it
// started (owner pids + app kinds); the NEXT host reclaims that record - but only when the recorded
// owner is provably dead and the instance holds no unsaved work. This test drives the reclaim and the
// safety gate (no provenance -> no action). Like the other WPS tests it closes what it can reach.
// Run: node test/orphan-reclaim.test.mjs
import { spawn, spawnSync } from "node:child_process";
import { existsSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";

// W1-7：宿主脚本直连所需的两个常量（场景 4 用；写法照抄 host-lease.test.mjs）。
const HOST = resolve("host/wps-com-host.ps1");
const POWERSHELL = process.env.WPS_OFFICE_POWERSHELL ||
  join(process.env.SystemRoot || "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const RECORD = join(homedir(), ".wps-office-mcp", "owned-apps.json");

function ps(cmd) { return String(spawnSync("powershell", ["-NoProfile", "-Command", cmd], { encoding: "utf8" }).stdout || "").trim(); }
function appCount() { return parseInt(ps("Get-Process wps,et,wpp -ErrorAction SilentlyContinue | Measure-Object | ForEach-Object { $_.Count }") || "0", 10) || 0; }
// FIXES 106: a WPS instance started through COM is a whole family - the frame wps.exe (started by
// svchost/DCOM), et.exe/wpp.exe, and WPS's own cloud service wpscloudsvr with its wps.exe workers.
// That service SURVIVES Quit() and keeps its workers alive: measured right after a successful reclaim,
// Kwps.Application is already gone from the ROT while wpscloudsvr + 2 workers are still running (and
// wpscloudsvr restarts a killed worker within ~1s). Those are not the orphan we reclaim, so the
// "our instance is gone" assertions must not count them - the old raw count could never reach 0.
function appInstanceCount() {
  if (appCount() === 0) return 0;   // fast path: nothing WPS-ish at all
  const rows = (ps("Get-CimInstance Win32_Process | Where-Object { @('wps','et','wpp','wpscloudsvr') -contains ($_.Name -replace '\\.exe$','') } | ForEach-Object { ($_.Name -replace '\\.exe$','') + ' ' + $_.ProcessId + ' ' + $_.ParentProcessId }") || "")
    .split(/\r?\n/).map((l) => l.trim().split(/\s+/)).filter((r) => r.length === 3)
    .map(([name, pid, ppid]) => ({ name, pid, ppid }));
  const byId = new Map(rows.map((r) => [r.pid, r]));
  const serviceOwned = (row) => {
    let cur = row;
    for (let hop = 0; hop < 10 && cur; hop++) {
      if (cur.name === 'wpscloudsvr') return true;
      cur = byId.get(cur.ppid);
    }
    return false;
  };
  return rows.filter((r) => r.name !== 'wpscloudsvr' && !serviceOwned(r)).length;
}
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
let left = appInstanceCount();
for (let i = 0; i < 20 && left > 0; i++) { await sleep(1000); left = appInstanceCount(); }
check("the next session reclaims the orphan", left === 0, "apps=" + left);
// The direct evidence that the reclaim did its job: the application left the running-object table.
// (The process count alone cannot show it - WPS's cloud service keeps its own workers alive.)
const rotAfter = ps("try { [void][System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application'); 'present' } catch { 'gone' }");
check("the reclaimed application is gone from the ROT", /gone/.test(rotAfter), rotAfter);
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

// --- 4) W1-7：实例里有**未保存**内容时，__shutdown 不许把它退出 -------------------------------
// 这是数据丢失级护栏（host/wps-actions.ps1 的 Close-WpsAppsStartedByUs）：宿主只退出"自己启动的、
// 且没有未保存内容"的实例，读状态抛异常时按"脏"处理。此前全测试目录**零覆盖**，只在文件头注释里提过。
// 这里直接驱动宿主（不走 MCP 客户端），才能精确控制 __shutdown 的时机。
killAll();
await sleep(2000);
rmSync(RECORD, { force: true });
const pidsBefore = new Set(ps("Get-Process wps,et,wpp -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }").split(/\s+/).filter(Boolean));
const hostFrames = [];
const hostProc = spawn(POWERSHELL, ["-NoProfile", "-NoLogo", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-File", HOST], {
  windowsHide: true, stdio: ["pipe", "pipe", "pipe"], env: { ...process.env, WPS_OFFICE_CLIENT_PID: String(process.pid) },
});
let hostBuf = "";
hostProc.stdout.on("data", (d) => { hostBuf += d.toString(); let i; while ((i = hostBuf.indexOf("\n")) >= 0) { const line = hostBuf.slice(0, i).trim(); hostBuf = hostBuf.slice(i + 1); if (!line) continue; try { hostFrames.push(JSON.parse(line)); } catch { /* non-protocol line */ } } });
hostProc.stderr.on("data", () => {});
const sendHost = (o) => hostProc.stdin.write(JSON.stringify(o) + "\n");
const waitFrame = async (pred, ms) => { const deadline = Date.now() + ms; while (Date.now() < deadline) { const f = hostFrames.find(pred); if (f) return f; await sleep(150); } return null; };

const ready = await waitFrame((f) => f.ready !== undefined, 30000);
check("W1-7: a host started for the unsaved-work probe", !!ready, JSON.stringify(hostFrames.slice(0, 1)).slice(0, 90));
sendHost({ id: 1, action: "createDocument", params: {} });
const created = await waitFrame((f) => f.id === 1, 90000);
check("W1-7: the probe created a document", !!created && created.ok === true && !!(created.result && created.result.success === true), JSON.stringify(created || {}).slice(0, 110));
await sleep(1500);
const runningDuring = appCount();
check("W1-7: an instance is running", runningDuring > 0, "apps=" + runningDuring);
// 空文档在 WPS 眼里 Saved 就是 True（实测），所以必须先**写入内容**再谈未保存。
sendHost({ id: 3, action: "insertText", params: { text: "未保存的探针内容", position: "end" } });
const typed = await waitFrame((f) => f.id === 3, 60000);
check("W1-7: the probe typed into the document", !!typed && typed.ok === true, JSON.stringify(typed || {}).slice(0, 110));
await sleep(1200);
// **前提检查**：这份文档必须真的是未保存。若 Saved=True，护栏本来就该退出它，整个探针就是无效的。
const savedFlag = ps("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application'); [string]$w.Documents.Item(1).Saved");
check("W1-7: the document really is unsaved (or the probe proves nothing)", savedFlag === "False", "Saved=" + JSON.stringify(savedFlag));
sendHost({ id: 2, action: "__shutdown", params: {} });
const shutdown = await waitFrame((f) => f.id === 2, 30000);
const closedKinds = shutdown && shutdown.data && Array.isArray(shutdown.data.closed) ? shutdown.data.closed : [];
check("W1-7: shutdown does not claim to have closed the word instance", !closedKinds.includes("word"), JSON.stringify(closedKinds));
await sleep(3000);
check("W1-7: the instance holding unsaved work is STILL RUNNING", appCount() > 0, "apps=" + appCount());
// 收尾：只杀"我们这次新起的"实例，绝不碰测试开始前就存在的那些。
const pidsAfter = ps("Get-Process wps,et,wpp -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }").split(/\s+/).filter(Boolean).filter((p) => !pidsBefore.has(p));
if (pidsAfter.length) ps("Stop-Process -Id " + pidsAfter.join(",") + " -Force -ErrorAction SilentlyContinue");
try { hostProc.stdin.end(); } catch { /* already gone */ }
await sleep(500);
try { hostProc.kill(); } catch { /* already gone */ }
// --- 5) FIXES 102：ready 必须意味着「启动工作已完成」 ----------------------------------------
// 以前 ready 在孤儿回收**之前**发出：客户端看到 ready 就发第一个请求，而宿主还没进主循环、
// 根本不读 stdin；回收里全是对遗留 WPS 的 COM 调用且**没有超时**，卡住就是无限等
// （param-contract 那次 2 小时挂起）。下面两条把这个顺序钉死。
// ① 静态顺序：源码里 ready 帧必须在回收调用之后。
const hostSrc = readFileSync("host/wps-com-host.ps1", "utf8");
const reclaimAt = hostSrc.indexOf("$null = Invoke-WpsOrphanReclaim");
const readyAt = hostSrc.indexOf('Write-Frame (\'{"ready":true');
check("FIXES 102: the ready frame is emitted only after the orphan reclaim", reclaimAt > 0 && readyAt > reclaimAt, "reclaim@" + reclaimAt + " ready@" + readyAt);
// ② 行为顺序：给一条「owner 已死」的回收记录再起宿主 —— 看到 ready 时记录必须已经被清掉
//    （回收末尾会 Clear-WpsOwnedApps）。固定后这条是确定的，不存在竞态。
rmSync(RECORD, { force: true });
writeFileSync(RECORD, JSON.stringify({ version: 1, hostPid: 999999, clientPid: 999998, apps: ["word"], savedUtc: new Date().toISOString() }));
const host3 = spawn(POWERSHELL, ["-NoProfile", "-NoLogo", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-File", HOST], {
  windowsHide: true, stdio: ["pipe", "pipe", "pipe"], env: { ...process.env, WPS_OFFICE_CLIENT_PID: String(process.pid) },
});
let readySeen = false; let recordAtReady = null; let buf3 = "";
host3.stdout.on("data", (d) => {
  buf3 += d.toString();
  let i;
  while ((i = buf3.indexOf("\n")) >= 0) {
    const line = buf3.slice(0, i).trim(); buf3 = buf3.slice(i + 1);
    if (!line) continue;
    let f; try { f = JSON.parse(line); } catch { continue; }
    // ready 到达的**那一刻**立刻看记录 —— 这一刻才是「客户端以为可以用了」时的真相
    if (f.ready !== undefined && !readySeen) { readySeen = true; recordAtReady = existsSync(RECORD); }
  }
});
host3.stderr.on("data", () => {});
for (let i = 0; i < 200 && !readySeen; i++) await sleep(100);
check("FIXES 102: a host reaches ready for the ordering probe", readySeen, "readySeen=" + readySeen);
check("FIXES 102: the reclaim is already done when ready arrives", recordAtReady === false, "recordAtReady=" + recordAtReady);
try { host3.stdin.end(); } catch { /* already gone */ }
await sleep(500);
try { host3.kill(); } catch { /* already gone */ }
rmSync(RECORD, { force: true });

killAll();
await sleep(1500);

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "ORPHAN RECLAIM TESTS OK (" + results.length + ")" : "ORPHAN RECLAIM TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

// S2 acceptance: the resident host holds a single-instance lease over WPS.
//
// Two DSH sessions each starting a host used to drive the same WPS instance at the same time,
// which surfaces as "it randomly hangs" with nothing to point at. The second host must now refuse
// with a readable Chinese message, and the lease must not be able to wedge WPS forever: it is
// taken over when the recorded owner is provably dead (its process, its client, or its heartbeat).
//
// This file needs no WPS and no mcp/dist build of the client - it drives host/wps-com-host.ps1
// directly. Run: node test/host-lease.test.mjs
import { spawn } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';

const HOST = resolve('host/wps-com-host.ps1');
const POWERSHELL = process.env.WPS_OFFICE_POWERSHELL ||
  join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
const STATE_FILE = join(homedir(), '.wps-office-mcp', 'com-host.json');

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** Start a host and collect its protocol frames until it says ready (or refuses). */
function startHost(clientPid) {
  const env = { ...process.env, WPS_OFFICE_CLIENT_PID: String(clientPid) };
  const child = spawn(POWERSHELL, ['-NoProfile', '-NoLogo', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', HOST], {
    windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'], env,
  });
  const state = { child, frames: [], exited: null, stderr: '' };
  let buf = '';
  child.stdout.on('data', (d) => {
    buf += d.toString();
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1);
      if (!line) continue;
      try { state.frames.push(JSON.parse(line)); } catch { /* non-protocol line */ }
    }
  });
  child.stderr.on('data', (d) => { state.stderr += d.toString(); });
  child.on('exit', (code) => { state.exited = code; });
  state.ready = async (ms = 20000) => {
    const deadline = Date.now() + ms;
    while (Date.now() < deadline) {
      const frame = state.frames.find((f) => f.ready !== undefined);
      if (frame) return frame;
      if (state.exited !== null) return null;
      await sleep(100);
    }
    return null;
  };
  state.stop = async () => {
    // Closing stdin is how the real client ends a host (ReadLine returns null, the loop breaks).
    // Force-killing immediately afterwards races the exit notification, so it is only a fallback.
    try { state.child.stdin.end(); } catch { /* already gone */ }
    const deadline = Date.now() + 4000;
    while (Date.now() < deadline && state.exited === null) await sleep(100);
    if (state.exited === null) { try { state.child.kill(); } catch { /* already gone */ } }
    await sleep(200);
  };
  return state;
}

const leaseOf = () => { try { return JSON.parse(readFileSync(STATE_FILE, 'utf8')); } catch { return null; } };
/** The exit event is asynchronous, so an assertion right after kill() races it. */
async function waitExit(state, ms = 5000) {
  const deadline = Date.now() + ms;
  while (Date.now() < deadline && state.exited === null) await sleep(100);
  return state.exited;
}

// ---------------- first host owns the lease ----------------
const a = startHost(process.pid);
const readyA = await a.ready();
check('first host becomes ready', !!readyA && readyA.ready === true, JSON.stringify(readyA).slice(0, 110));
check('ready frame carries its own pid and the client pid', !!readyA && Number(readyA.pid) > 0 && Number(readyA.clientPid) === process.pid, readyA ? 'pid=' + readyA.pid + ' clientPid=' + readyA.clientPid : 'no frame');

const lease = leaseOf();
check('lease file records the owner and a heartbeat', !!lease && Number(lease.hostPid) === Number(readyA.pid) && !!lease.updatedUtc, JSON.stringify(lease).slice(0, 140));
check('lease file records the owning client', !!lease && Number(lease.clientPid) === process.pid, lease ? 'clientPid=' + lease.clientPid : 'no lease');

// ---------------- second host refuses ----------------
const b = startHost(process.pid);
const readyB = await b.ready();
check('second host refuses instead of racing', !!readyB && readyB.ready === false, JSON.stringify(readyB || {}).slice(0, 110));
check('the refusal is a readable Chinese message naming the cause', !!readyB && /另一个 DSH 会话/.test(String(readyB.error)) && /WPS/.test(String(readyB.error)), readyB ? String(readyB.error).slice(0, 90) : 'no frame');
const bExit = await waitExit(b);
check('the refused host exits non-zero', bExit !== null && bExit !== 0, 'exit=' + bExit);
await b.stop();

// A refused host must not have disturbed the lease holder.
check('the lease holder is unaffected by the refusal', a.exited === null, 'exited=' + a.exited);

// ---------------- lease is released on exit ----------------
await a.stop();
const aExit = await waitExit(a);
check('lease holder exited', aExit !== null, 'exit=' + aExit);
const c = startHost(process.pid);
const readyC = await c.ready();
check('a new host takes the lease once the owner exits', !!readyC && readyC.ready === true, JSON.stringify(readyC || {}).slice(0, 110));
await c.stop();
check('second holder exits too', (await waitExit(c)) !== null, 'exit=' + c.exited);

// ---------------- a dead owner is taken over ----------------
// The owner is given a client pid that dies immediately, which is what a killed MCP server looks
// like from the host's side. The next host must reclaim rather than report a phantom conflict.
const dummy = spawn(process.execPath, ['-e', 'setTimeout(() => {}, 60000)'], { stdio: 'ignore', windowsHide: true });
const dummyPid = dummy.pid;
await sleep(200);
const d = startHost(dummyPid);
const readyD = await d.ready();
check('host with a doomed client starts', !!readyD && readyD.ready === true, JSON.stringify(readyD || {}).slice(0, 110));
check('that host is the recorded owner', Number(leaseOf()?.hostPid) === Number(readyD?.pid), 'recorded=' + leaseOf()?.hostPid + ' actual=' + readyD?.pid);
try { dummy.kill(); } catch { /* already gone */ }
await sleep(400);

const e = startHost(process.pid);
const readyE = await e.ready();
check('a stale owner (dead client) is taken over, not reported as a conflict', !!readyE && readyE.ready === true, JSON.stringify(readyE || {}).slice(0, 130));
const dExit = await waitExit(d);
check('the stale owner was stopped', dExit !== null, 'exit=' + dExit);
check('the takeover winner is the new host', Number(leaseOf()?.hostPid) === Number(readyE?.pid), 'recorded=' + leaseOf()?.hostPid + ' actual=' + readyE?.pid);
await d.stop();
await e.stop();

// ---------------- state file hygiene ----------------
const finalLease = leaseOf();
check('state file is valid JSON after the churn', !!finalLease, JSON.stringify(finalLease || {}).slice(0, 110));
check('state file never records a pid of 0', !!finalLease && Number(finalLease.hostPid) > 0, finalLease ? 'hostPid=' + finalLease.hostPid : 'none');

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'HOST LEASE TESTS OK (' + results.length + ')' : 'HOST LEASE TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

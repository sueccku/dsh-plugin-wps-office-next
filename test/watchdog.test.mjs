// S1 acceptance (client side): a timed-out call must say the state is unknown, must not retry
// silently, and must not leave the session paying a full timeout on every later call.
//
// The host is a stub here on purpose: it answers after a delay read from a file, so the timeout
// path can be exercised exactly and quickly instead of by wedging a real WPS. What is asserted:
//   1. a timeout rejects with the Chinese "state unknown" explanation and names the action;
//   2. the next call runs under the short suspect timeout, so a blocked WPS cannot pin the session;
//   3. one success restores the normal timeout;
//   4. each call reached the host exactly once - nothing was retried behind the caller's back.
//
// Needs mcp/dist built. Run: node test/watchdog.test.mjs
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { join, resolve } from 'node:path';

const ARTIFACTS = resolve('test/.artifacts/watchdog');
mkdirSync(ARTIFACTS, { recursive: true });
const stubPath = join(ARTIFACTS, 'stub-host.ps1');
const delayFile = join(ARTIFACTS, 'delay.txt');
const logFile = join(ARTIFACTS, 'requests.log');

const STUB = `$delayFile = $env:WPS_OFFICE_STUB_DELAY
$logFile = $env:WPS_OFFICE_STUB_LOG
[Console]::Out.WriteLine('{"ready":true,"pid":1,"protocol":1,"psVersion":"5.1.0.0"}')
[Console]::Out.Flush()
while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }
    $line = $line.Trim()
    if ($line.Length -eq 0) { continue }
    if ($logFile) { try { Add-Content -LiteralPath $logFile -Value $line } catch { } }
    $delay = 0
    try { if ($delayFile -and (Test-Path $delayFile)) { $delay = [int](Get-Content -LiteralPath $delayFile -Raw) } } catch { $delay = 0 }
    if ($delay -gt 0) { Start-Sleep -Milliseconds $delay }
    $req = $line | ConvertFrom-Json
    [Console]::Out.WriteLine('{"id":' + $req.id + ',"ok":true,"result":{"success":true,"data":{"stub":true}},"ms":0}')
    [Console]::Out.Flush()
}
`;
writeFileSync(stubPath, STUB.replace(/\n/g, '\r\n'), 'utf8');
rmSync(logFile, { force: true });
writeFileSync(delayFile, '0', 'utf8');

// The module reads these at import time, so they must be set before it is required.
process.env.WPS_OFFICE_HOST_SCRIPT = stubPath;
process.env.WPS_OFFICE_TIMEOUT_MS = '1500';
process.env.WPS_OFFICE_SUSPECT_TIMEOUT_MS = '800';
process.env.WPS_OFFICE_STARTUP_TIMEOUT_MS = '10000';
process.env.WPS_OFFICE_STUB_DELAY = delayFile;
process.env.WPS_OFFICE_STUB_LOG = logFile;

const require = createRequire(import.meta.url);
const { ComHost } = require('../mcp/dist/client/com-host.js');
if (!ComHost) { console.log('FAIL mcp/dist/client/com-host.js has no ComHost export'); process.exit(1); }

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
// Wall-clock numbers include PowerShell startup, which stretches under load, so the budget that was
// actually applied is read out of the message instead of being inferred from elapsed time.
const budgetOf = (message) => { const m = /超过 (\d+)ms/.exec(message || ''); return m ? Number(m[1]) : 0; };
const setDelay = (ms) => writeFileSync(delayFile, String(ms), 'utf8');
const host = new ComHost();

async function timed(action) {
  const started = Date.now();
  try {
    const value = await host.invoke(action, {});
    return { ok: true, value, ms: Date.now() - started };
  } catch (error) {
    return { ok: false, error: String(error && error.message), ms: Date.now() - started };
  }
}

// 1. A call that never comes back must fail with the state-unknown explanation.
setDelay(60000);
const first = await timed('ping');
check('a blocked call rejects instead of hanging forever', first.ok === false, first.ok ? 'it succeeded' : first.ms + 'ms');
check('the timeout names the action', !first.ok && first.error.includes('ping'), (first.error || '').slice(0, 80));
check('the timeout says the state is unknown', !first.ok && first.error.includes('状态未知'), (first.error || '').slice(0, 110));
check('the timeout points at a blocking dialog as the likely cause', !first.ok && first.error.includes('对话框'), (first.error || '').slice(0, 110));
check('the timeout offers a next step and refuses to kill WPS', !first.ok && first.error.includes('不会自动关闭 WPS') && first.error.includes('重试'), (first.error || '').slice(0, 140));
check('the first timeout used the normal budget', budgetOf(first.error) === 1500 && first.ms >= 1400, first.ms + 'ms, budget=' + budgetOf(first.error));
check('the host is marked suspect after the timeout', host.isSuspect === true, 'isSuspect=' + host.isSuspect);

// 2. The next call must fail fast: a blocked WPS cannot make every later call pay a full timeout.
// Both numbers include the fresh host's startup, so the budget itself is read from the message.
const second = await timed('ping');
check('the next call runs under the short suspect timeout', budgetOf(second.error) === 800, 'budget=' + budgetOf(second.error));
check('the short leash is a strictly smaller budget than the normal one', budgetOf(second.error) > 0 && budgetOf(second.error) < budgetOf(first.error), 'normal=' + budgetOf(first.error) + 'ms short=' + budgetOf(second.error) + 'ms');
check('the short-leash call still explains itself', !second.ok && second.error.includes('状态未知'), (second.error || '').slice(0, 80));
check('it is still suspect after a second failure', host.isSuspect === true, 'isSuspect=' + host.isSuspect);

// 3. One success clears the suspicion and restores the normal budget.
setDelay(0);
const recovered = await timed('ping');
check('a call that answers again succeeds', recovered.ok === true, JSON.stringify(recovered).slice(0, 90));
check('success clears the suspect flag', host.isSuspect === false, 'isSuspect=' + host.isSuspect);

setDelay(60000);
const fourth = await timed('ping');
check('the normal timeout is back after a success', budgetOf(fourth.error) === 1500 && fourth.ms >= 1400, fourth.ms + 'ms, budget=' + budgetOf(fourth.error));

// 4. Nothing may be retried behind the caller's back: four calls, four requests delivered.
const logged = existsSync(logFile) ? readFileSync(logFile, 'utf8').trim().split(/\r?\n/).filter(Boolean) : [];
check('no call was silently retried', logged.length === 4, logged.length + ' request(s) reached the host for 4 calls');

setDelay(0);
await host.stop();
rmSync(logFile, { force: true });
rmSync(delayFile, { force: true });
rmSync(stubPath, { force: true });

const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'WATCHDOG TESTS OK (' + results.length + ')' : 'WATCHDOG TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

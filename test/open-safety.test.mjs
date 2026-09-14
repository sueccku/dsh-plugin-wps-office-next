// S1 acceptance: opening a file must never pin the session on a modal dialog.
//
// The failure mode: WPS answers a password-protected document with a modal 文档已加密 prompt.
// The prompt blocks the STA thread the whole COM host runs on, so the request never returns, the
// client times out, and every later call is stuck behind the same dialog. Measured on WPS 12.1 x64
// (docs/FIXES.md 52): the raw call hangs forever, an empty PasswordDocument hangs too, and only a
// non-empty sentinel password turns it into a fast error.
//
// Every call here is watched: if one does not come back inside its budget the MCP child is killed
// and the run fails, which is exactly the behaviour a customer would experience as "it froze".
//
// Needs a real WPS. Run: node test/open-safety.test.mjs
import { spawn, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

const ARTIFACTS = resolve('test/.artifacts/open-safety');
mkdirSync(ARTIFACTS, { recursive: true });
const normalDoc = join(ARTIFACTS, 'normal.docx');
const normalWorkbook = join(ARTIFACTS, 'normal.xlsx');
const lockedDoc = join(ARTIFACTS, 'locked.docx');
const lockedWorkbook = join(ARTIFACTS, 'locked.xlsx');
const mismatchDoc = join(ARTIFACTS, 'mismatch.doc');

const POWERSHELL = process.env.WPS_OFFICE_POWERSHELL ||
  join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }

// ---------------- fixtures, straight through raw COM (never through the plugin) ----------------
// Each password-protected fixture is created in its own PowerShell process, and every fixture is
// then VERIFIED rather than assumed. WPS's Document.Password setter is unreliable by password
// length (see docs/FIXES.md 52): 15 characters blocked the call outright and 17 characters silently
// produced an unprotected file - an ordinary ZIP with no encryption at all - while 4-14 characters
// work. So the fixtures use a 12-character password and are checked to really refuse the sentinel;
// otherwise an unencrypted "encrypted" fixture would make this whole file pass for the wrong reason.
const FIXTURE_SCRIPT = `param([string]$Mode, [string]$Path)
$ErrorActionPreference = 'Stop'
switch ($Mode) {
  'plain-doc' {
    $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application')
    $d = $word.Documents.Add()
    $d.Content.Text = 'open safety: ordinary document'
    $d.SaveAs($Path, 16)
    $d.Close(0)
  }
  'plain-xls' {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application')
    $wb = $excel.Workbooks.Add()
    $wb.Sheets.Item(1).Range('A1').Value2 = 'open safety'
    $wb.SaveAs($Path, 51)
    $wb.Close($false)
  }
  'locked-doc' {
    $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application')
    $d = $word.Documents.Add()
    $d.Content.Text = 'open safety: document that needs a password'
    $d.Password = 'open-safe-pw'
    $d.SaveAs($Path, 16)
    $d.Close(0)
  }
  'locked-xls' {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application')
    $wb = $excel.Workbooks.Add()
    $wb.Sheets.Item(1).Range('A1').Value2 = 'locked'
    # SaveAs(FileName, FileFormat, Password, ...)
    $wb.SaveAs($Path, 51, 'open-safe-pw')
    $wb.Close($false)
  }
  'mismatch' {
    Set-Content -Path $Path -Value "plain text wearing a .doc extension" -Encoding ASCII
  }
  'probe-doc' {
    # The verification step: an encrypted document MUST refuse the sentinel password.
    $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application')
    $sentinel = 'dsh-wps-office-next/no-password-supplied'
    $d = $word.Documents.Open($Path, $false, $false, $true, $sentinel, "", $false, "", "", 0, 0, $false, $false, 0, $true)
    $d.Close(0)
    'IT OPENED - NOT ENCRYPTED'
  }
  'probe-xls' {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application')
    $sentinel = 'dsh-wps-office-next/no-password-supplied'
    $wb = $excel.Workbooks.Open($Path, 0, $false, $null, $sentinel, "", $true)
    $wb.Close($false)
    'IT OPENED - NOT ENCRYPTED'
  }
}
'OK'
`;

const fixturePath = join(ARTIFACTS, 'fixtures.ps1');
writeFileSync(fixturePath, FIXTURE_SCRIPT.replace(/\n/g, '\r\n'), 'utf8');

function fixture(mode, path, timeout = 90000) {
  return spawnSync(POWERSHELL, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', fixturePath, '-Mode', mode, '-Path', path], { encoding: 'utf8', timeout, windowsHide: true });
}

for (const [mode, path] of [['plain-doc', normalDoc], ['plain-xls', normalWorkbook], ['locked-doc', lockedDoc], ['locked-xls', lockedWorkbook], ['mismatch', mismatchDoc]]) {
  const run = fixture(mode, path);
  if (run.status !== 0) check('fixture ' + mode + ' built', false, 'exit=' + run.status + ' ' + String(run.stderr || run.error || '').slice(0, 200));
}

const fixturesOk = [normalDoc, normalWorkbook, lockedDoc, lockedWorkbook, mismatchDoc].every((p) => existsSync(p));
check('fixtures were built through raw COM', fixturesOk, fixturesOk ? ARTIFACTS : 'a fixture is missing');
if (!fixturesOk) {
  console.log('OPEN SAFETY TESTS FAILED (no fixtures)');
  process.exit(1);
}

// An unencrypted "encrypted" fixture would make the whole file vacuous, so prove it first.
const lockedDocProbe = fixture('probe-doc', lockedDoc, 30000);
const lockedXlsProbe = fixture('probe-xls', lockedWorkbook, 30000);
check('locked.docx really is password-protected', !String(lockedDocProbe.stdout || '').includes('IT OPENED'), 'probe output: ' + String(lockedDocProbe.stdout || '').trim().slice(0, 40));
check('locked.xlsx really is password-protected', !String(lockedXlsProbe.stdout || '').includes('IT OPENED'), 'probe output: ' + String(lockedXlsProbe.stdout || '').trim().slice(0, 40));

// ---------------- drive the real plugin ----------------
const child = spawn(process.execPath, ['mcp/dist/index.js'], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
let buf = '';
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + '\n'); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); }); }
child.stdout.on('data', (d) => { buf += d.toString(); let i; while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on('data', () => {});

function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }

let id = 10;
async function call(name, args, ms) {
  const started = Date.now();
  let timer;
  const race = new Promise((r) => { timer = setTimeout(() => r('__TIMEOUT__'), ms); });
  const res = await Promise.race([req(id++, 'tools/call', { name, arguments: args }), race]);
  clearTimeout(timer);
  const elapsed = Date.now() - started;
  if (res === '__TIMEOUT__') {
    console.log('FAIL ' + name + ' HUNG for ' + ms + 'ms -> killing the MCP child (this is the bug S1 exists to prevent)');
    child.kill();
    process.exit(1);
  }
  return { res, elapsed };
}
const viaAction = (method, params, ms) => call('wps_call', { tool: 'wps_execute_method', args: { method, params: params || {} } }, ms || 60000);
const openCount = async (method) => { const { res } = await viaAction(method); try { return JSON.parse(text(res)).data?.count ?? -1; } catch { return -1; } };
const closeAll = async (method, closeMethod) => {
  const { res } = await viaAction(method);
  let names = [];
  try { const data = JSON.parse(text(res)).data; names = (data.workbooks || data.documents || data.presentations || []).map((item) => item.name); } catch { }
  for (const name of names) await viaAction(closeMethod, { name, save: false });
};

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'open-safety', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

// Ordinary opens must keep working: the guard must not break the happy path.
const before = await openCount('getOpenDocuments');
const plain = await call('wps_word_open_document', { filePath: normalDoc }, 90000);
check('ordinary .docx opens', ok(plain.res), text(plain.res).replace(/\s+/g, ' ').slice(0, 90));
await closeAll('getOpenDocuments', 'closeDocument');
check('no document left open', (await openCount('getOpenDocuments')) === before, 'documents=' + (await openCount('getOpenDocuments')));

const wbBefore = await openCount('getOpenWorkbooks');
const plainXls = await call('wps_excel_open_workbook', { filePath: normalWorkbook }, 90000);
check('ordinary .xlsx opens', ok(plainXls.res), text(plainXls.res).replace(/\s+/g, ' ').slice(0, 90));
await closeAll('getOpenWorkbooks', 'closeWorkbook');
check('no workbook left open', (await openCount('getOpenWorkbooks')) === wbBefore, 'workbooks=' + (await openCount('getOpenWorkbooks')));

// A text file wearing a .doc extension: the format/extension mismatch used to be the other
// candidate for a conversion prompt. It must return, not hang.
const mismatch = await call('wps_word_open_document', { filePath: mismatchDoc }, 45000);
check('a mismatched-extension file returns instead of prompting', !!mismatch.res, 'returned in ' + mismatch.elapsed + 'ms, ok=' + ok(mismatch.res));
await closeAll('getOpenDocuments', 'closeDocument');

// The regression proper: an encrypted document must fail fast with an actionable Chinese message.
const lockedWord = await call('wps_word_open_document', { filePath: lockedDoc }, 30000);
const lockedWordText = text(lockedWord.res);
check('encrypted .docx returns instead of hanging', !!lockedWord.res, 'returned in ' + lockedWord.elapsed + 'ms');
check('encrypted .docx fails', !ok(lockedWord.res), lockedWordText.replace(/\s+/g, ' ').slice(0, 110));
check('encrypted .docx explains the password and the next step', /密码|加密/.test(lockedWordText) && /另存为/.test(lockedWordText), lockedWordText.replace(/\s+/g, ' ').slice(0, 150));
check('encrypted .docx fails inside the timeout budget', lockedWord.elapsed < 20000, lockedWord.elapsed + 'ms');

const lockedExcel = await call('wps_excel_open_workbook', { filePath: lockedWorkbook }, 30000);
const lockedExcelText = text(lockedExcel.res);
check('encrypted .xlsx returns instead of hanging', !!lockedExcel.res, 'returned in ' + lockedExcel.elapsed + 'ms');
check('encrypted .xlsx fails', !ok(lockedExcel.res), lockedExcelText.replace(/\s+/g, ' ').slice(0, 110));
check('encrypted .xlsx explains the password and the next step', /密码|加密/.test(lockedExcelText) && /另存为/.test(lockedExcelText), lockedExcelText.replace(/\s+/g, ' ').slice(0, 150));
check('encrypted .xlsx fails inside the timeout budget', lockedExcel.elapsed < 20000, lockedExcel.elapsed + 'ms');

// The real proof that nothing is left blocking: the host answers a fresh call afterwards.
const after = await call('wps_common_ping', {}, 20000);
check('the bridge still answers after the encrypted opens', ok(after.res), text(after.res).replace(/\s+/g, ' ').slice(0, 80));

// A missing file must be an error, never a wait.
const missing = await call('wps_word_open_document', { filePath: join(ARTIFACTS, 'definitely-not-here.docx') }, 30000);
check('a missing file fails fast with a readable reason', !ok(missing.res) && missing.elapsed < 20000, 'returned in ' + missing.elapsed + 'ms: ' + text(missing.res).replace(/\s+/g, ' ').slice(0, 90));

// Nothing may be left open by the failed opens.
const docLeak = await openCount('getOpenDocuments');
const xlsLeak = await openCount('getOpenWorkbooks');
check('no document leaked by the failed opens', docLeak === before, 'documents=' + docLeak);
check('no workbook leaked by the failed opens', xlsLeak === wbBefore, 'workbooks=' + xlsLeak);
// Even when the checks above fail, do not hand the next test a polluted WPS.
await closeAll('getOpenDocuments', 'closeDocument');
await closeAll('getOpenWorkbooks', 'closeWorkbook');

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'OPEN SAFETY TESTS OK (' + results.length + ')' : 'OPEN SAFETY TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

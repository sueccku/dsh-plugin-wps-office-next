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

# WPS can answer GetActiveObject with a hollow instance whose collections are null (the bridge
# guards against exactly this in Test-WpsAppUsable). Using the raw call here once left Excel
# unusable for every test that ran afterwards, so the fixtures validate before they accept one.
function Get-App([string]$progId, [string]$collection) {
    $candidates = @()
    try { $candidates += [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId) } catch { }
    try { $candidates += (New-Object -ComObject $progId) } catch { }
    foreach ($candidate in $candidates) {
        if ($null -eq $candidate) { continue }
        try { if ($null -ne $candidate.$collection) { return $candidate } } catch { }
    }
    throw "no usable $progId instance"
}

function Set-AlertsOff($app) {
    $previous = $null
    try { $previous = $app.DisplayAlerts; $app.DisplayAlerts = $false } catch { $previous = $null }
    return $previous
}

function Restore-Alerts($app, $previous) {
    if ($null -eq $previous) { return }
    try { $app.DisplayAlerts = $previous } catch { }
}

function Close-Item($app, $item, $save) {
    # A raw Close() can pop a modal save prompt, which then blocks every later call - the trap the
    # bridge's own close paths suppress with DisplayAlerts (FIXES 24-29).
    $previous = Set-AlertsOff $app
    try { $item.Close($save) } finally { Restore-Alerts $app $previous }
}

function Clear-Strays {
    # Whatever earlier runs left behind would otherwise be counted as this run's leak.
    # ONLY never-saved items (no path on disk) are touched: that is what a test zombie looks like,
    # and it means a real document someone has open - saved or not - is never closed by this test.
    $word = $null
    try { $word = Get-App 'Kwps.Application' 'Documents' } catch { }
    if ($null -ne $word) {
        $previous = Set-AlertsOff $word
        try {
            for ($i = $word.Documents.Count; $i -ge 1; $i--) {
                try { $item = $word.Documents.Item($i); if ([string]$item.Path -eq '') { Close-Item $word $item 0 } } catch { }
            }
        } finally { Restore-Alerts $word $previous }
    }
    $excel = $null
    try { $excel = Get-App 'Ket.Application' 'Workbooks' } catch { }
    if ($null -ne $excel) {
        $previous = Set-AlertsOff $excel
        try {
            for ($i = $excel.Workbooks.Count; $i -ge 1; $i--) {
                try { $item = $excel.Workbooks.Item($i); if ([string]$item.Path -eq '') { Close-Item $excel $item $false } } catch { }
            }
        } finally { Restore-Alerts $excel $previous }
    }
}

switch ($Mode) {
  'preflight' {
    # Proves the shared instance is usable BEFORE the test blames the plugin for anything: a hollow
    # instance or a pending modal dialog shows up here, once, with a clear message.
    Clear-Strays
    $word = Get-App 'Kwps.Application' 'Documents'
    $wp = Set-AlertsOff $word
    try {
      $d = $word.Documents.Add()
      $paragraphs = [int]$d.Paragraphs.Count
      Close-Item $word $d 0
      if ($paragraphs -lt 1) { throw "Word Add() produced a document with $paragraphs paragraph(s) - the instance is hollow" }
    } finally { Restore-Alerts $word $wp }
    $excel = Get-App 'Ket.Application' 'Workbooks'
    $ep = Set-AlertsOff $excel
    try {
      $wb = $excel.Workbooks.Add()
      $sheets = [int]$wb.Sheets.Count
      if ($sheets -ge 1) { $wb.Sheets.Item(1).Range('A1').Value2 = 'preflight' }
      Close-Item $excel $wb $false
      if ($sheets -lt 1) { throw "Excel Add() produced a workbook with $sheets sheet(s) - the instance is hollow" }
    } finally { Restore-Alerts $excel $ep }
    'PREFLIGHT OK'
  }
  'plain-doc' {
    $word = Get-App 'Kwps.Application' 'Documents'
    $previous = Set-AlertsOff $word
    try {
      $d = $word.Documents.Add()
      $d.Content.Text = 'open safety: ordinary document'
      $d.SaveAs($Path, 16)
      Close-Item $word $d 0
    } finally { Restore-Alerts $word $previous }
  }
  'plain-xls' {
    $excel = Get-App 'Ket.Application' 'Workbooks'
    $previous = Set-AlertsOff $excel
    try {
      $wb = $excel.Workbooks.Add()
      $wb.Sheets.Item(1).Range('A1').Value2 = 'open safety'
      $wb.SaveAs($Path, 51)
      Close-Item $excel $wb $false
    } finally { Restore-Alerts $excel $previous }
  }
  'locked-doc' {
    $word = Get-App 'Kwps.Application' 'Documents'
    $previous = Set-AlertsOff $word
    try {
      $d = $word.Documents.Add()
      $d.Content.Text = 'open safety: document that needs a password'
      $d.Password = 'open-safe-pw'
      $d.SaveAs($Path, 16)
      Close-Item $word $d 0
    } finally { Restore-Alerts $word $previous }
  }
  'locked-xls' {
    $excel = Get-App 'Ket.Application' 'Workbooks'
    $previous = Set-AlertsOff $excel
    try {
      $wb = $excel.Workbooks.Add()
      $wb.Sheets.Item(1).Range('A1').Value2 = 'locked'
      # SaveAs(FileName, FileFormat, Password, ...)
      $wb.SaveAs($Path, 51, 'open-safe-pw')
      Close-Item $excel $wb $false
    } finally { Restore-Alerts $excel $previous }
  }
  'mismatch' {
    Set-Content -Path $Path -Value "plain text wearing a .doc extension" -Encoding ASCII
  }
  'probe-doc' {
    # The verification step: an encrypted document MUST refuse the sentinel password.
    $word = Get-App 'Kwps.Application' 'Documents'
    $previous = Set-AlertsOff $word
    $sentinel = 'dsh-wps-office-next/no-password-supplied'
    try {
      $d = $word.Documents.Open($Path, $false, $false, $true, $sentinel, "", $false, "", "", 0, 0, $false, $false, 0, $true)
      Close-Item $word $d 0
      'IT OPENED - NOT ENCRYPTED'
    } catch {
      'REFUSED'
    } finally { Restore-Alerts $word $previous }
  }
  'probe-xls' {
    $excel = Get-App 'Ket.Application' 'Workbooks'
    $previous = Set-AlertsOff $excel
    $sentinel = 'dsh-wps-office-next/no-password-supplied'
    try {
      $wb = $excel.Workbooks.Open($Path, 0, $false, $null, $sentinel, "", $true)
      Close-Item $excel $wb $false
      'IT OPENED - NOT ENCRYPTED'
    } catch {
      'REFUSED'
    } finally { Restore-Alerts $excel $previous }
  }
}
'OK'
`;

const fixturePath = join(ARTIFACTS, 'fixtures.ps1');
writeFileSync(fixturePath, FIXTURE_SCRIPT.replace(/\n/g, '\r\n'), 'utf8');

function fixture(mode, path, timeout = 90000) {
  return spawnSync(POWERSHELL, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', fixturePath, '-Mode', mode, '-Path', path], { encoding: 'utf8', timeout, windowsHide: true });
}

// Health first. A wedged WPS (hollow instance, or a modal dialog left by some earlier run) makes
// every check below meaningless, so it is reported once, as an environment problem, instead of
// three confusing plugin failures. Recovery is documented in docs/HANDOFF.md section 10.
const preflight = fixture('preflight', '', 120000);
const preflightOut = String(preflight.stdout || '');
const preflightOk = preflightOut.includes('PREFLIGHT OK');
check('WPS is healthy before the test starts (pre-flight)', preflightOk, preflightOk ? 'instance usable' : 'see the message below');
if (!preflightOk) {
  console.log('');
  console.log('OPEN SAFETY TESTS FAILED (environment not healthy - not a plugin defect)');
  console.log('  stdout: ' + preflightOut.trim().slice(0, 200));
  console.log('  stderr: ' + String(preflight.stderr || preflight.error || '').trim().slice(0, 400));
  console.log('  修法：关掉残留的 WPS 模态对话框（class Qt*），或结束该应用的 WPS 进程后重跑。详见 docs/HANDOFF.md §10。');
  process.exit(1);
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

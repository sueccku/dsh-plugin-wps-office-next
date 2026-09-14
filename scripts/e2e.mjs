// One-command end-to-end acceptance run for dsh-plugin-wps-office-next.
//
// What it does, in order:
//   1. builds a fresh fixture workbook under test/.artifacts/e2e/<run>/ with raw WPS COM
//      (deliberately not through this plugin, so a tool bug cannot fake its own input)
//   2. runs one real headless DSH task against a profile that has this bundle installed
//   3. finds that run's session log, decodes every appended zstd frame, prints the tool trace
//   4. re-opens the produced workbook/document with raw COM and checks their contents
//   5. asserts behaviour: nothing left open, no defect markers in tool results, and the model
//      driving WPS through the plugin instead of writing its own COM script
//
// Usage:
//   node scripts/e2e.mjs --profile <name>          run the acceptance test
//   node scripts/e2e.mjs --profile <name> --setup  create that profile and install this repo first
//
// Exit code is 0 only when every check passes. Evidence stays in test/.artifacts/e2e/<run>/.
import { spawn, spawnSync } from 'node:child_process';
import { closeSync, existsSync, mkdirSync, openSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zstdDecompressSync } from 'node:zlib';

const ROOT = dirname(dirname(fileURLToPath(import.meta.url)));
const E2E_ROOT = join(ROOT, 'test', '.artifacts', 'e2e');
const DSH_HOME = process.env.DSH_HOME || join(homedir(), '.dsh');
const TASK_PREFIX = '用 WPS 处理表格';

// Scenario data. Totals are 华北 330 / 华南 280 / 华东 240 / 西南 165, sum 1015.
const FIXTURE_ROWS = [
  ['地区', '产品', '销售额'],
  ['华东', 'A', 120], ['华东', 'B', 80], ['华南', 'A', 150], ['华南', 'C', 60],
  ['华北', 'B', 90], ['华北', 'A', 110], ['华东', 'C', 40], ['华南', 'B', 70],
  ['华北', 'C', 130], ['西南', 'A', 55], ['西南', 'B', 45], ['西南', 'C', 65],
];
const EXPECTED_ROWS = [['华北', 330], ['华南', 280], ['华东', 240], ['西南', 165]];
const EXPECTED_TOTAL = 1015;

// Scenario 2 (P2-5): a separate order workbook that the same run has to turn into a real table
// (ListObject), conditional-format its amount column and make print-ready. Amounts above 500:
// 680, 720, 900, 510, 830, 600 - six of them.
const ORDER_ROWS = [
  ['订单号', '区域', '产品', '金额'],
  ['A001', '华东', 'A', 120], ['A002', '华东', 'B', 680], ['A003', '华南', 'A', 450],
  ['A004', '华南', 'C', 720], ['A005', '华北', 'B', 300], ['A006', '华北', 'A', 900],
  ['A007', '华东', 'C', 150], ['A008', '华南', 'B', 510], ['A009', '华北', 'C', 260],
  ['A010', '西南', 'A', 830], ['A011', '西南', 'B', 90], ['A012', '西南', 'C', 600],
];
const ORDER_LAST_ROW = ORDER_ROWS.length;

// Strings an earlier real run produced. Any of them in a tool result is a regression.
const DEFECT_MARKERS = [
  'unknown parameter(',
  'Unknown action',
  '页数: undefined',
  '[object Object]',
  '计算结果: null',
  'is out of range',
  'no presentation is open',
];
// Markers that mean the model gave up on the plugin and drove COM by hand, which is the whole
// value proposition this bundle replaces. Scanned only in the model's own pwsh commands.
const OWN_COM_MARKERS = ['New-Object -ComObject', 'Ket.Application', 'Kwps.Application', 'Kwpp.Application'];

const checks = [];
function check(name, ok, detail) {
  checks.push({ name, ok: !!ok, detail: detail || '' });
  console.log((ok ? 'PASS  ' : 'FAIL  ') + name + (detail ? '  ' + detail : ''));
}
function info(message) { console.log('      ' + message); }
function die(message) { console.error('e2e: ' + message); process.exit(2); }

function printHelp() {
  console.log([
    'Usage: node scripts/e2e.mjs --profile <name> [--setup] [--dsh-bin <lib/bin.js>] [--timeout <sec>] [--clean]',
    '',
    '  --profile <name>   DSH profile that has this bundle installed (required)',
    '  --setup            create that profile from the headless template and install this repo into it',
    '  --dsh-bin <path>   path to @deepseek-ai/dsh/lib/bin.js when it cannot be located automatically',
    '  --timeout <sec>    per-run wall clock limit (default 420; the observed two-scenario run takes about 2-3 min)',
    '  --clean            delete the generated fixture and output document when every check passes',
  ].join('\n'));
}

function parseArgs(argv) {
  const opts = { profile: '', setup: false, dshBin: '', timeoutSec: 420, clean: false };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--profile') opts.profile = argv[++i] || '';
    else if (arg === '--setup') opts.setup = true;
    else if (arg === '--dsh-bin') opts.dshBin = argv[++i] || '';
    else if (arg === '--timeout') opts.timeoutSec = Number(argv[++i] || 420);
    else if (arg === '--clean') opts.clean = true;
    else if (arg === '--help' || arg === '-h') { printHelp(); process.exit(0); }
    else die('unknown argument: ' + arg);
  }
  if (!opts.profile) { printHelp(); die('--profile <name> is required'); }
  return opts;
}

// DSH is launched through its real entry rather than the .cmd shim: argv then carries the task
// text verbatim, with no cmd.exe quoting to get wrong.
function resolveDshBin(explicit) {
  const rel = join('node_modules', '@deepseek-ai', 'dsh', 'lib', 'bin.js');
  const candidates = [];
  if (explicit) candidates.push(explicit);
  if (process.env.DSH_BIN) candidates.push(process.env.DSH_BIN);
  const exeDir = dirname(process.execPath);
  candidates.push(join(exeDir, rel));
  candidates.push(join(exeDir, '..', 'lib', rel));
  if (process.env.APPDATA) candidates.push(join(process.env.APPDATA, 'npm', rel));
  const where = spawnSync('where.exe', ['dsh'], { encoding: 'utf8', windowsHide: true });
  if (!where.error && where.status === 0) {
    for (const line of String(where.stdout).split(/\r?\n/)) {
      const shim = line.trim();
      if (shim) candidates.push(join(dirname(shim), rel));
    }
  }
  return candidates.find((candidate) => candidate && existsSync(candidate)) || '';
}

function runSync(bin, args, opts) {
  return spawnSync(bin, args, Object.assign({ encoding: 'utf8', windowsHide: true }, opts || {}));
}

// Generated PowerShell must be read as UTF-8 by Windows PowerShell 5.1, which only honours a BOM,
// and its stdout must be forced back to UTF-8 or the Chinese text comes home mangled.
function writePowerShell(path, scriptLines) {
  writeFileSync(path, '\uFEFF' + scriptLines.join('\r\n') + '\r\n', 'utf8');
}
function runPowerShell(path, args, timeoutMs) {
  return runSync('powershell.exe', ['-NoProfile', '-NoLogo', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', path].concat(args), { timeout: timeoutMs || 180000 });
}

function psArray(values) {
  return values.map((value) => {
    if (typeof value === 'number') return String(value);
    return "'" + String(value).replace(/'/g, "''") + "'";
  }).join(', ');
}

// Both fixtures come out of one raw-COM session; the second workbook is the P2-5 scenario.
function fixtureScript() {
  const rows = FIXTURE_ROWS.map((row) => '  @(' + psArray(row) + ')').join(',\r\n');
  const orderRows = ORDER_ROWS.map((row) => '  @(' + psArray(row) + ')').join(',\r\n');
  return [
    'param([string]$Path, [string]$OrdersPath)',
    "$ErrorActionPreference = 'Stop'",
    '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8',
    'if (Test-Path $Path) { Remove-Item $Path -Force }',
    'if (Test-Path $OrdersPath) { Remove-Item $OrdersPath -Force }',
    '$app = New-Object -ComObject Ket.Application',
    '$app.DisplayAlerts = $false',
    'try {',
    '    $wb = $app.Workbooks.Add()',
    '    $ws = $app.ActiveSheet',
    "    $ws.Name = '明细'",
    '    $rows = @(',
    rows,
    '    )',
    "    $data = New-Object 'object[,]' $rows.Count, 3",
    '    for ($r = 0; $r -lt $rows.Count; $r++) {',
    '        for ($c = 0; $c -lt 3; $c++) {',
    '            $v = $rows[$r][$c]',
    '            if ($v -is [int]) { $data[$r, $c] = [double]$v } else { $data[$r, $c] = $v }',
    '        }',
    '    }',
    "    $ws.Range('A1:C' + $rows.Count).Value2 = $data",
    '    $wb.SaveAs($Path, 51)',
    '    $wb.Close($false)',
    '',
    '    $wb2 = $app.Workbooks.Add()',
    '    $ws2 = $app.ActiveSheet',
    "    $ws2.Name = '订单'",
    '    $orderRows = @(',
    orderRows,
    '    )',
    "    $data2 = New-Object 'object[,]' $orderRows.Count, 4",
    '    for ($r = 0; $r -lt $orderRows.Count; $r++) {',
    '        for ($c = 0; $c -lt 4; $c++) {',
    '            $v = $orderRows[$r][$c]',
    '            if ($v -is [int]) { $data2[$r, $c] = [double]$v } else { $data2[$r, $c] = $v }',
    '        }',
    '    }',
    "    $ws2.Range('A1:D' + $orderRows.Count).Value2 = $data2",
    '    $wb2.SaveAs($OrdersPath, 51)',
    '    $wb2.Close($false)',
    '} finally {',
    '    $app.Quit() | Out-Null',
    '}',
    "Write-Output ('fixture rows=' + $rows.Count + ' orders=' + $orderRows.Count + ' path=' + $Path)",
  ];
}

function verifyScript() {
  return [
    'param([string]$Workbook, [string]$Document, [string]$Orders, [int]$OrderLastRow)',
    "$ErrorActionPreference = 'Stop'",
    '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8',
    '$result = @{}',
    '',
    '# Prefer the running instance so the report describes what the agent actually left behind;',
    '# fall back to a fresh one only when nothing is running any more.',
    'function Get-App([string]$progId) {',
    '    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId) } catch { }',
    '    try { return (New-Object -ComObject $progId) } catch { return $null }',
    '}',
    '',
    "$excel = Get-App 'Ket.Application'",
    'if ($null -ne $excel) {',
    '    try { $excel.DisplayAlerts = $false } catch { }',
    '    $wb = $null',
    '    try {',
    '        $wb = $excel.Workbooks.Open($Workbook, 0, $true)',
    '        $sheets = @()',
    '        for ($i = 1; $i -le $wb.Worksheets.Count; $i++) { $sheets += [string]$wb.Worksheets.Item($i).Name }',
    '        $result.sheets = $sheets',
    '        $target = $null',
    "        foreach ($name in $sheets) { if ($name -eq '汇总') { $target = $wb.Worksheets.Item($name) } }",
    '        if ($null -ne $target) {',
    "            $vals = $target.Range('A1:B5').Value2",
    '            $rows = @()',
    '            for ($r = 1; $r -le 5; $r++) {',
    '                $rows += ,@([string]$vals.GetValue($r, 1), [string]$vals.GetValue($r, 2))',
    '            }',
    '            $result.rows = $rows',
    '            $charts = @()',
    '            for ($i = 1; $i -le $target.ChartObjects().Count; $i++) {',
    '                $chart = $target.ChartObjects($i).Chart',
    "                $title = ''",
    '                try { if ($chart.HasTitle) { $title = [string]$chart.ChartTitle.Text } } catch { }',
    '                $charts += @{ type = [int]$chart.ChartType; title = $title }',
    '            }',
    '            $result.charts = $charts',
    "            $result.numberFormatB2 = [string]$target.Range('B2').NumberFormat",
    '        }',
    '    } catch {',
    '        $result.workbookError = $_.Exception.Message',
    '    } finally {',
    '        if ($null -ne $wb) { try { $wb.Close($false) } catch { } }',
    '    }',
    '} else {',
    "    $result.workbookError = 'no Excel/WPS instance available'",
    '}',
    '',
    'if (Test-Path $Document) {',
    "    $word = Get-App 'Kwps.Application'",
    '    if ($null -ne $word) {',
    '        try { $word.DisplayAlerts = 0 } catch { }',
    '        $doc = $null',
    '        try {',
    '            $doc = $word.Documents.Open($Document, $false, $true)',
    '            $paras = @()',
    '            for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {',
    '                $p = $doc.Paragraphs.Item($i)',
    "                $style = ''",
    '                try { $style = [string]$p.Style.NameLocal } catch { }',
    '                $paras += @{ style = $style; text = ([string]$p.Range.Text).Trim() }',
    '            }',
    '            $result.document = @{ paragraphs = $paras; characters = [int]$doc.Characters.Count }',
    '        } catch {',
    '            $result.documentError = $_.Exception.Message',
    '        } finally {',
    '            if ($null -ne $doc) { try { $doc.Close($false) } catch { } }',
    '        }',
    '    } else {',
    "        $result.documentError = 'no Word/WPS instance available'",
    '    }',
    '} else {',
    '    $result.documentMissing = $true',
    '}',
    '',

    "# Scenario 2 (P2-5): the order workbook must have become a real table with print-ready setup.",
    "if (Test-Path $Orders) {",
    "    $excel2 = Get-App 'Ket.Application'",
    "    if ($null -ne $excel2) {",
    "        try { $excel2.DisplayAlerts = $false } catch { }",
    "        $owb = $null",
    "        try {",
    "            $owb = $excel2.Workbooks.Open($Orders, 0, $true)",
    "            $osheet = $owb.Worksheets.Item('订单')",
    "            $objects = @()",
    "            for ($i = 1; $i -le $osheet.ListObjects.Count; $i++) {",
    "                $lo = $osheet.ListObjects.Item($i)",
    "                $addr = ''",
    "                try { $addr = [string]$lo.Range.Address() } catch { }",
    "                $objects += @{ name = [string]$lo.Name; range = $addr }",
    "            }",
    "            # Where the rule landed is the model's choice, so probe the likely ranges and keep the",
    "            # largest count: FormatConditions belongs to a range object, not to the sheet.",
    "            $formats = 0",
    "            foreach ($probe in @(\"D2:D$OrderLastRow\", \"D1:D$OrderLastRow\", 'D:D', 'D2', \"A1:D$OrderLastRow\")) {",
    "                try { $n = [int]$osheet.Range($probe).FormatConditions.Count; if ($n -gt $formats) { $formats = $n } } catch { }",
    "            }",
    "            $result.orders = @{",
    "                listCount = [int]$osheet.ListObjects.Count",
    "                objects = $objects",
    "                formatCount = $formats",
    "                orientation = [int]$osheet.PageSetup.Orientation",
    "                paperSize = [int]$osheet.PageSetup.PaperSize",
    "                centerFooter = [string]$osheet.PageSetup.CenterFooter",
    "                printTitleRows = [string]$osheet.PageSetup.PrintTitleRows",
    "            }",
    "        } catch {",
    "            $result.ordersError = $_.Exception.Message",
    "        } finally {",
    "            if ($null -ne $owb) { try { $owb.Close($false) } catch { } }",
    "        }",
    "    } else {",
    "        $result.ordersError = 'no Excel/WPS instance available'",
    "    }",
    "} else {",
    "    $result.ordersMissing = $true",
    "}",
    '$result | ConvertTo-Json -Depth 8 -Compress',
  ];
}

// --- session log ---------------------------------------------------------------------------

const ZSTD_MAGIC = Buffer.from([0x28, 0xb5, 0x2f, 0xfd]);
// DSH appends one zstd frame per flush and node's decompressor stops after the first, so decode
// frame by frame; a truncated slice tells us the candidate boundary was inside the payload.
function decompressFrames(buf) {
  const starts = [];
  for (let i = 0; i + 4 <= buf.length; i++) if (buf.compare(ZSTD_MAGIC, 0, 4, i, i + 4) === 0) starts.push(i);
  if (!starts.length) throw new Error('no zstd frame found');
  const bounds = starts.concat([buf.length]);
  const out = [];
  let i = 0;
  while (i < bounds.length - 1) {
    let done = false;
    for (let j = i + 1; j < bounds.length && !done; j++) {
      try { out.push(zstdDecompressSync(buf.subarray(bounds[i], bounds[j]))); i = j; done = true; } catch { }
    }
    if (!done) throw new Error('undecodable frame at offset ' + bounds[i]);
  }
  return Buffer.concat(out);
}

function loadRecords(file) {
  const text = decompressFrames(readFileSync(file)).toString('utf8');
  return text.split(/\r?\n/).filter((line) => line.trim()).map((line) => {
    try { return JSON.parse(line); } catch { return { __unparsed: line }; }
  });
}

function flatten(content) {
  if (content == null) return '';
  if (typeof content === 'string') return content;
  if (Array.isArray(content)) return content.map(flatten).join(' | ');
  if (typeof content === 'object') {
    if (typeof content.text === 'string') return content.text;
    if (content.content) return flatten(content.content);
  }
  return '';
}
function oneLine(value, max) { return String(value).replace(/\s+/g, ' ').trim().slice(0, max); }

// Layout is sessions/<encoded-cwd>/<session-id>/session.v3.jsonl.zstd; both depths are accepted
// because the workspace level is easy to forget and a silent miss would quietly skip every
// behaviour check below.
function sessionFiles() {
  const root = join(DSH_HOME, 'sessions');
  const files = [];
  if (!existsSync(root)) return files;
  for (const workspace of readdirSync(root)) {
    const workspaceDir = join(root, workspace);
    let entries;
    try { entries = readdirSync(workspaceDir); } catch { continue; }
    const direct = join(workspaceDir, 'session.v3.jsonl.zstd');
    if (existsSync(direct)) { files.push({ file: direct, workspace }); continue; }
    for (const entry of entries) {
      const nested = join(workspaceDir, entry, 'session.v3.jsonl.zstd');
      if (existsSync(nested)) files.push({ file: nested, workspace });
    }
  }
  return files;
}

function findSession(sinceMs) {
  const candidates = [];
  for (const entry of sessionFiles()) {
    const file = entry.file;
    const stat = statSync(file);
    if (stat.mtimeMs < sinceMs - 15000) continue;
    let text;
    try { text = decompressFrames(readFileSync(file)).toString('utf8'); } catch { continue; }
    if (!text.includes('sales.xlsx')) continue;
    // A PTC session (this script's own parent) never drives the WPS tools natively and its first
    // user message is the human request, so it must not be mistaken for the run under test.
    if (text.includes('"name":"run_code"')) continue;
    let header = null;
    try { header = JSON.parse(text.slice(0, text.indexOf('\n'))); } catch { }
    candidates.push({ file, mtimeMs: stat.mtimeMs, header });
  }
  if (!candidates.length) info('no session file under ' + join(DSH_HOME, 'sessions') + ' was touched since the run started');
  candidates.sort((a, b) => b.mtimeMs - a.mtimeMs);
  for (const candidate of candidates) {
    const records = loadRecords(candidate.file);
    const firstUser = records.find((record) => record.type === 'user/message');
    const text = firstUser ? flatten(firstUser.data && firstUser.data.content) : '';
    if (text.trim().startsWith(TASK_PREFIX)) return { file: candidate.file, records };
  }
  return null;
}

function summarizeTrace(records) {
  const lines = [];
  const toolCalls = [];
  const toolResults = [];
  const pwshCommands = [];
  const nameByCall = new Map();
  let skillLoads = 0;
  for (const record of records) {
    const data = record.data || {};
    if (record.type === 'user/message') {
      lines.push(String(record.seq).padStart(4) + ' USER   ' + oneLine(flatten(data.content), 150));
    } else if (record.type === 'assistant/message') {
      for (const part of (data.message && data.message.content) || []) {
        if (part.type === 'text' && part.text && part.text.trim()) lines.push(String(record.seq).padStart(4) + ' ASST   ' + oneLine(part.text, 150));
      }
    } else if (record.type === 'tool/call') {
      const name = String(data.name || '');
      nameByCall.set(data.callId, name);
      toolCalls.push({ name, args: String(data.arguments || '') });
      if (name === 'skill') skillLoads++;
      if (name === 'pwsh') {
        try { pwshCommands.push(String(JSON.parse(String(data.arguments)).command || '')); } catch { }
      }
      lines.push(String(record.seq).padStart(4) + ' CALL   ' + name + ' ' + oneLine(data.arguments, 150));
    } else if (record.type === 'tool/result') {
      const name = nameByCall.get(data.message && data.message.source && data.message.source.callId) || '?';
      const text = flatten(data.message && data.message.content);
      toolResults.push({ name, text });
      lines.push(String(record.seq).padStart(4) + ' RESULT ' + name + ' ' + oneLine(text, 150));
    }
  }
  return { lines, toolCalls, toolResults, pwshCommands, skillLoads };
}

function shortToolName(name) { return name.replace(/^mcp__wps-office-next__/, ''); }

// wps_batch and wps_call carry the real work inside their arguments, so a check that only looked at
// the top-level tool name would miss, for example, a Word document created inside a batch.
function expandToolNames(toolCalls) {
  const names = [];
  for (const call of toolCalls) {
    const name = shortToolName(call.name);
    names.push(name);
    if (name !== 'wps_batch' && name !== 'wps_call') continue;
    try {
      const args = JSON.parse(call.args);
      if (name === 'wps_call') names.push(shortToolName(String(args.tool || '')));
      for (const inner of args.calls || []) names.push(shortToolName(String(inner.tool || '')));
    } catch { }
  }
  return names.filter((entry) => entry && entry.startsWith('wps_'));
}

// Session logs, traces and reports are read by Windows PowerShell 5.1 as often as by an editor, and
// it only honours UTF-8 when a BOM is present.
function writeUtf8Bom(path, text) { writeFileSync(path, '﻿' + text, 'utf8'); }

// --- our own tools, used for the leak check ------------------------------------------------

function mcpCalls(calls) {
  return new Promise((resolvePromise) => {
    const child = spawn(process.execPath, [join(ROOT, 'mcp', 'dist', 'index.js')], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
    const pending = new Map();
    let buf = '';
    let id = 0;
    const send = (payload) => child.stdin.write(JSON.stringify(payload) + '\n');
    child.stdout.on('data', (chunk) => {
      buf += chunk.toString();
      let index;
      while ((index = buf.indexOf('\n')) >= 0) {
        const line = buf.slice(0, index).trim();
        buf = buf.slice(index + 1);
        if (!line) continue;
        let message;
        try { message = JSON.parse(line); } catch { continue; }
        if (message.id && pending.has(message.id)) { pending.get(message.id)(message); pending.delete(message.id); }
      }
    });
    child.stderr.on('data', () => {});
    const request = (method, params) => new Promise((resolve) => {
      const myId = ++id;
      const timer = setTimeout(() => { if (pending.has(myId)) { pending.delete(myId); resolve(null); } }, 60000);
      pending.set(myId, (message) => { clearTimeout(timer); resolve(message); });
      send({ jsonrpc: '2.0', id: myId, method, params });
    });
    (async () => {
      await request('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'e2e', version: '1' } });
      send({ jsonrpc: '2.0', method: 'notifications/initialized' });
      const results = [];
      for (const [name, args] of calls) {
        const message = await request('tools/call', { name, arguments: args });
        let text = '';
        try { text = message.result.content[0].text; } catch { }
        results.push({ name, text, failed: !message || !message.result || !!message.result.isError });
      }
      child.kill();
      resolvePromise(results);
    })();
  });
}

// null means the answer did not state a count, so an unexpected format cannot be read as "clean".
function parseOpenCount(text) {
  if (/没有打开任何文档|没有打开的演示|: 无$|: 无\s*$/.test(text)) return 0;
  const match = /\((?:共)?(\d+)个\)/.exec(text) || /共(\d+)个/.exec(text);
  return match ? Number(match[1]) : null;
}

// --- main ----------------------------------------------------------------------------------

const opts = parseArgs(process.argv.slice(2));
const dshBin = resolveDshBin(opts.dshBin);
if (!dshBin) die('could not locate @deepseek-ai/dsh/lib/bin.js; pass --dsh-bin <path> or set DSH_BIN');
info('dsh entry: ' + dshBin);

if (opts.setup) {
  info('creating profile ' + opts.profile + ' from the headless template and installing ' + ROOT);
  const manifest = join(DSH_HOME, 'profiles', opts.profile, 'package.json');
  const created = runSync(process.execPath, [dshBin, '--profile', opts.profile, '--from-default-profile', 'headless', '--dump-config'], { timeout: 300000 });
  // Checked against the profile it writes rather than against its stdout: the composed tree prints
  // bundle ids with forward slashes, so a path-joined needle never matches.
  if (created.error || !existsSync(manifest)) {
    die('could not create profile ' + opts.profile + ': ' + (created.error ? created.error.message : String(created.stderr || created.stdout || '').slice(0, 400)));
  }
  const added = runSync(process.execPath, [dshBin, 'plugin', '--profile', opts.profile, 'add', ROOT], { timeout: 600000 });
  const manifestText = existsSync(manifest) ? readFileSync(manifest, 'utf8') : '';
  if (added.error || !manifestText.includes('dsh-plugin-wps-office-next')) {
    die('could not install the bundle into ' + opts.profile + ': ' + (added.error ? added.error.message : String(added.stderr || added.stdout || '').slice(0, 400)));
  }
  info('profile ready: ' + manifest);
}

const runId = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19) + '-' + process.pid;
const runDir = join(E2E_ROOT, runId);
const fixtureDir = join(runDir, 'fixture');
mkdirSync(fixtureDir, { recursive: true });
const workbookPath = join(fixtureDir, 'sales.xlsx');
const ordersPath = join(fixtureDir, 'orders.xlsx');
const documentPath = join(fixtureDir, '结论.docx');
info('run directory: ' + runDir);

// 1. fixture
const fixturePs1 = join(runDir, 'fixture.ps1');
writePowerShell(fixturePs1, fixtureScript());
const fixtureRun = runPowerShell(fixturePs1, ['-Path', workbookPath, '-OrdersPath', ordersPath]);
if (fixtureRun.error || fixtureRun.status !== 0) die('fixture generation failed: ' + (fixtureRun.error ? fixtureRun.error.message : String(fixtureRun.stderr || '').slice(0, 400)));
check('fixture workbook created', existsSync(workbookPath), workbookPath);
check('fixture order workbook created', existsSync(ordersPath), ordersPath);
if (existsSync(documentPath)) rmSync(documentPath, { force: true });

// 2. the task
const task = [
  TASK_PREFIX + ' ' + workbookPath + '：',
  '1) 打开它，按「地区」汇总「销售额」，把结果写入同一工作簿的新工作表「汇总」，两列（地区、销售额合计），按销售额从高到低排序；',
  '2) 在「汇总」表里，用这份汇总数据画一张簇状柱形图；',
  '3) 再新建一个 Word 文档，把各地区合计写成一段简短结论，保存为 ' + documentPath + '；',
  '4) 保存工作簿并关闭这两个文件，不要留下未保存的改动。',
  '',
  '还有一个文件 ' + ordersPath + '：',
  '5) 打开它，把「订单」表上的 A1:D' + ORDER_LAST_ROW + ' 变成一张真正的表（ListObject），表名用 Orders；',
  '6) 给「金额」列加一条条件格式：大于 500 的单元格标红；',
  '7) 让它可以直接打印：横向、A4、页脚居中写「第 &P 页 / 共 &N 页」，并且第 1 行在每一页都重复；',
  '8) 保存并关闭它，同样不要留下未保存的改动。',
  '请优先使用现成的 WPS 工具（可参考已注册的 WPS 技能），不要自己写 COM 脚本。完成后报告：两个工作簿的实际路径、各地区合计数值、订单表的表名、以及过程中遇到的任何报错。',
].join('\n');
// The run must exercise the advertised tool surface, so the optional PTC mode is cleared even if
// the caller's environment opted in; otherwise the model would reach WPS through run_code instead.
const runEnv = Object.assign({}, process.env, { DSH_PERMISSION_MODE: 'danger-full-access' });
delete runEnv.DSH_TOOLS_MODE;
const startedAt = Date.now();
const stdoutPath = join(runDir, 'run.stdout.txt');
const stderrPath = join(runDir, 'run.stderr.txt');
const stdoutFd = openSync(stdoutPath, 'w');
const stderrFd = openSync(stderrPath, 'w');
info('running headless task (limit ' + opts.timeoutSec + 's) ...');
const runResult = await new Promise((resolvePromise) => {
  const child = spawn(process.execPath, [dshBin, '--profile', opts.profile, task], {
    cwd: ROOT,
    windowsHide: true,
    env: runEnv,
    stdio: ['ignore', stdoutFd, stderrFd],
  });
  const timer = setTimeout(() => { info('  time limit reached, killing the run'); child.kill(); }, opts.timeoutSec * 1000);
  child.on('exit', (code, signal) => { clearTimeout(timer); resolvePromise({ code, signal: signal || '', ms: Date.now() - startedAt }); });
});
closeSync(stdoutFd);
closeSync(stderrFd);
const answer = readFileSync(stdoutPath, 'utf8').trim();
const reasoning = readFileSync(stderrPath, 'utf8');
check('headless task exited 0', runResult.code === 0, 'exit=' + runResult.code + ' elapsed=' + Math.round(runResult.ms / 1000) + 's');
check('run finished inside the time limit', runResult.ms < opts.timeoutSec * 1000, Math.round(runResult.ms / 1000) + 's');
check('produced a final answer', answer.length > 80, answer.split(/\r?\n/)[0].slice(0, 90));
info('reasoning bytes: ' + reasoning.length);

// 3. session trace
const session = findSession(startedAt);
check('found the run session log', !!session, session ? session.file : 'no session newer than the run start');
let trace = { lines: [], toolCalls: [], toolResults: [], pwshCommands: [], skillLoads: 0 };
if (session) {
  trace = summarizeTrace(session.records);
  writeUtf8Bom(join(runDir, 'trace.txt'), trace.lines.join('\n') + '\n');
  info('tool calls: ' + trace.toolCalls.length + ', skill loads: ' + trace.skillLoads);
}

// 4. artifacts, re-read with raw COM
const verifyPs1 = join(runDir, 'verify.ps1');
writePowerShell(verifyPs1, verifyScript());
const verifyRun = runPowerShell(verifyPs1, ['-Workbook', workbookPath, '-Document', documentPath, '-Orders', ordersPath, '-OrderLastRow', String(ORDER_LAST_ROW)]);
let findings = {};
try { findings = JSON.parse(String(verifyRun.stdout || '{}')); } catch { }
writeUtf8Bom(join(runDir, 'verify.json'), JSON.stringify(findings, null, 2) + '\n');
if (verifyRun.error || !findings.sheets) {
  check('independent COM verification ran', false, verifyRun.error ? verifyRun.error.message : 'no sheets; stderr: ' + String(verifyRun.stderr || '').slice(0, 200));
} else {
  check('independent COM verification ran', true, '');
  const sheets = findings.sheets || [];
  check('workbook has 明细 and 汇总', sheets.includes('明细') && sheets.includes('汇总'), sheets.join(', '));
  const rows = (findings.rows || []).map((row) => [String(row[0]), String(row[1])]);
  const expected = [['地区', '销售额合计']].concat(EXPECTED_ROWS.map((row) => [row[0], String(row[1])]));
  const rowsOk = rows.length === expected.length && rows.every((row, index) => row[0] === expected[index][0] && row[1].replace(/,/g, '') === expected[index][1]);
  check('汇总 holds the descending totals', rowsOk, JSON.stringify(rows));
  const charts = Array.isArray(findings.charts) ? findings.charts : findings.charts ? [findings.charts] : [];
  check('a clustered column chart is on 汇总', charts.length >= 1 && charts.some((chart) => Number(chart.type) === 51), JSON.stringify(charts));
  const paragraphs = Array.isArray(findings.document && findings.document.paragraphs) ? findings.document.paragraphs : [];
  const nonEmpty = paragraphs.filter((part) => part.text);
  const documentText = nonEmpty.map((part) => part.text).join(' ');
  // The task asks for the per-region totals; whether the grand total also appears, and whether the
  // model gives the text a heading style, are choices, so they are reported rather than asserted.
  check('document states every regional total', EXPECTED_ROWS.every((row) => documentText.includes(String(row[1]))), documentText.slice(0, 90));
  check('document is not empty', nonEmpty.length >= 1, nonEmpty.length + ' non-empty paragraph(s)');
  info('document styles: ' + (nonEmpty.map((part) => part.style).join(', ') || '(none)') + '; grand total stated: ' + documentText.includes(String(EXPECTED_TOTAL)) + '; characters: ' + ((findings.document && findings.document.characters) || 0));

  // Scenario 2 (P2-5): a second real Excel task in the same run - table + conditional format +
  // print setup. The checks are emitted unconditionally so the total count stays stable.
  const orders = findings.orders || null;
  const objects = orders && Array.isArray(orders.objects) ? orders.objects : orders && orders.objects ? [orders.objects] : [];
  check('independent COM verification read the order workbook', !!orders, orders ? 'ok' : (findings.ordersError || 'orders.xlsx missing'));
  check('订单 became a real table (ListObject)', !!orders && Number(orders.listCount) >= 1, objects.map((o) => o.name + '@' + o.range).join(', ') || 'none');
  check('the table is named Orders', objects.some((o) => o.name === 'Orders'), objects.map((o) => o.name).join(', ') || 'none');
  check('the table covers A1:D' + ORDER_LAST_ROW, objects.some((o) => String(o.range).replace(/\$/g, '').indexOf('A1:D' + ORDER_LAST_ROW) >= 0), objects.map((o) => o.range).join(', ') || 'none');
  check('a conditional format is on the amount column', !!orders && Number(orders.formatCount) >= 1, orders ? 'FormatConditions=' + orders.formatCount : 'n/a');
  check('the order sheet prints landscape A4', !!orders && Number(orders.orientation) === 2 && Number(orders.paperSize) === 9, orders ? 'orientation=' + orders.orientation + ' paper=' + orders.paperSize : 'n/a');
  check('the footer carries the page-number field', !!orders && String(orders.centerFooter || '').includes('&P'), orders ? String(orders.centerFooter || '') : 'n/a');
  check('print titles repeat the header row', !!orders && String(orders.printTitleRows || '').length > 0, orders ? String(orders.printTitleRows || '') : 'n/a');
}

// 5. behaviour: nothing left open, no defect markers, no hand-written COM
const leakCalls = await mcpCalls([
  ['wps_excel_get_open_workbooks', {}],
  ['wps_call', { tool: 'wps_word_get_open_documents', args: {} }],
  ['wps_call', { tool: 'wps_ppt_get_open_presentations', args: {} }],
]);
const leakCounts = leakCalls.map((entry) => parseOpenCount(entry.text));
const leakDetail = leakCalls.map((entry, index) => shortToolName(entry.name) + '=' + (leakCounts[index] === null ? 'unknown(' + oneLine(entry.text, 40) + ')' : leakCounts[index])).join(', ');
check('no document left open after the run', leakCounts.every((count) => count === null || count === 0), leakDetail);
check('leak probe understood every answer', leakCounts.every((count) => count !== null), leakDetail);

// Only operation results are scanned: wps_help returns documentation and skill loads return the
// skill text, and both legitimately quote these very strings while explaining how to handle them.
const defectHits = [];
let scannedResults = 0;
for (const result of trace.toolResults) {
  const name = shortToolName(result.name);
  if (!name.startsWith('wps_') || name === 'wps_help') continue;
  scannedResults++;
  for (const marker of DEFECT_MARKERS) if (result.text.includes(marker)) defectHits.push(name + ' -> ' + marker);
}
check('tool results contain no defect markers', defectHits.length === 0, defectHits.slice(0, 4).join(' | ') || scannedResults + ' operation result(s) scanned for ' + DEFECT_MARKERS.length + ' markers');

const comHits = [];
for (const command of trace.pwshCommands) {
  for (const marker of OWN_COM_MARKERS) if (command.includes(marker)) comHits.push(marker);
}
check('model never fell back to its own COM script', comHits.length === 0, comHits.slice(0, 3).join(', ') || trace.pwshCommands.length + ' pwsh command(s) scanned');

const wpsTools = expandToolNames(trace.toolCalls);
check('used wps_status', wpsTools.includes('wps_status'), wpsTools.slice(0, 3).join(', '));
check('drove WPS through the plugin', wpsTools.length >= 8, wpsTools.length + ' wps_* calls (batch and wps_call members included)');
check('loaded the WPS skills', trace.skillLoads >= 2, trace.skillLoads + ' skill load(s)');
check('created the Word document with a plugin tool', wpsTools.includes('wps_word_create_document'), wpsTools.filter((name) => name.includes('create')).join(', ') || 'none');

// 6. report
const failed = checks.filter((entry) => !entry.ok);
writeUtf8Bom(join(runDir, 'report.json'), JSON.stringify({
  runId, profile: opts.profile, dshBin, task, ordersPath, exitCode: runResult.code, elapsedMs: runResult.ms,
  sessionFile: session ? session.file : null,
  toolCalls: expandToolNames(trace.toolCalls),
  checks,
}, null, 2) + '\n');

console.log('');
console.log(failed.length === 0
  ? 'E2E OK (' + checks.length + ' checks) in ' + Math.round(runResult.ms / 1000) + 's'
  : 'E2E FAILED (' + failed.length + '/' + checks.length + ' checks) after ' + Math.round(runResult.ms / 1000) + 's');
console.log('evidence: ' + runDir);
if (failed.length) {
  for (const entry of failed) console.log('  - ' + entry.name + (entry.detail ? ': ' + entry.detail : ''));
  console.log('  inspect trace.txt / run.stderr.txt in the run directory');
}

if (opts.clean && failed.length === 0) {
  rmSync(fixtureDir, { recursive: true, force: true });
  info('removed the generated fixture and document (--clean)');
}
process.exit(failed.length === 0 ? 0 : 1);

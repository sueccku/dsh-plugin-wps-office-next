// FIXES 72: a range that is too big must fail fast and say so, instead of tying up the single-threaded
// COM host until the 60s timeout kills it. Also covers the span-wise insert and the long-path message.
// Needs a real WPS installation. Run: node test/range-limits.test.mjs
import { spawn } from 'node:child_process';
import { readFileSync } from 'node:fs';

const child = spawn(process.execPath, ['mcp/dist/index.js'], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
let buf = '';
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + '\n'); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); }); }
child.stdout.on('data', (d) => { buf += d.toString(); let i; while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on('data', () => {});
await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'range-limits', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });
let id = 10;
const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
async function call(name, args) {
  const started = Date.now();
  const r = await req(id++, 'tools/call', { name, arguments: args || {} });
  const isErr = !!(r.result && r.result.isError);
  const text = String((r.result && r.result.content && r.result.content[0].text) || '').replace(/\s+/g, ' ');
  return { isErr, text, ms: Date.now() - started };
}

await call('wps_excel_create_workbook', {});

// ---- read budget -------------------------------------------------------------------------------------
const wholeColumn = await call('wps_excel_read_range', { range: 'A:A' });
check('reading a whole column is refused', wholeColumn.isErr && /范围太大/.test(wholeColumn.text), wholeColumn.ms + 'ms ' + wholeColumn.text.slice(0, 110));
check('the refusal names the read budget and the action', /50000/.test(wholeColumn.text) && /read_range/.test(wholeColumn.text), wholeColumn.text.slice(0, 130));
check('it fails fast, not after the 60s timeout', wholeColumn.ms < 20000, wholeColumn.ms + 'ms');

await call('wps_excel_write_range', { range: 'A1:C2', data: [[1, 2, 3], [4, 5, 6]] });
const small = await call('wps_excel_read_range', { range: 'A1:C2' });
check('a small range still reads back', !small.isErr && /2行 x 3列/.test(small.text), small.text.slice(0, 90));
check('and keeps the 2D shape after the preallocation change', /1 \| 2 \| 3/.test(small.text) && /4 \| 5 \| 6/.test(small.text), small.text.slice(0, 140));

// ---- per-cell budget (clean_data) --------------------------------------------------------------------
const clean = await call('wps_excel_clean_data', { range: 'A:A', operations: ['trim'] });
check('cleaning a whole column is refused', clean.isErr && /范围太大/.test(clean.text) && /20000/.test(clean.text), clean.ms + 'ms ' + clean.text.slice(0, 120));

// ---- scan budget must not block ordinary searches ---------------------------------------------------
const find = await call('wps_excel_find_in_sheet', { searchText: 'zzz-no-such-value' });
check('an ordinary find is not blocked by the scan budget', !find.isErr, find.ms + 'ms ' + find.text.slice(0, 90));
const fit = await call('wps_excel_auto_fit', {});
check('an ordinary auto_fit is not blocked by the layout budget', !fit.isErr, fit.ms + 'ms ' + fit.text.slice(0, 90));

// ---- insert_columns: one span, not count COM round trips ---------------------------------------------
await call('wps_excel_create_workbook', {});
await call('wps_excel_write_range', { range: 'A1:D1', data: [['a', 'b', 'c', 'd']] });
const ins = await call('wps_excel_insert_columns', { column: 'B', count: 2 });
const shifted = await call('wps_excel_read_range', { range: 'A1:F1' });
check('inserting two columns reports and performs one span insert', !ins.isErr && /1行 x 6列/.test(shifted.text), ins.text.slice(0, 70) + ' | ' + shifted.text.slice(0, 90));
check('the data really shifted right by two', /\| b \| c \| d/.test(shifted.text), shifted.text.slice(0, 140));
await call('wps_excel_close_workbook', { save: false });

// ---- long path must explain itself -------------------------------------------------------------------
const long = 'C:\\' + 'x'.repeat(300) + '.png';
const longPath = await call('wps_excel_insert_excel_image', { path: long, cell: 'A1' });
check('a >260-character path explains the Windows limit', longPath.isErr && /260/.test(longPath.text) && /上限/.test(longPath.text), longPath.text.slice(0, 160));
check('and does not pretend it is simply missing', !/^.*not found: C:\\\\x+/.test(longPath.text), longPath.text.slice(0, 90));

// ---- static guard: the budget helper must stay exit-free and the tiers must stay --------------------
const bridge = readFileSync('mcp/scripts/wps-com.ps1', 'utf8');
check('the budget helper returns an error instead of exiting', /function Get-WpsRangeBudgetError/.test(bridge) && !/Assert-WpsRangeBudget/.test(bridge), 'no helper-level exit');
check('all four tiers are declared', /read = 50000/.test(bridge) && /scan = 200000/.test(bridge) && /layout = 200000/.test(bridge) && /percell = 20000/.test(bridge), 'read/scan/layout/percell');

try { child.kill(); } catch { }
const failed = results.filter((r) => !r.ok);
console.log('');
console.log(failed.length ? 'RANGE LIMIT TESTS FAILED (' + failed.length + '/' + results.length + ')' : 'RANGE LIMIT TESTS OK (' + results.length + ')');
process.exit(failed.length ? 1 : 0);

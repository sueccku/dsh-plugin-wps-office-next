// FIXES 73: bridge warnings must reach the model through FIRST-PARTY tools, not only through the
// pass-through surface (wps_call / wps_execute_method / wps_batch). Before this, every first-party
// handler built its own data object and dropped them, so the C7 ambiguity warning and every
// best-effort failure were invisible on ~260 tools. Needs a real WPS installation.
// Run: node test/warning-channel.test.mjs
import { spawn } from 'node:child_process';
import { readFileSync } from 'node:fs';
// FIXES 95（W1-4）：这个文件验证的是"警告有没有传到"，本身合理。但它有两处**前置条件**是自证的：
// "确实开了两个工作簿/两个文档"。前置条件要是不成立，后面整串警告断言就是白测 —— 补裸 COM 交叉验证。
// 正则闸门（要求纯数字）：com() 读失败返回空串，Number('') === 0 会让断言**假通过**。
import { com } from './lib/oracle.mjs';

const child = spawn(process.execPath, ['mcp/dist/index.js'], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
let buf = '';
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + '\n'); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); }); }
child.stdout.on('data', (d) => { buf += d.toString(); let i; while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on('data', () => {});
await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'warning-channel', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });
let id = 10;
const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
async function call(name, args) {
  const r = await req(id++, 'tools/call', { name, arguments: args || {} });
  const isErr = !!(r.result && r.result.isError);
  const text = String((r.result && r.result.content && r.result.content[0].text) || '').replace(/\s+/g, ' ');
  return { isErr, text };
}

// clean slate: close whatever is open, then open exactly two workbooks
for (let i = 0; i < 8; i++) { const r = await call('wps_excel_close_workbook', { save: false }); if (/失败/.test(r.text)) break; }
await call('wps_excel_create_workbook', {});
await call('wps_excel_create_workbook', {});
const open = await call('wps_excel_get_open_workbooks', {});
check('two workbooks are open for the ambiguity warning', /\(2个\)/.test(open.text), open.text.slice(0, 80));
// ---- 独立 oracle：前置条件经裸 COM 交叉验证 ----
const rawWb2 = String(com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application'); [int]$w.Workbooks.Count")).trim();
check('raw COM confirms the two-workbook precondition', /^\d+$/.test(rawWb2) && Number(rawWb2) === 2, 'raw=' + JSON.stringify(rawWb2));

// ---- first-party tool (read_range) -------------------------------------------------------------------
const read = await call('wps_excel_read_range', { range: 'A1' });
check('read_range still succeeds', !read.isErr, read.text.slice(0, 90));
check('the ambiguity warning reaches a first-party tool', /检测到 2 个打开的工作簿/.test(read.text), read.text.slice(0, 160));
check('and it arrives as an explicit note block', /注意（1 条/.test(read.text), read.text.slice(0, 160));

// ---- another first-party tool (write_range) ---------------------------------------------------------
const wrote = await call('wps_excel_write_range', { range: 'B1', data: [['x']] });
check('write_range carries the same warning', !wrote.isErr && /检测到 2 个打开的工作簿/.test(wrote.text) && /注意（1 条/.test(wrote.text), wrote.text.slice(0, 160));

// ---- pass-through surface must not be doubled -------------------------------------------------------
// wps_call of a first-party tool is still a first-party call, so the note is expected exactly once.
const viaCall = await call('wps_call', { tool: 'wps_excel_read_range', args: { range: 'A1' } });
check('wps_call of a first-party tool carries the note once', /检测到 2 个打开的工作簿/.test(viaCall.text) && (viaCall.text.match(/注意（/g) || []).length === 1, viaCall.text.slice(0, 120));
// wps_execute_method returns the bridge JSON verbatim, which already contains the warning string:
// the registry must recognise that and not append a second copy.
const verbatim = await call('wps_call', { tool: 'wps_execute_method', args: { method: 'getRangeData', params: { range: 'A1' } } });
check('the verbatim pass-through is not duplicated', /检测到 2 个打开的工作簿/.test(verbatim.text) && !/注意（/.test(verbatim.text), verbatim.text.slice(0, 130));

// ---- explicit target silences it --------------------------------------------------------------------
const named = await call('wps_excel_read_range', { range: 'A1', sheet: 'Sheet1' });
check('naming the sheet keeps the result clean', !named.isErr && !/检测到 2 个打开的工作簿/.test(named.text), named.text.slice(0, 120));

for (let i = 0; i < 4; i++) { const r = await call('wps_excel_close_workbook', { save: false }); if (/失败/.test(r.text)) break; }

// ---- Word now has a shared resolver too (FIXES 64 residual) ------------------------------------------
for (let i = 0; i < 8; i++) { const r = await call('wps_word_close_document', { save: false }); if (/失败/.test(r.text)) break; }
await call('wps_word_create_document', {});
await call('wps_word_create_document', {});
const wOpen = await call('wps_word_get_open_documents', {});
check('two documents are open for the Word warning', /2个/.test(wOpen.text), wOpen.text.slice(0, 80));
const rawDoc2 = String(com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application'); [int]$w.Documents.Count")).trim();
check('raw COM confirms the two-document precondition', /^\d+$/.test(rawDoc2) && Number(rawDoc2) === 2, 'raw=' + JSON.stringify(rawDoc2));
const wRead = await call('wps_word_get_document_text', {});
check('a Word first-party tool reports the ambiguity', /检测到 2 个打开的文档/.test(wRead.text) && /注意（1 条/.test(wRead.text), wRead.text.slice(0, 170));
for (let i = 0; i < 4; i++) { const r = await call('wps_word_close_document', { save: false }); if (/失败/.test(r.text)) break; }

// ---- static guard: the channel must stay wired -------------------------------------------------------
const client = readFileSync('mcp/src/client/wps-client.ts', 'utf8');
const registry = readFileSync('mcp/src/server/tool-registry.ts', 'utf8');
check('the client collects warnings from the bridge result', /collectToolWarnings\(result\?\.warnings\)/.test(client) && /collectToolWarnings\(\(result\?\.data/.test(client), 'both attachment points');
check('the registry attaches them to the tool result', /attachWarnings\(result, warnings\)/.test(registry) && /runWithWarningCollector/.test(registry), 'collector wired');

try { child.kill(); } catch { }
const failed = results.filter((r) => !r.ok);
console.log('');
console.log(failed.length ? 'WARNING CHANNEL TESTS FAILED (' + failed.length + '/' + results.length + ')' : 'WARNING CHANNEL TESTS OK (' + results.length + ')');
process.exit(failed.length ? 1 : 0);

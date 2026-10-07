// Verifies the Excel "missing half" tools (P2 wave 1) against a real WPS instance:
// sheet info / autofit x3 / wrap text / find-with-locations / named ranges read+delete.
// Run: node test/excel-missing-halves.test.mjs
import { spawn } from 'node:child_process';
import { resolve } from 'node:path';
// FIXES 95（W1-4）：命名区域的 set/delete 都只走插件自己的 get_named_ranges。
// 补一对**以文件为准**的核对：定义的名字真的写进了工作簿 XML，删除后**真的不在**了
// （"删掉了"如果只删了内存副本，用户下次打开文件会看到那个名字还在）。
import { zipScan } from './lib/oracle.mjs';

const child = spawn(process.execPath, ['mcp/dist/index.js'], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
let buf = '';
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + '\n'); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); }); }
child.stdout.on('data', (d) => { buf += d.toString(); let i; while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on('data', () => {});

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
function textOf(res) { try { return res.result.content[0].text; } catch { return ''; } }
function ok(res) { return !!(res && res.result && !res.result.isError); }
let id = 10;
const call = (name, args) => req(id++, 'tools/call', { name, arguments: args });
const viaCall = (tool, args) => call('wps_call', { tool, args });

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'mh', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

check('scratch workbook created', ok(await viaCall('wps_excel_create_workbook', {})), '');
await viaCall('wps_excel_write_range', { range: 'A1', data: [['地区', '产品', '备注'], ['华东', 'A', '一段很长很长的说明文字用于验证自动换行'], ['华南', 'A', '短'], ['华北', 'B', '华东地区重复出现']] });

const info = textOf(await call('wps_excel_get_sheet_info', {}));
check('sheet info reports the used range', /已用范围: \$?A\$?1:\$?C\$?4/.test(info), info.replace(/\n/g, ' | ').slice(0, 120));
check('sheet info lists the headers', info.includes('A=地区') && info.includes('C=备注'), info.replace(/\n/g, ' | ').slice(0, 120));
check('sheet info names the workbook and sheet', info.includes('工作簿:') && info.includes('全部工作表:'), info.replace(/\n/g, ' | ').slice(0, 120));

check('auto_fit succeeds', ok(await call('wps_excel_auto_fit', {})), textOf(await call('wps_excel_auto_fit', {})).slice(0, 60));
check('auto_fit_columns accepts a column', ok(await call('wps_excel_auto_fit_columns', { column: 'A' })), '');
check('auto_fit_rows accepts a row', ok(await call('wps_excel_auto_fit_rows', { row: 1 })), '');
check('auto_fit_columns rejects an unknown column instead of pretending', !ok(await call('wps_excel_auto_fit_columns', { column: 'ZZZ' })), textOf(await call('wps_excel_auto_fit_columns', { column: 'ZZZ' })).slice(0, 60));

const wrapped = textOf(await call('wps_excel_set_wrap_text', { range: 'A1:C4' }));
check('wrap text on reports the state', ok(await call('wps_excel_set_wrap_text', { range: 'A1:C4' })) && wrapped.includes('打开'), wrapped.slice(0, 60));
check('wrap text off reports the state', textOf(await call('wps_excel_set_wrap_text', { range: 'A1:C4', wrap: false })).includes('关闭'), '');
check('wrap text without range is rejected', !ok(await call('wps_excel_set_wrap_text', {})), '');

const found = textOf(await call('wps_excel_find_in_sheet', { searchText: '华东' }));
check('find_in_sheet returns cell addresses', found.includes('A2') && found.includes('C4'), found.replace(/\n/g, ' | ').slice(0, 120));
check('find_in_sheet reports a miss clearly', textOf(await call('wps_excel_find_in_sheet', { searchText: '不存在的词' })).includes('没有找到'), '');

check('named range set (existing tool)', ok(await viaCall('wps_excel_set_named_range', { name: '测试范围', range: 'A1:B2' })), '');
const names = textOf(await call('wps_excel_get_named_ranges', {}));
check('get_named_ranges lists it', names.includes('测试范围'), names.replace(/\n/g, ' | ').slice(0, 100));
// ---- 独立 oracle ①：名字真的写进了文件 ----
const MH1 = resolve('test/.artifacts/missing-halves-defined.xlsx');
await viaCall('wps_common_save_as', { outputPath: MH1, format: 'xlsx' });
const nameBefore = zipScan(MH1, '测试范围', 'xl/');
check('the named range really landed in the file', !!nameBefore && nameBefore.count >= 1, JSON.stringify(nameBefore));
check('delete_named_range succeeds', ok(await call('wps_excel_delete_named_range', { name: '测试范围' })), '');
const after = textOf(await call('wps_excel_get_named_ranges', {}));
check('the deleted name is gone', !after.includes('测试范围'), after.replace(/\n/g, ' | ').slice(0, 100));
// ---- 独立 oracle ②：删除后文件里**真的没有**这个名字了 ----
const MH2 = resolve('test/.artifacts/missing-halves-deleted.xlsx');
await viaCall('wps_common_save_as', { outputPath: MH2, format: 'xlsx' });
const nameAfter = zipScan(MH2, '测试范围', 'xl/');
check('the deleted name is really gone from the file', !!nameAfter && nameAfter.count === 0, JSON.stringify(nameAfter));

await viaCall('wps_excel_close_workbook', { save: false });
const open = textOf(await call('wps_excel_get_open_workbooks', {}));
check('no workbook left open', /\(0个\)/.test(open), open.replace(/\n/g, ' | ').slice(0, 60));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'EXCEL MISSING-HALVES TESTS OK (' + results.length + ')' : 'EXCEL MISSING-HALVES TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

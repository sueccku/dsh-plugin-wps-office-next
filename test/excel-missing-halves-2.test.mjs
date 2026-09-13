// Verifies the P2 wave-1 remainder (10 more Excel tools) against a real WPS instance:
// copy/clear formats, conditional-format read+remove, data-validation read+remove,
// refresh links, consolidate, force recalculate, group columns.
// Run: node test/excel-missing-halves-2.test.mjs
import { spawn } from 'node:child_process';

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
const info = async (cell) => textOf(await call('wps_excel_get_cell_info', { sheet: SHEET, cell }));
let SHEET = 'Sheet1';

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'mh2', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

check('scratch workbook created', ok(await viaCall('wps_excel_create_workbook', {})), '');
const sheetInfo = textOf(await viaCall('wps_excel_get_sheet_info', {}));
const sheetMatch = /当前工作表: ([^\n|]+)/.exec(sheetInfo);
if (sheetMatch) SHEET = sheetMatch[1].trim();
await viaCall('wps_excel_write_range', { range: 'A1', data: [[1], [2], [3], [4]] });

// ---- 格式刷 / 清除格式：用 get_cell_info 观察真实效果 ----
check('set_cell_format on A1', ok(await viaCall('wps_excel_set_cell_format', { range: 'A1', format: { bold: true } })), '');
check('A1 is bold before the brush', (await info('A1')).includes('"bold": true'), '');
const copied = textOf(await call('wps_excel_copy_format', { source: 'A1', target: 'A2:A4' }));
check('copy_format succeeds', ok(await call('wps_excel_copy_format', { source: 'A1', target: 'A2:A4' })) && copied.includes('已把'), copied.slice(0, 80));
const a3Before = await info('A3');
check('the brush really copied the bold', a3Before.includes('"bold": true'), a3Before.replace(/\s+/g, ' ').slice(0, 90));
check('copy_format without target is rejected', !ok(await call('wps_excel_copy_format', { source: 'A1' })), '');
const cleared = textOf(await call('wps_excel_clear_formats', { range: 'A2:A4' }));
const a3After = await info('A3');
check('clear_formats succeeds and the bold is gone', ok(await call('wps_excel_clear_formats', { range: 'A2:A4' })) && a3After.includes('"bold": false'), cleared.slice(0, 60) + ' | ' + a3After.replace(/\s+/g, ' ').slice(0, 70));
check('clear_formats without range is rejected', !ok(await call('wps_excel_clear_formats', {})), '');

// ---- 条件格式：读与删 ----
check('set_conditional_format (existing tool)', ok(await viaCall('wps_excel_set_conditional_format', { range: 'A1:A4', condition: '>2', format: 'red_fill' })), '');
const cfs = textOf(await call('wps_excel_get_conditional_formats', { range: 'A1:A4' }));
check('get_conditional_formats reports the rule', ok(await call('wps_excel_get_conditional_formats', { range: 'A1:A4' })) && /类型 \d/.test(cfs), cfs.replace(/\n/g, ' | ').slice(0, 100));
check('get_conditional_formats needs a range', !ok(await call('wps_excel_get_conditional_formats', {})), '');
check('remove_conditional_format by index', ok(await call('wps_excel_remove_conditional_format', { range: 'A1:A4', index: 1 })), '');
const cfsAfter = textOf(await call('wps_excel_get_conditional_formats', { range: 'A1:A4' }));
check('the rule is gone', cfsAfter.includes('没有条件格式'), cfsAfter.replace(/\n/g, ' | ').slice(0, 80));
await viaCall('wps_excel_set_conditional_format', { range: 'A1:A4', condition: '>1', format: 'bold' });
check('remove_conditional_format without index removes all', ok(await call('wps_excel_remove_conditional_format', { range: 'A1:A4' })) && textOf(await call('wps_excel_get_conditional_formats', { range: 'A1:A4' })).includes('没有条件格式'), '');

// ---- 数据验证：读与删 ----
check('set_data_validation (existing tool)', ok(await viaCall('wps_excel_set_data_validation', { range: 'B1:B4', type: 'list', formula: '甲,乙,丙' })), '');
const dvs = textOf(await call('wps_excel_get_data_validations', { range: 'B1:B4' }));
check('get_data_validations reports type and source', ok(await call('wps_excel_get_data_validations', { range: 'B1:B4' })) && dvs.includes('甲'), dvs.replace(/\n/g, ' | ').slice(0, 100));
check('get_data_validations needs a range', !ok(await call('wps_excel_get_data_validations', {})), '');
check('remove_data_validation succeeds', ok(await call('wps_excel_remove_data_validation', { range: 'B1:B4' })), '');
const dvsAfter = textOf(await call('wps_excel_get_data_validations', { range: 'B1:B4' }));
check('the validation is gone', !dvsAfter.includes('甲'), dvsAfter.replace(/\n/g, ' | ').slice(0, 80));

// ---- 重算 / 外部链接 / 列分组 ----
const calcAll = textOf(await call('wps_excel_calculate', { all: true }));
check('calculate(all) succeeds', ok(await call('wps_excel_calculate', { all: true })) && calcAll.includes('整个工作簿'), calcAll.slice(0, 60));
const calcOne = textOf(await call('wps_excel_calculate', {}));
check('calculate(sheet) succeeds', ok(await call('wps_excel_calculate', {})) && calcOne.includes('工作表'), calcOne.slice(0, 60));
const links = textOf(await call('wps_excel_refresh_links', {}));
check('refresh_links reports zero links honestly', ok(await call('wps_excel_refresh_links', {})) && links.includes('0 条'), links.slice(0, 60));
const grouped = textOf(await call('wps_excel_group_columns', { startColumn: 'B', endColumn: 'C' }));
check('group_columns succeeds', ok(await call('wps_excel_group_columns', { startColumn: 'B', endColumn: 'C' })) && grouped.includes('已分组列'), grouped.slice(0, 60));
check('group_columns without endColumn is rejected', !ok(await call('wps_excel_group_columns', { startColumn: 'B' })), '');

// ---- 合并计算：两张来源表汇总到目标格 ----
await viaCall('wps_excel_create_sheet', { name: 'S2' });
await viaCall('wps_excel_switch_sheet', { name: 'S2' });
await viaCall('wps_excel_write_range', { range: 'A1', data: [[10], [20]] });
await viaCall('wps_excel_switch_sheet', { name: SHEET });
const consSources = [SHEET + '!A1:A4', 'S2!A1:A2'];
const cons = textOf(await call('wps_excel_consolidate', { destination: 'D1', sources: consSources, function: 'sum' }));
check('consolidate succeeds', ok(await call('wps_excel_consolidate', { destination: 'D1', sources: consSources, function: 'sum' })) && cons.includes('合并到'), cons.replace(/\n/g, ' | ').slice(0, 90));
// Consolidate is positional: one result cell per source cell. D1 = 1+10, D2 = 2+20, and the
// non-overlapping A3/A4 stay as they are. (topRow=true would match by label and write nothing.)
const dcell = textOf(await viaCall('wps_excel_read_range', { range: 'D1:D4' }));
check('consolidate adds the overlapping cells (D1 = 1+10, D2 = 2+20)', dcell.includes('11') && dcell.includes('22'), dcell.replace(/\n/g, ' | ').slice(0, 110));
check('consolidate without sources is rejected', !ok(await call('wps_excel_consolidate', { destination: 'E1' })), '');

// ---- 分类汇总：有工具、此前零测试覆盖，且与 consolidate 同一个常量 bug ----
await viaCall('wps_excel_write_range', { range: 'A1', data: [['地区', '产品', '金额'], ['华东', 'A', 10], ['华东', 'B', 20], ['华南', 'A', 30]] });
const subRes = await call('wps_excel_subtotal', { range: 'A1:C4', groupBy: '1', function: 'sum', columns: [3] });
const sub = textOf(subRes);
check('subtotal succeeds', ok(subRes) && sub.includes('分类汇总完成'), sub.slice(0, 90));
const afterSub = textOf(await viaCall('wps_excel_get_sheet_info', {}));
check('subtotal really inserted summary rows', /已用范围: \$?A\$?1:\$?[A-Z]+\$?[5-9]/.test(afterSub), afterSub.replace(/\n/g, ' | ').slice(0, 130));

await viaCall('wps_excel_close_workbook', { save: false });
const open = textOf(await call('wps_excel_get_open_workbooks', {}));
check('no workbook left open', /\(0个\)/.test(open), open.replace(/\n/g, ' | ').slice(0, 60));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'EXCEL MISSING-HALVES-2 TESTS OK (' + results.length + ')' : 'EXCEL MISSING-HALVES-2 TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

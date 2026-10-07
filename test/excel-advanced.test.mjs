// Verifies the P2-4 advanced tools against a real WPS instance: pivot table listing/refresh/clear,
// workbook-wide refresh, goal seek, sparklines, chart labels and chart deletion.
// Run: node test/excel-advanced.test.mjs
import { spawn } from 'node:child_process';
import { resolve } from 'node:path';
// FIXES 95（W1-4）：这个文件全程只信插件自己的回报。现在加三个真值核对：
// ① 单变量求解后**直接读那两个单元格**（不是读工具说的"A10 = 25"）；
// ② 图表标题是不是真的写进了 chart part；③ 迷你图组在文件里的真实条数。
import { comGrid, zipScan } from './lib/oracle.mjs';

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

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'adv', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

check('scratch workbook created', ok(await viaCall('wps_excel_create_workbook', {})), '');
const sheetInfo = textOf(await viaCall('wps_excel_get_sheet_info', {}));
const sheetMatch = /当前工作表: ([^\n|]+)/.exec(sheetInfo);
const SHEET = sheetMatch ? sheetMatch[1].trim() : 'Sheet1';
await viaCall('wps_excel_write_range', { range: 'A1', data: [['Region', 'Product', 'Amount'], ['East', 'A', 10], ['East', 'B', 20], ['West', 'A', 30], ['West', 'B', 40], ['North', 'A', 50], ['North', 'B', 60]] });

// ---- 透视表：列表 / 刷新 / 清除 ----
const noPivotRes = await call('wps_excel_get_pivot_tables', { sheet: SHEET });
check('no pivot tables before creating one', textOf(noPivotRes).includes('没有透视表'), textOf(noPivotRes).replace(/\n/g, ' | ').slice(0, 90));
const pivotCreated = await viaCall('wps_excel_create_pivot_table', { sourceRange: 'A1:C7', destinationCell: 'F1', rowFields: ['Region'], valueFields: [{ field: 'Amount', aggregation: 'SUM' }] });
check('pivot table created (existing tool)', ok(pivotCreated), textOf(pivotCreated).replace(/\n/g, ' | ').slice(0, 90));
const listRes = await call('wps_excel_get_pivot_tables', { sheet: SHEET });
const listed = textOf(listRes);
check('get_pivot_tables finds it', ok(listRes) && listed.includes('共 1 张透视表'), listed.replace(/\n/g, ' | ').slice(0, 120));
check('the listing names the row field', listed.includes('行字段 Region'), listed.replace(/\n/g, ' | ').slice(0, 160));
const pivotName = /「([^」]+)」/.exec(listed)?.[1] || '数据透视表1';
const refreshRes = await call('wps_excel_refresh_pivot_tables', { sheet: SHEET });
check('refresh_pivot_tables refreshes the sheet', ok(refreshRes) && textOf(refreshRes).includes('已刷新 1 张透视表'), textOf(refreshRes).slice(0, 100));
const refreshOneRes = await call('wps_excel_refresh_pivot_tables', { sheet: SHEET, pivotTable: pivotName });
check('refresh_pivot_tables accepts a table name', ok(refreshOneRes) && textOf(refreshOneRes).includes(pivotName), textOf(refreshOneRes).slice(0, 110));
const refreshBadRes = await call('wps_excel_refresh_pivot_tables', { sheet: SHEET, pivotTable: '不存在的透视表' });
check('refreshing an unknown pivot table is rejected', !ok(refreshBadRes), textOf(refreshBadRes).slice(0, 70));
const refreshAllRes = await call('wps_excel_refresh_all_data', {});
check('refresh_all_data refreshes the workbook', ok(refreshAllRes) && textOf(refreshAllRes).includes('已刷新工作簿'), textOf(refreshAllRes).slice(0, 90));
const clearRes = await call('wps_excel_clear_pivot_table', { sheet: SHEET, pivotTable: pivotName });
check('clear_pivot_table clears the report and reports the leftover object honestly', ok(clearRes) && textOf(clearRes).includes('已清除透视表') && textOf(clearRes).includes('还有 1 个透视表对象'), textOf(clearRes).replace(/\n/g, ' | ').slice(0, 160));
const clearNoNameRes = await call('wps_excel_clear_pivot_table', { sheet: SHEET });
check('clear_pivot_table without a name is rejected', !ok(clearNoNameRes), textOf(clearNoNameRes).slice(0, 60));

// ---- 单变量求解 ----
await viaCall('wps_excel_write_range', { range: 'A10', data: [[2]] });
await viaCall('wps_excel_set_formula', { sheet: SHEET, range: 'B10', formula: '=A10*2' });
const seekRes = await call('wps_excel_goal_seek', { sheet: SHEET, cell: 'B10', goal: 50, changingCell: 'A10' });
const seek = textOf(seekRes);
check('goal_seek solves for the changing cell', ok(seekRes) && seek.includes('求出解') && seek.includes('A10 = 25'), seek.replace(/\n/g, ' | ').slice(0, 140));
check('goal_seek reports the achieved result', seek.includes('B10 现在 = 50'), seek.replace(/\n/g, ' | ').slice(0, 140));
// ---- 独立 oracle ①：直接读 A10 / B10 的真实值（工具说"求出解"，但格子到底变没变？）----
const solved = comGrid('A10:B10', SHEET);
check('the changing cell and the target really hold the solved values', !!solved && Number(solved[0][0]) === 25 && Number(solved[0][1]) === 50, JSON.stringify(solved));
const seekBadRes = await call('wps_excel_goal_seek', { sheet: SHEET, cell: 'B10', goal: 50 });
check('goal_seek without changingCell is rejected', !ok(seekBadRes), textOf(seekBadRes).slice(0, 60));

// ---- 迷你图 ----
const sparkRes = await call('wps_excel_add_sparkline', { sheet: SHEET, range: 'C2:C7', location: 'E2:E7' });
check('add_sparkline creates one group', ok(sparkRes) && textOf(sparkRes).includes('共 1 组'), textOf(sparkRes).slice(0, 110));
const sparkColRes = await call('wps_excel_add_sparkline', { sheet: SHEET, range: 'C2:C7', location: 'I2:I7', sparklineType: 'column', markers: true });
check('add_sparkline accepts the column type', ok(sparkColRes) && textOf(sparkColRes).includes('column'), textOf(sparkColRes).slice(0, 110));
const sparkBadRes = await call('wps_excel_add_sparkline', { sheet: SHEET, range: 'C2:C7', sparklineType: 'nope' });
check('add_sparkline without location is rejected', !ok(sparkBadRes), textOf(sparkBadRes).slice(0, 60));
const sparkClearRes = await call('wps_excel_clear_sparkline', { sheet: SHEET, location: 'E2:E7' });
check('clear_sparkline removes them', ok(sparkClearRes) && textOf(sparkClearRes).includes('剩余 0 组'), textOf(sparkClearRes).slice(0, 100));

// ---- 图表：标题与删除 ----
const chartRes = await viaCall('wps_excel_create_chart', { range: 'A1:B7', chartType: 'column_clustered', title: '临时' });
check('chart created (existing tool)', ok(chartRes), textOf(chartRes).replace(/\n/g, ' | ').slice(0, 90));
const labelsRes = await call('wps_excel_set_chart_labels', { sheet: SHEET, title: '月度销售', categoryAxisTitle: '地区', valueAxisTitle: '金额' });
const labels = textOf(labelsRes);
check('set_chart_labels sets the title and both axes', ok(labelsRes) && labels.includes('月度销售') && labels.includes('地区') && labels.includes('金额'), labels.replace(/\n/g, ' | ').slice(0, 170));
check('the applied list names all three', labels.includes('已更新') && labels.includes('valueAxisTitle'), labels.split('\n')[0].slice(0, 110));
// ---- 独立 oracle ②：标题是不是真的写进了 chart part ----
const ADV1 = resolve('test/.artifacts/excel-advanced-1.xlsx');
check('workbook saved for independent verification (1)', ok(await viaCall('wps_common_save_as', { outputPath: ADV1, format: 'xlsx' })), '');
const titleScan = zipScan(ADV1, '月度销售', 'xl/charts/');
check('file says the chart carries the new title', !!titleScan && titleScan.count >= 1, JSON.stringify(titleScan));
const axisScan = zipScan(ADV1, '地区', 'xl/charts/');
check('file says the category axis title landed too', !!axisScan && axisScan.count >= 1, JSON.stringify(axisScan));
const delRes = await call('wps_excel_delete_chart', { sheet: SHEET });
check('delete_chart removes the only chart', ok(delRes) && textOf(delRes).includes('还剩 0 张'), textOf(delRes).slice(0, 110));
const delAgainRes = await call('wps_excel_delete_chart', { sheet: SHEET });
check('deleting again is rejected honestly', !ok(delAgainRes) && textOf(delAgainRes).includes('chart not found'), textOf(delAgainRes).slice(0, 90));

// ---- 独立 oracle ③：文件里到底还剩几组迷你图（加了两组、清了一组 → 应剩 1 组）----
const ADV2 = resolve('test/.artifacts/excel-advanced-2.xlsx');
check('workbook saved for independent verification (2)', ok(await viaCall('wps_common_save_as', { outputPath: ADV2, format: 'xlsx' })), '');
const sparkScan = zipScan(ADV2, '<x14:sparklineGroup ', 'xl/worksheets/');
check('file says exactly one sparkline group survived the clear', !!sparkScan && sparkScan.count === 1, JSON.stringify(sparkScan));

await viaCall('wps_excel_close_workbook', { save: false });
const open = textOf(await call('wps_excel_get_open_workbooks', {}));
check('no workbook left open', /\(0个\)/.test(open), open.replace(/\n/g, ' | ').slice(0, 60));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'EXCEL ADVANCED TESTS OK (' + results.length + ')' : 'EXCEL ADVANCED TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

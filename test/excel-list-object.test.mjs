// Verifies the P2-2 ListObject tools against a real WPS instance: build a table from a range, read
// its structure, add/delete rows, totals row (SUBTOTAL + structured reference), style/name options,
// resize, and unlist back to a plain range.
// Run: node test/excel-list-object.test.mjs
import { spawn } from 'node:child_process';
import { resolve } from 'node:path';
// FIXES 95（W1-4）：这个文件全程只信插件自己的回报。现在在三个关键节点**存盘后用 openpyxl 读文件里的表定义**
// （表名 / 范围 / 列名 / 表格样式），并在 unlist 之后确认文件里**真的不再有表**、而数据还在。
import { xlsx, xlsxTables } from './lib/oracle.mjs';

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

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'lo', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

check('scratch workbook created', ok(await viaCall('wps_excel_create_workbook', {})), '');
const sheetInfo = textOf(await viaCall('wps_excel_get_sheet_info', {}));
const sheetMatch = /当前工作表: ([^\n|]+)/.exec(sheetInfo);
const SHEET = sheetMatch ? sheetMatch[1].trim() : 'Sheet1';
await viaCall('wps_excel_write_range', { range: 'A1', data: [['Region', 'Product', 'Amount'], ['East', 'A', 10], ['East', 'B', 20], ['West', 'A', 30]] });

const noneRes = await call('wps_excel_get_list_objects', {});
check('no tables before the first create', textOf(noneRes).includes('没有表'), textOf(noneRes).replace(/\n/g, ' | ').slice(0, 80));

// ---- 建表与读结构 ----
const createdRes = await call('wps_excel_create_list_object', { range: 'A1:C4', name: 'Sales' });
const created = textOf(createdRes);
check('create_list_object builds the table', ok(createdRes) && created.includes('Sales'), created.replace(/\n/g, ' | ').slice(0, 120));
check('the table reports its structure', /3 列 \/ 3 行/.test(created) && created.includes('Region, Product, Amount'), created.replace(/\n/g, ' | ').slice(0, 150));
const noRangeRes = await call('wps_excel_create_list_object', {});
check('create_list_object without range is rejected', !ok(noRangeRes), '');

const listedRes = await call('wps_excel_get_list_objects', {});
const listed = textOf(listedRes);
check('get_list_objects finds it and offers structured references', listed.includes('Sales') && listed.includes('Sales[Amount]'), listed.replace(/\n/g, ' | ').slice(0, 170));
check('the reported range is the real one', /\$?A\$?1:\$?C\$?4/.test(listed), listed.replace(/\n/g, ' | ').slice(0, 110));

// ---- 加行 / 删行 ----
const addedRes = await call('wps_excel_add_list_row', { table: 'Sales', values: ['West', 'C', 40] });
const added = textOf(addedRes);
check('add_list_row appends and reports 4 rows', ok(addedRes) && /3 列 \/ 4 行/.test(added), added.replace(/\n/g, ' | ').slice(0, 120));
const row5 = textOf(await viaCall('wps_excel_read_range', { range: 'A5:C5' }));
check('the appended row really landed in the sheet', row5.includes('West') && row5.includes('40'), row5.replace(/\n/g, ' | ').slice(0, 90));

// ---- 独立 oracle ①：建表 + 加行之后，读磁盘上的 xlsx ----
const LO1 = resolve('test/.artifacts/list-object-1.xlsx');
check('workbook saved for independent verification (1)', ok(await viaCall('wps_common_save_as', { outputPath: LO1, format: 'xlsx' })), '');
const t1 = xlsxTables(LO1);
check('oracle can read the table out of the file', !!t1 && !!t1.Sales, JSON.stringify(Object.keys(t1 || {})));
if (t1 && t1.Sales) {
  check('file says the table spans A1:C5 after the append', String(t1.Sales.ref).replace(/\$/g, '') === 'A1:C5', JSON.stringify(t1.Sales.ref));
  check('file says the columns are the header row', JSON.stringify(t1.Sales.columns) === JSON.stringify(['Region', 'Product', 'Amount']), JSON.stringify(t1.Sales.columns));
}
const missingTableRes = await call('wps_excel_add_list_row', { table: 'Nope', values: ['x'] });
check('add_list_row on a missing table is rejected', !ok(missingTableRes), '');

const deletedRes = await call('wps_excel_delete_list_row', { table: 'Sales', rowIndex: 4 });
const deleted = textOf(deletedRes);
check('delete_list_row removes it and reports the shrunken table', ok(deletedRes) && /3 列 \/ 3 行/.test(deleted), deleted.replace(/\n/g, ' | ').slice(0, 120));
const oobRes = await call('wps_excel_delete_list_row', { table: 'Sales', rowIndex: 99 });
check('an out-of-range row index is rejected', !ok(oobRes), textOf(oobRes).slice(0, 60));

// ---- 总计行：SUBTOTAL + 结构化引用 ----
const totalsRes = await call('wps_excel_set_list_object_totals', { table: 'Sales', column: 'Amount', function: 'sum' });
const totals = textOf(totalsRes);
check('set_list_object_totals turns the totals row on', ok(totalsRes) && totals.includes('总计行 开'), totals.replace(/\n/g, ' | ').slice(0, 130));
const formulaRes = await viaCall('wps_excel_get_formula', { sheet: SHEET, cell: 'C5' });
const formula = textOf(formulaRes);
check('the totals cell holds a SUBTOTAL over the structured reference', formula.includes('SUBTOTAL(109') && formula.includes('[Amount]'), formula.replace(/\n/g, ' | ').slice(0, 110));
const countedRes = await call('wps_excel_set_list_object_totals', { table: 'Sales', column: 'Amount', function: 'count' });
const formula2 = textOf(await viaCall('wps_excel_get_formula', { sheet: SHEET, cell: 'C5' }));
check('changing the function rewrites the formula', ok(countedRes) && formula2.includes('SUBTOTAL(103'), formula2.replace(/\n/g, ' | ').slice(0, 110));
const totalsOffRes = await call('wps_excel_set_list_object_totals', { table: 'Sales', show: false });
check('the totals row can be turned off again', ok(totalsOffRes) && textOf(totalsOffRes).includes('总计行 关'), textOf(totalsOffRes).replace(/\n/g, ' | ').slice(0, 110));

// ---- 样式 / 改名 / 调整范围 / 转回区域 ----
const styledRes = await call('wps_excel_update_list_object', { table: 'Sales', tableStyle: 'TableStyleMedium2', showAutoFilter: false });
check('update_list_object changes the style', ok(styledRes) && textOf(styledRes).includes('TableStyleMedium2'), textOf(styledRes).replace(/\n/g, ' | ').slice(0, 130));
const renamedRes = await call('wps_excel_update_list_object', { table: 'Sales', name: 'SalesData' });
check('update_list_object renames the table', ok(renamedRes) && textOf(renamedRes).includes('SalesData'), textOf(renamedRes).replace(/\n/g, ' | ').slice(0, 110));
const oldNameRes = await call('wps_excel_update_list_object', { table: 'Sales', tableStyle: 'TableStyleLight1' });
check('the old name no longer resolves', !ok(oldNameRes), '');

const resizedRes = await call('wps_excel_resize_list_object', { table: 'SalesData', range: 'A1:C6' });
check('resize_list_object grows the table', ok(resizedRes) && /3 列 \/ 5 行/.test(textOf(resizedRes)), textOf(resizedRes).replace(/\n/g, ' | ').slice(0, 130));
const resizeNoRange = await call('wps_excel_resize_list_object', { table: 'SalesData' });
check('resize_list_object without range is rejected', !ok(resizeNoRange), '');

// ---- 独立 oracle ②：改名 + 调整范围 + 样式之后，文件里到底是什么 ----
const LO2 = resolve('test/.artifacts/list-object-2.xlsx');
check('workbook saved for independent verification (2)', ok(await viaCall('wps_common_save_as', { outputPath: LO2, format: 'xlsx' })), '');
const t2 = xlsxTables(LO2);
check('file says the table was renamed (old name gone)', !!t2 && !!t2.SalesData && !t2.Sales, JSON.stringify(Object.keys(t2 || {})));
if (t2 && t2.SalesData) {
  check('file says the table now spans A1:C6', String(t2.SalesData.ref).replace(/\$/g, '') === 'A1:C6', JSON.stringify(t2.SalesData.ref));
  check('file carries the applied table style', /TableStyleMedium2/.test(String(t2.SalesData.style)), JSON.stringify(t2.SalesData.style));
}

const unlistedRes = await call('wps_excel_unlist_list_object', { table: 'SalesData' });
check('unlist_list_object converts back to a range', ok(unlistedRes) && textOf(unlistedRes).includes('转回普通区域'), textOf(unlistedRes).replace(/\n/g, ' | ').slice(0, 130));
const afterRes = await call('wps_excel_get_list_objects', {});
check('no table remains after unlist', textOf(afterRes).includes('没有表'), textOf(afterRes).replace(/\n/g, ' | ').slice(0, 90));
const kept = textOf(await viaCall('wps_excel_read_range', { range: 'A1:C4' }));
check('unlist kept the data', kept.includes('East') && kept.includes('West'), kept.replace(/\n/g, ' | ').slice(0, 110));

// ---- 独立 oracle ③：转回区域之后 —— 文件里**真的不再有表**，而数据一格不少 ----
const LO3 = resolve('test/.artifacts/list-object-3.xlsx');
check('workbook saved for independent verification (3)', ok(await viaCall('wps_common_save_as', { outputPath: LO3, format: 'xlsx' })), '');
const t3 = xlsxTables(LO3);
check('file says no table remains after unlist', !!t3 && Object.keys(t3).length === 0, JSON.stringify(t3));
const v3 = xlsx(LO3, 'A1:C4');
check('file still carries the data after unlist', !!v3 && JSON.stringify(v3.values).includes('East') && JSON.stringify(v3.values).includes('West'), JSON.stringify(v3 && v3.values));

await viaCall('wps_excel_close_workbook', { save: false });
const open = textOf(await call('wps_excel_get_open_workbooks', {}));
check('no workbook left open', /\(0个\)/.test(open), open.replace(/\n/g, ' | ').slice(0, 60));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'EXCEL LIST-OBJECT TESTS OK (' + results.length + ')' : 'EXCEL LIST-OBJECT TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

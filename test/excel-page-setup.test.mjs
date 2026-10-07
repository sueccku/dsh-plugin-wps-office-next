// Verifies the P2-3 page/print/appearance tools against a real WPS instance: page setup read-back,
// orientation/paper/margins/scale, print titles, header/footer, sheet visibility and tab colour,
// outline levels, page-break reset, and formula auditing (precedents/dependents).
// Run: node test/excel-page-setup.test.mjs
import { spawn } from 'node:child_process';
import { resolve } from 'node:path';
// FIXES 95（W1-4）：这个文件 25 条断言**全部**在验证插件自己的回读文案，而打印设置恰恰是
// "回报写了 ≠ 文件里真是那样"的重灾区。现在存盘后用 openpyxl 直接读文件本身做独立核对。
import { xlsxProps } from './lib/oracle.mjs';

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

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'ps', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

check('scratch workbook created', ok(await viaCall('wps_excel_create_workbook', {})), '');
const sheetInfo = textOf(await viaCall('wps_excel_get_sheet_info', {}));
const sheetMatch = /当前工作表: ([^\n|]+)/.exec(sheetInfo);
const SHEET = sheetMatch ? sheetMatch[1].trim() : 'Sheet1';
await viaCall('wps_excel_write_range', { range: 'A1', data: [['Region', 'Amount'], ['East', 10], ['West', 20], ['North', 30], ['South', 40], ['All', 50]] });
check('second sheet created', ok(await viaCall('wps_excel_create_sheet', { name: 'Extra' })), '');

// ---- 读设置 ----
const initialRes = await call('wps_excel_get_sheet_settings', { sheet: SHEET });
const initial = textOf(initialRes);
check('get_sheet_settings returns a readable snapshot', ok(initialRes) && initial.includes('页面设置') && initial.includes('页边距') && initial.includes('打印区域'), initial.replace(/\n/g, ' | ').slice(0, 150));

// ---- 页面设置 ----
const landRes = await call('wps_excel_set_sheet_page_setup', { sheet: SHEET, orientation: 'landscape', paperSize: 'A3' });
const land = textOf(landRes);
check('orientation and paper size are applied and read back', ok(landRes) && land.includes('landscape') && land.includes('A3'), land.replace(/\n/g, ' | ').slice(0, 150));
check('the applied list names what changed', land.includes('已应用') && land.includes('orientation'), land.split('\n')[0].slice(0, 110));
const marginRes = await call('wps_excel_set_sheet_page_setup', { sheet: SHEET, topMargin: 36, leftMargin: 36, centerHorizontally: true });
const margin = textOf(marginRes);
check('margins and centring land', ok(marginRes) && margin.includes('上 36') && margin.includes('左 36') && margin.includes('水平 是'), margin.replace(/\n/g, ' | ').slice(0, 170));
const fitRes = await call('wps_excel_set_sheet_page_setup', { sheet: SHEET, fitToPagesWide: 1 });
check('fit-to-page switches the scale mode', ok(fitRes) && textOf(fitRes).includes('按页适配'), textOf(fitRes).replace(/\n/g, ' | ').slice(0, 130));
const zoomRes = await call('wps_excel_set_sheet_page_setup', { sheet: SHEET, zoom: 90 });
check('zoom switches back to a percentage', ok(zoomRes) && textOf(zoomRes).includes('90%'), textOf(zoomRes).replace(/\n/g, ' | ').slice(0, 130));
const badPaperRes = await call('wps_excel_set_sheet_page_setup', { sheet: SHEET, paperSize: 'nonsense' });
check('an unknown paper size is rejected, not ignored', !ok(badPaperRes) && textOf(badPaperRes).includes('unknown paperSize'), textOf(badPaperRes).slice(0, 90));

// ---- 打印标题 / 页眉页脚 ----
const titleRes = await call('wps_excel_set_sheet_print_titles', { sheet: SHEET, printTitleRows: '$1:$1' });
check('print titles are set and read back', ok(titleRes) && textOf(titleRes).includes('行 $1:$1'), textOf(titleRes).replace(/\n/g, ' | ').slice(0, 140));
const hfRes = await call('wps_excel_set_sheet_header_footer', { sheet: SHEET, leftHeader: '&F', centerFooter: '第 &P 页 / 共 &N 页' });
const hf = textOf(hfRes);
check('header and footer are set and read back', ok(hfRes) && hf.includes('页脚') && hf.includes('&P'), hf.replace(/\n/g, ' | ').slice(0, 180));

// ---- 外观：标签色与可见性 ----
const colorRes = await call('wps_excel_set_sheet_appearance', { sheet: SHEET, tabColor: '#FF9900' });
const color = textOf(colorRes);
check('tab colour is applied (BGR 39423 for #FF9900)', ok(colorRes) && color.includes('标签色: 39423'), color.split('\n').slice(-1)[0].slice(0, 110));
const clearRes = await call('wps_excel_set_sheet_appearance', { sheet: SHEET, tabColor: '' });
check('an empty colour restores the default', ok(clearRes) && textOf(clearRes).includes('标签色: (默认)'), textOf(clearRes).split('\n').slice(-1)[0].slice(0, 110));
const hideRes = await call('wps_excel_set_sheet_appearance', { sheet: 'Extra', visible: 'hidden' });
check('a sheet can be hidden', ok(hideRes) && textOf(hideRes).includes('工作表: 隐藏'), textOf(hideRes).split('\n').slice(-1)[0].slice(0, 100));
const veryHiddenRes = await call('wps_excel_set_sheet_appearance', { sheet: 'Extra', visible: 'veryHidden' });
check('a sheet can be very hidden', ok(veryHiddenRes) && textOf(veryHiddenRes).includes('空格') === false && textOf(veryHiddenRes).includes('工作表: 深度隐藏'), textOf(veryHiddenRes).split('\n').slice(-1)[0].slice(0, 100));
const showRes = await call('wps_excel_set_sheet_appearance', { sheet: 'Extra', visible: 'visible' });
check('a hidden sheet can be shown again', ok(showRes) && textOf(showRes).includes('工作表: 可见'), textOf(showRes).split('\n').slice(-1)[0].slice(0, 100));
const badVisibleRes = await call('wps_excel_set_sheet_appearance', { sheet: 'Extra', visible: 'sometimes' });
check('an unknown visibility value is rejected', !ok(badVisibleRes) && textOf(badVisibleRes).includes('unknown visible'), textOf(badVisibleRes).slice(0, 80));

// ---- 分级显示与分页符 ----
await viaCall('wps_excel_group_rows', { startRow: 2, endRow: 4 });
const outlineRes = await call('wps_excel_set_outline_levels', { sheet: SHEET, rowLevels: 1, summaryRow: 'below' });
check('outline levels are accepted and the summary position is read back', ok(outlineRes) && textOf(outlineRes).includes('汇总行/列位置: 下'), textOf(outlineRes).slice(0, 120));
const breaksRes = await call('wps_excel_reset_page_breaks', { sheet: SHEET });
check('page breaks reset reports the remaining count', ok(breaksRes) && textOf(breaksRes).includes('剩余手动分页符 0 条'), textOf(breaksRes).slice(0, 100));

// ---- 公式审计 ----
await viaCall('wps_excel_set_formula', { range: 'B7', formula: '=B2+B3' });
const auditRes = await call('wps_excel_get_formula_audit', { sheet: SHEET, cell: 'B7' });
const audit = textOf(auditRes);
check('audit finds the precedents of a formula', ok(auditRes) && audit.includes('引用来源') && audit.includes('2 处'), audit.replace(/\n/g, ' | ').slice(0, 150));
const depRes = await call('wps_excel_get_formula_audit', { sheet: SHEET, cell: 'B2' });
check('audit finds the dependents of a cell', ok(depRes) && textOf(depRes).includes('被引用') && /被引用（dependents）: 1 处/.test(textOf(depRes)), textOf(depRes).replace(/\n/g, ' | ').slice(0, 150));
const constRes = await call('wps_excel_get_formula_audit', { sheet: SHEET, cell: 'A1' });
check('a constant cell is reported honestly', ok(constRes) && textOf(constRes).includes('不是公式'), textOf(constRes).replace(/\n/g, ' | ').slice(0, 110));
const noCellRes = await call('wps_excel_get_formula_audit', {});
check('audit without a cell is rejected', !ok(noCellRes), textOf(noCellRes).slice(0, 60));

// ---- 独立 oracle：把设置**存进文件**再读回来（W1-4）------------------------------------------
// 不以插件的回读为准，改以**磁盘上的 xlsx**为准：方向 / 纸张 / 页边距 / 居中 / 打印标题 /
// 页眉页脚 / 标签色 / 工作表可见性。单位提醒：openpyxl 的页边距是**英寸**（36 磅 = 0.5），paperSize 是**数字**（A3 = 8）。
const OUT = resolve('test/.artifacts/page-setup.xlsx');
const saved = await viaCall('wps_common_save_as', { outputPath: OUT, format: 'xlsx' });
check('workbook saved for independent verification', ok(saved), textOf(saved).replace(/\n/g, ' ').slice(0, 80));
const props = xlsxProps(OUT);
check('oracle can read the saved workbook', !!props && !!props[SHEET] && !!props.Extra, JSON.stringify(Object.keys(props || {})));
if (props && props[SHEET]) {
  const p = props[SHEET];
  check('file says landscape (not just the report)', p.orientation === 'landscape', JSON.stringify(p.orientation));
  check('file says A3 paper (openpyxl paperSize 8)', Number(p.paperSize) === 8, String(p.paperSize));
  check('file says 36pt margins = 0.5 inch', Math.abs(Number(p.marginTop) - 0.5) < 0.01 && Math.abs(Number(p.marginLeft) - 0.5) < 0.01, 'top=' + p.marginTop + ' left=' + p.marginLeft);
  check('file says horizontally centred', p.horizontalCentered === true, String(p.horizontalCentered));
  check('file says print titles are row 1', String(p.printTitleRows || '').replace(/\$/g, '') === '1:1', JSON.stringify(p.printTitleRows));
  check('file carries the header/footer fields', /&F/.test(String(p.headerLeft)) && /&P/.test(String(p.footerCenter)), JSON.stringify({ header: p.headerLeft, footer: p.footerCenter }));
  check('file says the last scale mode won (zoom 90, not fit-to-page)', Number(p.scale) === 90 && p.fitToPage === false, 'scale=' + p.scale + ' fitToPage=' + p.fitToPage);
  check('file says the tab colour was really cleared', !p.tabColor || p.tabColor === '00000000', JSON.stringify(p.tabColor));
}
if (props && props.Extra) {
  check('file says the Extra sheet is visible again', props.Extra.sheetState === 'visible', JSON.stringify(props.Extra.sheetState));
}

await viaCall('wps_excel_close_workbook', { save: false });
const open = textOf(await call('wps_excel_get_open_workbooks', {}));
check('no workbook left open', /\(0个\)/.test(open), open.replace(/\n/g, ' | ').slice(0, 60));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'EXCEL PAGE-SETUP TESTS OK (' + results.length + ')' : 'EXCEL PAGE-SETUP TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

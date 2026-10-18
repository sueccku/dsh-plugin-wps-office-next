// S6 走查回归（真机）：sheet 名字对不上时，桥必须回一条能照着做的中文错误
// ——「哪本工作簿、缺哪张表、这本里现有哪几张」，而不是把裸 COM
// 「Invalid index. (异常来自 HRESULT:0x8002000B (DISP_E_BADINDEX))」丢给模型。
// 真机依据：S6 attempt 1 / attempt 2 各实测到一次（FIXES 111；走查报告已随 2026-10-18 文档清理删除）。
//
// 反向取证（先红后绿）：把 host 的两处调用还原成 `$wb.Sheets.Item(...)`
// （python test/.artifacts/scenario-s6/tools/make_prefix_host.py），再用环境变量
// S6_PREFIX_HOST 指向那份副本的 wps-com-host.ps1 跑本文件，本文件必须变红。
//
// Run: node test/excel-sheet-resolve.test.mjs
import { startServer, createReport } from './lib/harness.mjs';

const srv = startServer({
  clientName: 'sheet-resolve',
  env: process.env.S6_PREFIX_HOST ? { WPS_OFFICE_HOST_SCRIPT: process.env.S6_PREFIX_HOST } : {},
});
const { call, text, ok, stop } = srv;
const { check, finish } = createReport('EXCEL SHEET-RESOLVE TESTS');
const viaCall = (tool, args) => call('wps_call', { tool, args });
const flat = (s) => String(s).replace(/\s+/g, ' ').slice(0, 160);

const created = await viaCall('wps_excel_create_workbook', {});
const infoTxt = text(await call('wps_excel_get_sheet_info', {}));
const wbName = (infoTxt.match(/工作簿[:：]\s*([^\s|]+)/) || [])[1] || null;
check('scratch workbook created', ok(created) && !!wbName, flat(infoTxt));
check('marker written to the first sheet', ok(await call('wps_excel_write_range', { range: 'A1', data: [['标记']] })), '');
check('second sheet "Data" created', ok(await viaCall('wps_excel_create_sheet', { name: 'Data' })), '');
check('marker written to sheet Data', ok(await call('wps_excel_write_range', { range: 'A1', sheet: 'Data', data: [['数据']] })), '');

// ① 名字不存在：中文诊断 + 现有表清单，不再漏裸 COM 码
const missRes = await call('wps_excel_read_range', { range: 'A1:B2', sheet: '费用明细' });
const miss = text(missRes);
check('unknown sheet name is rejected (no silent read)', !ok(missRes), flat(miss));
check('error names the missing sheet', miss.includes('费用明细'), '');
check('error names the workbook it looked in', miss.includes('工作簿'), '');
check('error lists the sheets that do exist', miss.includes('现有工作表') && miss.includes('Data'), '');
check('error no longer leaks the raw COM code', !/Invalid index|DISP_E_BADINDEX/i.test(miss), '');
check('error carries a next step (error contract)', miss.includes('下一步'), '');

// ② 名字对得上：大小写不敏感仍然有效
const caseRes = await call('wps_excel_read_range', { range: 'A1:B2', sheet: 'data' });
check('sheet lookup stays case-insensitive', ok(caseRes) && text(caseRes).includes('数据'), flat(text(caseRes)));

// ③ 序号照旧可用（不能被新解析器吃掉）
const idx2 = await call('wps_excel_read_range', { range: 'A1:B2', sheet: '2' });
check('numeric sheet index still works', ok(idx2) && text(idx2).includes('数据'), flat(text(idx2)));
const idx1 = await call('wps_excel_read_range', { range: 'A1:B2', sheet: '1' });
check('index 1 reads the first sheet', ok(idx1) && text(idx1).includes('标记'), flat(text(idx1)));
const idxOut = await call('wps_excel_read_range', { range: 'A1:B2', sheet: '9' });
check('an out-of-range index is rejected with the same clear message',
  !ok(idxOut) && text(idxOut).includes('现有工作表'), flat(text(idxOut)));

await viaCall('wps_excel_close_workbook', { name: wbName, save: false });
const open = text(await call('wps_excel_get_open_workbooks', {}));
check('our scratch workbook is closed', !!wbName && !open.includes(wbName), flat(open));

stop();
finish();

// D4 回归（真机）：新进广告面的 9 个工具必须「**直接调用**就能用」。
//
// 背景：S6 真机走查里模型直接调 `wps_excel_freeze_panes` 拿到 `unknown tool`（不在广告面），
// 白花一次调用并重试。D4 把这些高频隐藏工具提到广告面之后，本文件是那条链路的回归：
//   ① tools/list（standard 档）里直接列出它们，且与 spec/advertised.json 一致、仍在预算内；
//   ② 每一个都**按名字直接调用**一次真机操作，并核对效果（不是只看"没挂住"）。
//
// Run: node test/advertised-surface.test.mjs
import { readFileSync } from 'node:fs';
import { startServer, createReport } from './lib/harness.mjs';

const PROMOTED = [
  'wps_excel_evaluate_formula', 'wps_excel_set_column_width', 'wps_excel_copy_range',
  'wps_excel_set_sheet_appearance', 'wps_excel_unmerge_cells', 'wps_excel_delete_rows',
  'wps_excel_freeze_panes', 'wps_excel_clear_range', 'wps_word_replace_range',
  // P2-2（2026-10-21，80 场景候选）：广告面 95 → 100。
  'wps_word_get_fields', 'wps_word_protect', 'wps_word_compare_documents',
  'wps_ppt_export_notes', 'wps_excel_get_links',
];
const BUDGET = { maxTools: 120, maxSchemaBytes: 80000 };

const srv = startServer({ clientName: 'advertised-surface' });
const { call, req, text, ok } = srv;
const { check, finish } = createReport('ADVERTISED SURFACE TESTS');
const viaCall = (tool, args) => call('wps_call', { tool, args });
const flat = (s) => String(s).replace(/\s+/g, ' ').slice(0, 140);

// ---------------- ① 广告面本身 ----------------
await call('wps_status', {});                       // 等 initialize 完成
const listed = await req('tools/list', {});
const tools = (listed.result && listed.result.tools) || [];
const names = tools.map((t) => t.name);
const spec = JSON.parse(readFileSync('spec/advertised.json', 'utf8'));
const missing = PROMOTED.filter((n) => !names.includes(n));
check('tools/list carries every promoted tool', missing.length === 0, missing.join(', ') || PROMOTED.length + ' names');
check('advertised surface matches spec/advertised.json', names.slice().sort().join(',') === spec.slice().sort().join(','), 'live=' + names.length + ' spec=' + spec.length);
const bytes = Buffer.byteLength(JSON.stringify(tools), 'utf8');
check('advertised tools within budget', tools.length <= BUDGET.maxTools, tools.length + ' / ' + BUDGET.maxTools);
check('advertised schema bytes within budget', bytes <= BUDGET.maxSchemaBytes, bytes + ' / ' + BUDGET.maxSchemaBytes);

// ---------------- ② 真机：直接按名字调用 ----------------
check('scratch workbook created', ok(await viaCall('wps_excel_create_workbook', {})), '');
const wbName = (text(await call('wps_excel_get_sheet_info', {})).match(/工作簿[:：]\s*([^\s|]+)/) || [])[1] || null;
check('scratch workbook name resolved', !!wbName, 'wb=' + wbName);
check('seed data written', ok(await call('wps_excel_write_range', {
  range: 'A1', data: [['项目', '数量', '金额'], ['甲', 2, 120.5], ['乙', 3, 80.25], ['丙', 1, 49.25]],
})), '');

check('set_column_width direct', ok(await call('wps_excel_set_column_width', { column: 'A', width: 18 })), '');
check('freeze_panes direct', ok(await call('wps_excel_freeze_panes', { row: 1 })), '');
const sum = await call('wps_excel_evaluate_formula', { formula: '=SUM(C2:C4)' });
check('evaluate_formula direct returns the real value', ok(sum) && /250/.test(text(sum)), flat(text(sum)));

check('copy_range direct', ok(await call('wps_excel_copy_range', { source: 'A1:C4', destination: 'E1' })), '');
const copied = text(await call('wps_excel_read_range', { range: 'E1:G4' }));
check('copy_range really landed the values', copied.includes('甲') && copied.includes('120.5'), flat(copied));

check('seed row for delete', ok(await call('wps_excel_write_range', { range: 'A10', data: [['待删'], ['保留']] })), '');
check('delete_rows direct', ok(await call('wps_excel_delete_rows', { startRow: 10, count: 1 })), '');
const afterDelete = text(await call('wps_excel_read_range', { range: 'A10:A11' }));
check('delete_rows really shifted the sheet', afterDelete.includes('保留') && !afterDelete.includes('待删'), flat(afterDelete));

check('merge_cells (setup)', ok(await viaCall('wps_excel_merge_cells', { range: 'A6:B6' })), '');
check('unmerge_cells direct', ok(await call('wps_excel_unmerge_cells', { range: 'A6:B6' })), '');
check('clear_range direct', ok(await call('wps_excel_write_range', { range: 'D1', data: [['X']] })) && ok(await call('wps_excel_clear_range', { range: 'D1' })), '');
const cleared = text(await call('wps_excel_read_range', { range: 'D1' }));
check('clear_range really emptied the cell', !/X/.test(cleared), flat(cleared));
check('set_sheet_appearance direct', ok(await call('wps_excel_set_sheet_appearance', { visible: 'visible', tabColor: '#FF9900' })), '');

check('word document created', ok(await call('wps_word_create_document', {})), '');
check('word text seeded', ok(await call('wps_word_insert_text', { text: 'ABCDEFGHIJ' })), '');
const replaced = await call('wps_word_replace_range', { startPos: 0, endPos: 5, text: '一二三' });
check('word_replace_range direct', ok(replaced), flat(text(replaced)));
const docText = text(await call('wps_word_get_document_text', {}));
check('word_replace_range really rewrote the range', docText.includes('一二三') && !docText.includes('ABCDE'), flat(docText));

// ---------------- 收尾 ----------------
const docName = (text(await call('wps_word_get_active_document', {})).match(/文档[:：]\s*([^\s|]+)/) || [])[1] || null;
await viaCall('wps_word_close_document', { save: false });
const openDocs = text(await call('wps_word_get_open_documents', {}));
check('word document closed', !docName || !openDocs.includes(docName), flat(openDocs));
await viaCall('wps_excel_close_workbook', { name: wbName, save: false });
const openWbs = text(await call('wps_excel_get_open_workbooks', {}));
check('scratch workbook closed', !!wbName && !openWbs.includes(wbName), flat(openWbs));

srv.stop();
finish();

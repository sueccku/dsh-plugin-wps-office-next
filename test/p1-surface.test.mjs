// P1 新工具面的静态验收：只走 MCP 握手 + tools/list，不需要 WPS。
// 目的：新工具必须真的注册进表面，且关键参数在 schema 里；同时让 coverage 门禁看到这些名字。
import { spawn } from 'node:child_process';

const NEW_TOOLS = [
  'wps_word_get_fields',
  'wps_word_update_fields',
  'wps_word_select_range',
  'wps_word_protect',
  'wps_word_get_protection',
  'wps_ppt_set_slide_hidden',
  'wps_ppt_get_hidden_slides',
  'wps_ppt_get_fonts',
  'wps_excel_allow_edit_range',
  'wps_excel_get_links',
  'wps_excel_insert_page_break',
  'wps_word_insert_caption',
  'wps_word_set_table_formula',
  'wps_excel_disconnect_link',
  'wps_excel_change_link',
  'wps_excel_get_charts',
  'wps_word_insert_table_of_figures',
  'wps_word_compare_documents',
  'wps_word_update_style',
  'wps_excel_set_chart_axis',
  'wps_ppt_export_notes',
  'wps_word_set_content_control',
  'wps_word_resolve_revisions',
  'wps_ppt_set_theme_colors',
  'wps_ppt_set_theme_fonts',
  'wps_ppt_add_section',
  'wps_excel_add_icon_set',
  'wps_excel_add_trendline',
  'wps_excel_import_csv',
  'wps_excel_add_pivot_calculated_field',
  'wps_excel_advanced_filter',
  'wps_excel_set_series_axis',
  'wps_word_create_labels',
  'wps_word_create_envelope',
  'wps_ppt_set_placeholder_lock',
  'wps_ppt_duplicate_master_layout',
  'wps_ppt_embed_fonts',
  'wps_ppt_compress_images',
  'wps_ppt_export_pdf',
  'wps_excel_group_pivot_dates',
];

const child = spawn(process.execPath, ['mcp/dist/index.js'], {
  stdio: ['pipe', 'pipe', 'pipe'],
  windowsHide: true,
  env: { ...process.env, WPS_OFFICE_TOOLSET: 'full' },
});
let buf = '';
const pending = new Map();
child.stdout.on('data', (d) => {
  buf += d.toString();
  let i;
  while ((i = buf.indexOf('\n')) >= 0) {
    const line = buf.slice(0, i).trim();
    buf = buf.slice(i + 1);
    if (!line) continue;
    let m;
    try { m = JSON.parse(line); } catch { continue; }
    if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); }
  }
});
child.stderr.on('data', () => {});
const req = (id, method, params) => new Promise((resolve) => {
  pending.set(id, resolve);
  child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
});

const results = [];
function check(name, ok, detail) {
  results.push({ name, ok });
  console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : ''));
}

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'p1-surface', version: '1' } });
child.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n');
const listed = await req(2, 'tools/list', {});
const tools = listed.result?.tools || [];
const byName = new Map(tools.map((t) => [t.name, t]));

check('tools/list is non-empty', tools.length > 0, tools.length + ' tools');
for (const name of NEW_TOOLS) {
  const tool = byName.get(name);
  check('registered: ' + name, !!tool, tool ? '' : 'missing from tools/list');
}
const schema = (name) => byName.get(name)?.inputSchema || {};
const props = (name) => Object.keys(schema(name).properties || {});
check('read_range advertises startRow/maxRows', props('wps_excel_read_range').includes('startRow') && props('wps_excel_read_range').includes('maxRows'));
check('get_document_text advertises start/end', props('wps_word_get_document_text').includes('start') && props('wps_word_get_document_text').includes('end'));
check('batch advertises maxResultChars', props('wps_batch').includes('maxResultChars'));
check('save_as advertises createDirs', props('wps_common_save_as').includes('createDirs'));
check('help advertises offset/limit', props('wps_help').includes('offset') && props('wps_help').includes('limit'));
check('word insert_table advertises data', props('wps_word_insert_table').includes('data'));
check('ppt insert_table advertises data', props('wps_ppt_insert_table').includes('data'));
check('ppt get_slide_notes advertises all', props('wps_ppt_get_slide_notes').includes('all'));
check('sort_range advertises keys/hasHeader', props('wps_excel_sort_range').includes('keys') && props('wps_excel_sort_range').includes('hasHeader'));
check('slide size advertises scaleContent', props('wps_ppt_set_slide_size').includes('scaleContent'));
check('header advertises firstPageDifferent', props('wps_word_insert_header').includes('firstPageDifferent'));
check('ppt export_notes advertises outputPath', props('wps_ppt_export_notes').includes('outputPath'));
check('word compare advertises original/revised', props('wps_word_compare_documents').includes('original') && props('wps_word_compare_documents').includes('revised'));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'P1 SURFACE TESTS OK (' + results.length + ')' : 'P1 SURFACE TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

// Verifies the P3 Word deep tools against a real WPS Writer instance: bookmarks / comments /
// document stats / hyperlinks, then the table family (read, cell write, add/delete lines,
// merge, split, format, convert to text).
// Run: node test/word-deep.test.mjs
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
const text = async (name, args) => textOf(await call(name, args));

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'wd', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

check('scratch document created', ok(await call('wps_word_create_document', {})), '');
check('insert_text works', ok(await call('wps_word_insert_text', { text: 'P3 深水区测试文档。', position: 'end' })), '');
const stats = await text('wps_word_get_document_stats', {});
check('get_document_stats reports page/word/paragraph counts', /页数 \d+/.test(stats) && /字数 \d+/.test(stats) && /段落 \d+/.test(stats), stats.replace(/\n/g, ' | ').slice(0, 130));
check('insert_bookmark works', ok(await call('wps_word_insert_bookmark', { name: 'P3Mark' })), '');
const bookmarks = await text('wps_word_get_bookmarks', {});
check('get_bookmarks lists it', bookmarks.includes('P3Mark'), bookmarks.replace(/\n/g, ' | ').slice(0, 110));
check('insert_comment works', ok(await call('wps_word_insert_comment', { text: '这是一条测试批注' })), '');
const comments = await text('wps_word_get_comments', {});
check('get_comments lists it', comments.includes('测试批注') && comments.includes('批注（'), comments.replace(/\n/g, ' | ').slice(0, 140));
const hyper = await text('wps_word_insert_hyperlink', { url: 'https://example.com', text: '示例链接' });
check('insert_hyperlink inserts a link', hyper.includes('示例链接') && hyper.includes('example.com'), hyper.slice(0, 110));
check('insert_hyperlink without url is rejected', !ok(await call('wps_word_insert_hyperlink', {})), '');

check('insert_table works (existing tool)', ok(await call('wps_word_insert_table', { rows: 3, cols: 3 })), '');
const tables = await text('wps_word_get_tables', {});
check('get_tables finds it with its shape', tables.includes('3 行 x 3 列'), tables.replace(/\n/g, ' | ').slice(0, 130));
const cell = await text('wps_word_set_table_cell', { table: 1, row: 1, column: 1, text: '地区' });
check('set_table_cell writes and reads back', cell.includes('地区'), cell.slice(0, 110));
const data = await text('wps_word_get_table_data', { table: 1 });
check('get_table_data shows the written cell', data.includes('地区') && data.includes('第1行'), data.replace(/\n/g, ' | ').slice(0, 140));
const addedRow = await text('wps_word_add_table_lines', { table: 1, kind: 'row', count: 1 });
check('add_table_lines adds a row', addedRow.includes('4 行'), addedRow.slice(0, 120));
const addedCol = await text('wps_word_add_table_lines', { table: 1, kind: 'column' });
check('add_table_lines adds a column', addedCol.includes('4 列'), addedCol.slice(0, 120));
const deleted = await text('wps_word_delete_table_line', { table: 1, kind: 'row', lineIndex: 4 });
check('delete_table_line removes a row', deleted.includes('3 行'), deleted.slice(0, 120));
check('an out-of-range line is rejected', !ok(await call('wps_word_delete_table_line', { table: 1, kind: 'row', lineIndex: 99 })), '');
const merged = await text('wps_word_merge_table_cells', { table: 1, startRow: 1, startColumn: 1, endRow: 1, endColumn: 2 });
check('merge_table_cells merges', merged.includes('已合并单元格'), merged.slice(0, 130));
const split = await text('wps_word_split_table_cell', { table: 1, row: 1, column: 1, rows: 1, columns: 2 });
check('split_table_cell splits', split.includes('已拆分单元格'), split.slice(0, 130));
const formatted = await text('wps_word_set_table_format', { table: 1, borders: true, headerShading: '#D9E2F3', autoFit: 'content' });
check('set_table_format reports what it applied', formatted.includes('borders') && formatted.includes('headerShading'), formatted.slice(0, 130));
check('an unrecognized shading colour is rejected', !ok(await call('wps_word_set_table_format', { table: 1, headerShading: 'nonsense' })), '');
check('get_table_data with a bad index is rejected', !ok(await call('wps_word_get_table_data', { table: 99 })), '');
check('set_table_cell without text is rejected', !ok(await call('wps_word_set_table_cell', { table: 1, row: 2, column: 2 })), '');
const converted = await text('wps_word_convert_table_to_text', { table: 1, separator: 'tab' });
check('convert_table_to_text removes the table object', converted.includes('还剩 0 张表'), converted.slice(0, 120));
const tablesAfter = await text('wps_word_get_tables', {});
check('no table remains after conversion', tablesAfter.includes('没有表格'), tablesAfter.replace(/\n/g, ' | ').slice(0, 90));

check('document closed', ok(await call('wps_word_close_document', { save: false })), '');
const open = await text('wps_word_get_open_documents', {});
check('no document left open', /共?0个|没有打开/.test(open), open.replace(/\n/g, ' | ').slice(0, 70));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'WORD DEEP TESTS OK (' + results.length + ')' : 'WORD DEEP TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

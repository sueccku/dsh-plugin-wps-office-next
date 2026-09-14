// Verifies the P3-3 Word document-production tools against a real WPS Writer instance: page numbers,
// columns, revision listing/accept/reject, and comment deletion.
// Run: node test/word-produce.test.mjs
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

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'wp', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

check('scratch document created', ok(await call('wps_word_create_document', {})), '');
check('insert_text works', ok(await call('wps_word_insert_text', { text: 'P3-3 生产族测试。', position: 'end' })), '');

const footer = await text('wps_word_insert_page_numbers', { alignment: 'center' });
check('page numbers land in the footer and report the count', footer.includes('页脚') && footer.includes('居中') && footer.includes('1 个页码域'), footer.slice(0, 130));
const header = await text('wps_word_insert_page_numbers', { position: 'header', alignment: 'right' });
check('page numbers can go in the header instead', header.includes('页眉') && header.includes('右对齐'), header.slice(0, 130));
check('a bad section index is rejected', !ok(await call('wps_word_insert_page_numbers', { section: 99 })), '');

const twoCols = await text('wps_word_set_columns', { count: 2, spacing: 20 });
check('set_columns makes two columns', twoCols.includes('现在是 2 栏') && twoCols.includes('spacing'), twoCols.slice(0, 140));
const oneCol = await text('wps_word_set_columns', { count: 1 });
check('set_columns can go back to one column', oneCol.includes('现在是 1 栏'), oneCol.slice(0, 120));
check('an out-of-range column count is rejected', !ok(await call('wps_word_set_columns', { count: 99 })), '');

check('track changes enabled (existing tool)', ok(await call('wps_word_enable_track_changes', { enable: true })), '');
await call('wps_word_insert_text', { text: '这是带修订的一行。', position: 'end' });
const revisions = await text('wps_word_get_revisions', {});
check('get_revisions sees the tracked change', revisions.includes('修订跟踪: 已打开') && revisions.includes('修订（') && revisions.includes('插入'), revisions.replace(/\n/g, ' | ').slice(0, 160));
const rejected = await text('wps_word_reject_revisions', { index: 1 });
check('reject_revisions drops that change', rejected.includes('已拒绝 1 处') && rejected.includes('还剩 0 处'), rejected.slice(0, 120));
await call('wps_word_insert_text', { text: '再改一次。', position: 'end' });
const accepted = await text('wps_word_accept_revisions', {});
check('accept_revisions accepts everything', accepted.includes('已接受 1 处') && accepted.includes('还剩 0 处'), accepted.slice(0, 120));
const revisionsAfter = await text('wps_word_get_revisions', {});
check('no revisions remain once handled', revisionsAfter.includes('没有修订'), revisionsAfter.replace(/\n/g, ' | ').slice(0, 110));
await call('wps_word_enable_track_changes', { enable: false });

await call('wps_word_insert_comment', { text: '第一条意见' });
await call('wps_word_insert_comment', { text: '第二条意见' });
const comments = await text('wps_word_get_comments', {});
check('two comments are in the document', comments.includes('批注（2 条）'), comments.replace(/\n/g, ' | ').slice(0, 130));
const deletedOne = await text('wps_word_delete_comment', { index: 1 });
check('delete_comment removes one', deletedOne.includes('已删除 1 条') && deletedOne.includes('还剩 1 条'), deletedOne.slice(0, 120));
const deletedAll = await text('wps_word_delete_comment', {});
check('delete_comment with no index removes the rest', deletedAll.includes('已删除 1 条') && deletedAll.includes('还剩 0 条'), deletedAll.slice(0, 120));
const commentsAfter = await text('wps_word_get_comments', {});
check('no comments remain', commentsAfter.includes('没有批注'), commentsAfter.replace(/\n/g, ' | ').slice(0, 90));

check('document closed', ok(await call('wps_word_close_document', { save: false })), '');
const open = await text('wps_word_get_open_documents', {});
check('no document left open', /共?0个|没有打开/.test(open), open.replace(/\n/g, ' | ').slice(0, 70));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'WORD PRODUCE TESTS OK (' + results.length + ')' : 'WORD PRODUCE TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

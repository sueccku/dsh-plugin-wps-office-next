// Verifies the P3-4 Word long-tail tools against a real WPS Writer instance: content controls,
// footnotes/endnotes, index, cross reference, and mail merge from a CSV.
// Run: node test/word-longtail.test.mjs
import { spawn } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = dirname(dirname(fileURLToPath(import.meta.url)));
const artifacts = join(ROOT, 'test', '.artifacts', 'word-longtail');
mkdirSync(artifacts, { recursive: true });
const csvPath = join(artifacts, 'merge-data.csv');
writeFileSync(csvPath, 'Name,City\r\nAlice,Beijing\r\nBob,Shanghai\r\n', 'utf8');

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

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'wlt', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

check('scratch document created', ok(await call('wps_word_create_document', {})), '');
check('insert_text works', ok(await call('wps_word_insert_text', { text: 'P3-4 长尾测试。', position: 'end' })), '');

const cc = await text('wps_word_add_content_control', { type: 'plainText', title: '姓名', tag: 'name', text: '张三' });
check('add_content_control inserts a typed control', cc.includes('纯文本') && cc.includes('姓名') && cc.includes('张三'), cc.slice(0, 130));
const controls = await text('wps_word_get_content_controls', {});
check('get_content_controls lists it', controls.includes('内容控件（1 个）') && controls.includes('纯文本「姓名」：张三'), controls.replace(/\n/g, ' | ').slice(0, 150));
check('an unknown control type is rejected', !ok(await call('wps_word_add_content_control', { type: 'nonsense' })), '');

const footnote = await text('wps_word_add_footnote', { text: '脚注内容ABC' });
check('add_footnote inserts a footnote', footnote.includes('脚注内容ABC') && footnote.includes('共 1 条'), footnote.slice(0, 110));
const endnote = await text('wps_word_add_endnote', { text: '尾注内容XYZ' });
check('add_endnote inserts an endnote', endnote.includes('尾注内容XYZ') && endnote.includes('共 1 条'), endnote.slice(0, 110));
const notes = await text('wps_word_get_notes', {});
check('get_notes reads both note bodies back', notes.includes('脚注 1 条，尾注 1 条') && notes.includes('脚注内容ABC') && notes.includes('尾注内容XYZ'), notes.replace(/\n/g, ' | ').slice(0, 160));
check('add_footnote without text is rejected', !ok(await call('wps_word_add_footnote', {})), '');

const index = await text('wps_word_insert_index', {});
check('insert_index adds an index object', index.includes('已插入索引'), index.slice(0, 120));

check('insert_bookmark works (existing tool)', ok(await call('wps_word_insert_bookmark', { name: 'RefTarget' })), '');
const crossRef = await text('wps_word_insert_cross_reference', { referenceType: 2, referenceKind: -1, referenceItem: 'RefTarget' });
check('insert_cross_reference accepts a bookmark reference', crossRef.includes('已插入交叉引用') && crossRef.includes('RefTarget'), crossRef.slice(0, 140));
check('insert_cross_reference without an item is rejected', !ok(await call('wps_word_insert_cross_reference', {})), '');

check('mail_merge without a data file is rejected', !ok(await call('wps_word_mail_merge', {})), '');
check('mail_merge with a missing file is rejected', !ok(await call('wps_word_mail_merge', { dataFile: join(artifacts, 'nope.csv') })), '');
const merged = await text('wps_word_mail_merge', { dataFile: csvPath, fields: ['Name', 'City'] });
check('mail_merge produces a new document', merged.includes('新文档') && merged.includes('字段: Name, City'), merged.replace(/\n/g, ' | ').slice(0, 170));
check('the merged document carries the data rows', merged.includes('Alice') && merged.includes('Bob'), merged.replace(/\n/g, ' | ').slice(0, 170));

// After Execute() the merged copy is the active document; close it, then the master.
check('merged document closed', ok(await call('wps_word_close_document', { save: false })), '');
check('master document closed', ok(await call('wps_word_close_document', { save: false })), '');
const open = await text('wps_word_get_open_documents', {});
check('no document left open', /共?0个|没有打开/.test(open), open.replace(/\n/g, ' | ').slice(0, 70));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'WORD LONGTAIL TESTS OK (' + results.length + ')' : 'WORD LONGTAIL TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

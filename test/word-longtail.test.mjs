// Verifies the P3-4 Word long-tail tools against a real WPS Writer instance: content controls,
// footnotes/endnotes, index, cross reference, and mail merge from a CSV.
// Run: node test/word-longtail.test.mjs
import { spawn } from 'node:child_process';
import { resolve } from 'node:path';
// FIXES 95（W1-4）：脚注、尾注、内容控件全都只走插件自己的回读。
// 它们在 docx 里是**独立的 part**（word/footnotes.xml、endnotes.xml、document.xml 里的 <w:sdt>），
// 直接读文件最直接：回报"插入了脚注"但文件里没有，用户存盘后打开就是空的。
import { zipScan } from './lib/oracle.mjs';
import { mkdirSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
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

// 这些文档只为本文件的断言存在，且由 wps_common_save_as 落盘：每轮先清掉，否则上一轮的产物
// 会让"份数"这类断言假红（本轮就踩过：张三.docx 与张三_1.docx 同时存在）。
for (const p of ['word-longtail.docx']) {
  try { rmSync(join(artifacts, p), { force: true }); } catch {}
}
for (const d of ['merge-split-out', 'merge-split-out2']) {
  try { rmSync(join(artifacts, d), { recursive: true, force: true }); } catch {}
}

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

// ---- 邮件合并：既判数据真的落盘，又判母版没被动过（S5 场景走查查出的 P0 缺陷的回归）----
// 旧实现把字段插到当前选区再 MailMerge.Execute()，实测会把母版正文顶掉、占位符从不被替换，
// 却回 success；而旧断言只查了**工具自述文本**里的 "新文档/字段"，所以一直没红。
const csvBare = join(artifacts, 'merge-bare.csv');
const csvBom = join(artifacts, 'merge-bom.csv');
const csvData = join(artifacts, 'merge-data2.csv');
writeFileSync(csvBare, 'Name,City\r\nAlice,Beijing\r\nBob,Shanghai\r\n', 'utf8');
writeFileSync(csvBom, '\ufeff姓名,城市\r\n王五,杭州\r\n', 'utf8');
writeFileSync(csvData, '\ufeff姓名,城市\r\n张三,北京\r\n李四,上海\r\n王五,杭州\r\n', 'utf8');
const mDocx = join(artifacts, 'merge-master.docx');
const sDocx = join(artifacts, 'merge-split.docx');
const bDocx = join(artifacts, 'merge-bom.docx');
const nDocx = join(artifacts, 'merge-noplaceholder.docx');
const mergeOut = join(artifacts, 'merge-single-out.docx');
const splitDir = join(artifacts, 'merge-split-out');

async function writeTemplate(path, rows) {
  await call('wps_word_create_document', {});
  for (const line of rows) await call('wps_word_insert_text', { text: line, position: 'end' });
  return call('wps_common_save_as', { outputPath: path });
}
// zipScan 在"没找到"时返回的也是**真值对象** {count:0}，所以断言必须看 count，
// 不能写 !zipScan(...)（那会永远为假 —— 这个坑本文件自己踩过一次）。
function hits(file, needle) { const r = zipScan(file, needle); return !!(r && r.count > 0); }

const tplRows = ['入 职 通 知 书', '姓名：⟨姓名⟩', '城市：⟨城市⟩', '宁川精工机械'];
check('merge template written', ok(await writeTemplate(mDocx, tplRows)), '');
check('merge template (BOM case) written', ok(await writeTemplate(bDocx, tplRows)), '');
check('merge template (split case) written', ok(await writeTemplate(sDocx, tplRows)), '');
check('merge template (no placeholder) written', ok(await writeTemplate(nDocx, ['这是一份没有占位符的文档。'])), '');

// (0) 没带占位符的文档：必须被拒，而不是产出一份"看起来成功、其实没数据"的文档
check('mail_merge without a data file is rejected', !ok(await call('wps_word_mail_merge', {})), '');
check('mail_merge with a missing file is rejected', !ok(await call('wps_word_mail_merge', { dataFile: join(artifacts, 'nope.csv') })), '');
check('mail_merge on a document without placeholders is rejected',
  !ok(await call('wps_word_mail_merge', { dataFile: csvBare, fields: ['Name', 'City'] })), '');
check('bare document closed', ok(await call('wps_word_close_document', { save: false })), '');

// (1) 单个新文档：数据必须在，占位符必须没了，母版必须原样
check('mail_merge master reopened for the single-document case', ok(await call('wps_word_open_document', { path: mDocx })), '');
const single = await text('wps_word_mail_merge', { dataFile: csvData, fields: ['姓名', '城市'], outputPath: mergeOut });
check('mail_merge single document reports the data rows', single.includes('张三') && single.includes('王五'), single.replace(/\n/g, ' | ').slice(0, 170));
check('single merged docx really contains the first data row', hits(mergeOut, '张三') && hits(mergeOut, '李四'), JSON.stringify(zipScan(mergeOut, '张三')));
check('single merged docx replaced every placeholder', !hits(mergeOut, '⟨姓名⟩'), JSON.stringify(zipScan(mergeOut, '⟨姓名⟩')));
check('single merged docx has no field code left', !hits(mergeOut, 'MERGEFIELD'), '');
check('single merged docx kept the master text', hits(mergeOut, '入 职 通 知 书'), '');
check('single merged document closed', ok(await call('wps_word_close_document', { save: false })), '');
check('single-document master closed', ok(await call('wps_word_close_document', { save: false })), '');
// 母版未被动过的判据看**文件内容**而不是文件 sha256：wps_common_save_as 会重写 docx 包
// （docProps 等），拿二进制哈希比一定假红 —— 这是本文件自己的一个假指控，已修。
check('the master document still carries its placeholders', hits(mDocx, '⟨姓名⟩') && hits(mDocx, '城市'), JSON.stringify(zipScan(mDocx, '⟨姓名⟩')));

// (2) splitDir：每个数据行一份，文件名取自指定字段
check('mail_merge master reopened for the split case', ok(await call('wps_word_open_document', { path: sDocx })), '');
const split = await text('wps_word_mail_merge', { dataFile: csvData, fields: ['姓名', '城市'], splitDir, fileNameField: '姓名' });
check('mail_merge split mode reports 3 written files', split.includes('3 份'), split.replace(/\n/g, ' | ').slice(0, 170));
const splitFiles = readdirSync(splitDir).filter((f) => f.endsWith('.docx'));
check('mail_merge split mode wrote one docx per data row', splitFiles.length === 3, splitFiles.join(', '));
check('mail_merge split mode honours fileNameField as a column name',
  splitFiles.includes('张三.docx') && splitFiles.includes('李四.docx'), splitFiles.join(', '));
const s1 = join(splitDir, splitFiles.find((f) => f.startsWith('张三')) || 'x');
check('a split file carries its own row only',
  hits(s1, '张三') && hits(s1, '北京') && !hits(s1, '李四'), JSON.stringify(zipScan(s1, '李四')));
check('split master closed', ok(await call('wps_word_close_document', { save: false })), '');
check('the master document was not modified by the split merge', hits(sDocx, '⟨姓名⟩'), JSON.stringify(zipScan(sDocx, '⟨姓名⟩')));

// fileNameField 是一段带 {列名} 的模板时也要生效（列可以不在 fields 里）
const splitDir2 = join(artifacts, 'merge-split-out2');
const split2 = await text('wps_word_mail_merge', {
  dataFile: csvData, fields: ['姓名', '城市'], splitDir: splitDir2,
  fileNameField: '通知书_{城市}_{姓名}',
});
check('mail_merge split mode accepts a {column} filename pattern', split2.includes('3 份'), split2.replace(/\n/g, ' | ').slice(0, 170));
const splitFiles2 = readdirSync(splitDir2).filter((f) => f.endsWith('.docx'));
check('the {column} pattern produced the expected filenames',
  splitFiles2.includes('通知书_北京_张三.docx') && splitFiles2.includes('通知书_杭州_王五.docx'), splitFiles2.join(', '));

// (3) UTF-8 BOM 的中文表头（Excel 导出的 CSV 就是这样）
check('mail_merge master reopened for the BOM case', ok(await call('wps_word_open_document', { path: bDocx })), '');
const bom = await text('wps_word_mail_merge', { dataFile: csvBom, fields: ['姓名', '城市'] });
check('mail_merge reads a UTF-8 BOM CSV with Chinese columns', bom.includes('王五') && bom.includes('杭州'), bom.replace(/\n/g, ' | ').slice(0, 170));
check('BOM-case merged document closed', ok(await call('wps_word_close_document', { save: false })), '');
check('BOM-case master closed', ok(await call('wps_word_close_document', { save: false })), '');

// 回到脚注/尾注那份文档：它随宿主会话存活，没落过盘，重开磁盘文件是拿不到的 ——
// 这里改成重新建一份并明确落盘，末尾的独立 oracle 以**磁盘上的文件**为准。
const WL = resolve('test/.artifacts/word-longtail.docx');
for (const tool of ['wps_word_add_footnote', 'wps_word_add_endnote']) {
  await call(tool, { text: tool.includes('footnote') ? '脚注内容ABC' : '尾注内容XYZ' });
}
check('long-tail document rebuilt for the disk oracle', ok(await call('wps_common_save_as', { outputPath: WL })), '');

// ---- 独立 oracle：以磁盘上的 docx 为准 ----
const fnScan = zipScan(WL, '脚注内容ABC', 'word/');
check('the footnote text really landed in the docx', !!fnScan && fnScan.count >= 1, JSON.stringify(fnScan));
const enScan = zipScan(WL, '尾注内容XYZ', 'word/');
check('the endnote text really landed in the docx', !!enScan && enScan.count >= 1, JSON.stringify(enScan));
const sdtScan = zipScan(WL, '<w:sdt>', 'word/');
check('the content control really landed in the docx', !!sdtScan && sdtScan.count >= 1, JSON.stringify(sdtScan));

await call('wps_word_close_document', { save: false });
// 逐个收尾，不假设哪个文档当前活动（合并/拆分都会新建临时文档）
for (let i = 0; i < 6; i++) {
  const list = await text('wps_word_get_open_documents', {});
  if (/共?0个|没有打开/.test(list)) break;
  await call('wps_word_close_document', { save: false });
}
const open = await text('wps_word_get_open_documents', {});
check('no document left open', /共?0个|没有打开/.test(open), open.replace(/\n/g, ' | ').slice(0, 70));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'WORD LONGTAIL TESTS OK (' + results.length + ')' : 'WORD LONGTAIL TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

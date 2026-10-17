// A4（改进指南 §9 P1）：**目标身份必须能读回来**。
//
// 背景：本轮没有、也不假装有 documentId。能在不引入句柄的前提下给出的最小可用能力是
// "读回目标身份，写前写后各读一次来对账"。这条测试就是把那份能力钉死：
//   ① Excel：get_sheet_info 必须回报 workbookName + workbookPath（路径是能对账的那一半）；
//      未落盘的簿必须如实回报空路径，而不是编一个名字。
//   ② PPT：get_slide_info 必须回报 presentationName / presentationPath 与**画布尺寸**
//      （slideWidth/slideHeight 是磅，slideWidthPx/slideHeightPx 是像素）——
//      形状坐标是磅、set_slide_size 与导出尺寸是像素，不知道画布就排不对版。
//   ③ Word：get_active_document 回报 name + path。
// 全部用裸 COM 交叉验证（不读插件自述），临时文件用完即删。
// Run: node test/target-identity.test.mjs   (needs WPS)
import { spawn } from 'node:child_process';
import { mkdirSync, rmSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { com } from './lib/oracle.mjs';

const OUT = join(tmpdir(), 'wps-target-identity');
rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
const XLSX = join(OUT, '身份探针.xlsx');
const PPTX = join(OUT, '身份探针.pptx');
const DOCX = join(OUT, '身份探针.docx');

const child = spawn(process.execPath, ['mcp/dist/index.js'], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
let buf = '';
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + '\n'); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); }); }
child.stdout.on('data', (d) => { buf += d.toString(); let i; while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on('data', () => {});

const results = [];
function check(name, ok, detail) { results.push({ name, ok: !!ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }
function payload(res) { try { return JSON.parse(text(res)); } catch { return {}; } }

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'target-identity', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });
let id = 10;
const call = (name, args) => req(id++, 'tools/call', { name, arguments: args });
const viaAction = (method, params) => call('wps_call', { tool: 'wps_execute_method', args: { method, params: params || {} } });
const raw = async (method, params) => payload(await viaAction(method, params));
async function closeAll(method, appType) {
  for (let i = 0; i < 10; i++) {
    const res = await call('wps_call', { tool: 'wps_execute_method', args: { method, params: { save: false }, appType } });
    if (!ok(res)) break;
  }
}

// ---------------- Excel ----------------
// 注意：第一方工具返回的是**给人看的文本**，不是桥的原始 JSON。身份字段必须出现在那段文本里，
// 否则"桥有、模型看不到"等于没有这个能力 —— 所以这里断言的是工具输出文本。
await closeAll('closeWorkbook', 'et');
check('scratch workbook created', ok(await call('wps_excel_create_workbook', {})), '');
const unsavedText = text(await call('wps_excel_get_sheet_info', {}));
check('A4: an unsaved workbook honestly says it has no disk path', /工作簿路径: \(尚未保存到磁盘\)/.test(unsavedText), unsavedText.replace(/\s+/g, ' ').slice(0, 140));
check('A4: an unsaved workbook still reports its name', /工作簿: \S+/.test(unsavedText), unsavedText.split('\n')[0]);

check('workbook saved to a scratch path', ok(await call('wps_common_save_as', { outputPath: XLSX })), '');
const infoText = text(await call('wps_excel_get_sheet_info', {}));
check('A4: get_sheet_info reports workbookPath', /工作簿路径: .+\\/.test(infoText), infoText.replace(/\s+/g, ' ').slice(0, 160));
check('A4: the reported path is the file we saved', infoText.includes(XLSX), infoText.replace(/\s+/g, ' ').slice(0, 200));
check('A4: the file really exists on disk', existsSync(XLSX), XLSX);
// 独立路径：裸 COM 读同一个属性，证明这不是插件自述
const rawPath = String(com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application'); [string]$w.ActiveWorkbook.FullName")).trim();
check('A4: raw COM agrees on the workbook path', rawPath.toLowerCase() === XLSX.toLowerCase(), JSON.stringify(rawPath));
const rawName = String(com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application'); [string]$w.ActiveWorkbook.Name")).trim();
check('A4: workbookName agrees with raw COM', infoText.includes('工作簿: ' + rawName), JSON.stringify(rawName));
check('A4: currentSheet is reported for identity', /当前工作表: \S+/.test(infoText), infoText.replace(/\s+/g, ' ').slice(0, 160));
await closeAll('closeWorkbook', 'et');

// ---------------- PowerPoint ----------------
await closeAll('closePresentation', 'wpp');
check('scratch presentation created', ok(await call('wps_ppt_create_presentation', {})), '');
check('one slide added', ok(await call('wps_ppt_add_slide', { layout: 'blank' })), '');
check('presentation saved to a scratch path', ok(await call('wps_common_save_as', { outputPath: PPTX })), '');
const pinfoText = text(await call('wps_ppt_get_slide_info', { slideIndex: 1 }));
check('A4: get_slide_info reports the presentation name', /演示文稿: \S+/.test(pinfoText), pinfoText.replace(/\s+/g, ' ').slice(0, 140));
check('A4: get_slide_info reports the presentation path', pinfoText.includes(PPTX), pinfoText.replace(/\s+/g, ' ').slice(0, 200));
check('A4: the file really exists on disk', existsSync(PPTX), PPTX);
const rawW = Number(String(com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Kwpp.Application'); [string]$w.ActivePresentation.PageSetup.SlideWidth")).trim());
const rawH = Number(String(com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Kwpp.Application'); [string]$w.ActivePresentation.PageSetup.SlideHeight")).trim());
const canvas = /画布尺寸: ([\d.]+) × ([\d.]+) 磅（([\d.]+) × ([\d.]+) 像素）/.exec(pinfoText);
check('A4: the canvas is printed with both units', !!canvas, pinfoText.replace(/\s+/g, ' ').slice(0, 140));
check('A4: slideWidth matches raw COM (points)', !!canvas && Math.abs(Number(canvas[1]) - rawW) < 0.5, 'tool=' + (canvas && canvas[1]) + ' raw=' + rawW);
check('A4: slideHeight matches raw COM (points)', !!canvas && Math.abs(Number(canvas[2]) - rawH) < 0.5, 'tool=' + (canvas && canvas[2]) + ' raw=' + rawH);
// 画布像素换算：磅 = 像素 × 0.75 ⇒ 像素 = 磅 ÷ 0.75。Skill 里的单位表就靠这条成立。
check('A4: the pixel figure is the documented 磅/0.75 conversion', !!canvas && Math.abs(Number(canvas[3]) - rawW / 0.75) < 0.05 && Math.abs(Number(canvas[4]) - rawH / 0.75) < 0.05, 'px=' + (canvas && [canvas[3], canvas[4]].join(' × ')) + ' expected=' + (rawW / 0.75) + ' × ' + (rawH / 0.75));
await closeAll('closePresentation', 'wpp');

// ---------------- Word ----------------
for (let i = 0; i < 10; i++) {
  const res = await call('wps_call', { tool: 'wps_execute_method', args: { method: 'closeDocument', params: { save: false }, appType: 'wps' } });
  if (!ok(res)) break;
}
check('word document created', ok(await call('wps_word_create_document', {})), '');
check('word text seeded', ok(await call('wps_word_insert_text', { text: '身份探针' })), '');
check('document saved to a scratch path', ok(await call('wps_common_save_as', { outputPath: DOCX })), '');
const winfoText = text(await call('wps_word_get_active_document', {}));
check('A4: get_active_document reports the name', /当前文档: \S+/.test(winfoText), winfoText.replace(/\s+/g, ' ').slice(0, 140));
check('A4: get_active_document reports the path', winfoText.includes(DOCX), winfoText.replace(/\s+/g, ' ').slice(0, 200));
check('A4: the document file exists on disk', existsSync(DOCX), DOCX);
await viaAction('closeDocument', { save: false });

child.kill();
rmSync(OUT, { recursive: true, force: true });
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'TARGET IDENTITY TESTS OK (' + results.length + ')' : 'TARGET IDENTITY TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

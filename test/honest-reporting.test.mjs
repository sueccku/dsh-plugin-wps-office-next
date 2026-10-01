// FIXES 71: a result must describe what happened, not what was asked for.
//   - close(save=true) reports saved only after a real Save() ran
//   - protect / unprotect read the protection state back
//   - deleteColumns reports the whole span it deletes
//   - refreshLinks scalarises the link list
//   - convertToPDF on PPT copies the file instead of repointing the source
// Needs a real WPS installation (it drives the apps); the static half also guards the source shape.
// Run: node test/honest-reporting.test.mjs
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';

const DIR = 'test/.artifacts/honest-reporting';
rmSync(DIR, { recursive: true, force: true });
mkdirSync(DIR, { recursive: true });

const child = spawn(process.execPath, ['mcp/dist/index.js'], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
let buf = '';
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + '\n'); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); }); }
child.stdout.on('data', (d) => { buf += d.toString(); let i; while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on('data', () => {});
await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'honest-reporting', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });
let id = 10;
const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
async function call(name, args) {
  const r = await req(id++, 'tools/call', { name, arguments: args || {} });
  const isErr = !!(r.result && r.result.isError);
  const text = String((r.result && r.result.content && r.result.content[0].text) || '').replace(/\s+/g, ' ');
  return { isErr, text };
}
const abs = (n) => join(process.cwd(), DIR, n).replace(/\//g, '\\');

// ---------- close(save=true) only reports saved after a real Save ------------------------------------
const savedBook = abs('saved.xlsx');
await call('wps_excel_create_workbook', {});
const saveAs = await call('wps_common_save_as', { filePath: savedBook });
check('a workbook is on disk before the close test', !saveAs.isErr && existsSync(savedBook), saveAs.text.slice(0, 80));
await call('wps_excel_write_range', { range: 'A1', data: [[1]] });
const closedSaved = await call('wps_excel_close_workbook', { save: true });
check('closing a saved workbook reports 已保存', !closedSaved.isErr && /已保存/.test(closedSaved.text), closedSaved.text.slice(0, 90));

await call('wps_excel_create_workbook', {});
await call('wps_excel_write_range', { range: 'A1', data: [[2]] });
const closedNever = await call('wps_excel_close_workbook', { save: true });
check('a never-saved workbook is not reported as saved', !closedNever.isErr && /未保存/.test(closedNever.text), closedNever.text.slice(0, 90));
check('and the drop is explained as a warning', /注意/.test(closedNever.text) && /never saved/.test(closedNever.text), closedNever.text.slice(0, 140));

// ---------- protect / unprotect read back ------------------------------------------------------------
await call('wps_excel_create_workbook', {});
const protectedRes = await call('wps_excel_protect_sheet', { protect: true, password: 'pw' });
check('protecting a sheet succeeds', !protectedRes.isErr && /成功/.test(protectedRes.text), protectedRes.text.slice(0, 70));
const wrongPw = await call('wps_excel_protect_sheet', { protect: false, password: 'definitely-wrong' });
check('unprotecting with the wrong password fails instead of reporting success', wrongPw.isErr, wrongPw.text.slice(0, 110));
const rightPw = await call('wps_excel_protect_sheet', { protect: false, password: 'pw' });
check('unprotecting with the right password succeeds', !rightPw.isErr && /成功/.test(rightPw.text), rightPw.text.slice(0, 70));
const noLinks = await call('wps_excel_refresh_links', {});
check('refreshLinks reports zero links, not a character count', !noLinks.isErr && /0 条/.test(noLinks.text), noLinks.text.slice(0, 80));
await call('wps_excel_close_workbook', { save: false });

// ---------- deleteColumns covers the whole span -------------------------------------------------------
await call('wps_excel_create_workbook', {});
await call('wps_excel_write_range', { range: 'A1:F1', data: [['a', 'b', 'c', 'd', 'e', 'f']] });
const deleted = await call('wps_excel_delete_columns', { column: 'B', count: 3 });
check('deleteColumns reports the span it removed', !deleted.isErr && /3列/.test(deleted.text) && /3 个非空/.test(deleted.text), deleted.text.slice(0, 110));
const afterDelete = await call('wps_excel_read_range', { range: 'A1:C1' });
check('the three columns are really gone', /a \| e \| f/.test(afterDelete.text), afterDelete.text.slice(0, 90));
await call('wps_excel_close_workbook', { save: false });

// ---------- PPT export must not repoint the source ---------------------------------------------------
const pptx = abs('source.pptx');
const pdf = abs('source.pdf');
await call('wps_ppt_create_presentation', {});
await call('wps_ppt_add_slide', {});
const pptSaveAs = await call('wps_common_save_as', { filePath: pptx });
check('the presentation is on disk before the export test', !pptSaveAs.isErr && existsSync(pptx), pptSaveAs.text.slice(0, 80));
const beforeList = await call('wps_ppt_get_open_presentations', {});
const pathBefore = (/路径: ([^ ]+.pptx)/.exec(beforeList.text) || [])[1];
await call('wps_ppt_add_slide', {});
const exported = await call('wps_convert_to_pdf', { appType: 'ppt', outputPath: pdf });
check('converting a presentation to PDF succeeds', !exported.isErr && /导出成功|成功/.test(exported.text), exported.text.slice(0, 100));
check('the PDF really exists', existsSync(pdf), pdf);
const afterList = await call('wps_ppt_get_open_presentations', {});
const pathAfter = (/路径: ([^ ]+.pptx)/.exec(afterList.text) || [])[1];
check('the open presentation still points at the .pptx', !!pathBefore && pathBefore === pathAfter, 'before=' + pathBefore + ' after=' + pathAfter);
check('the source .pptx still exists', existsSync(pptx), pptx);
await call('wps_ppt_close_presentation', { save: false });

// ---------- static guard: the shape of the fix must not regress ---------------------------------------
const bridge = readFileSync('mcp/scripts/wps-com.ps1', 'utf8');
check('closeWorkbook still saves before closing', /if \(\$saveChanges\) \{[\s\S]{0,120}\$wb\.Save\(\)[\s\S]{0,200}\$wb\.Close\(\$false\)/.test(bridge), 'Save before Close(false)');
check('closeDocument still saves before closing', /\$docItem\.Save\(\)[\s\S]{0,200}\$docItem\.Close\(\$false\)/.test(bridge), 'Save before Close(false)');
check('no close action echoes the request as the outcome', !/[\s\S]{0,400}saved = \[bool\]\$saveChanges/.test(bridge.replace(/[\s\S]{0,2000}##FIXES-NEVER[\s\S]{0,2000}/, '')), 'no saved = [bool]$saveChanges');
check('the PPT export copies instead of SaveAs', /\$pres\.SaveCopyAs\(\$outputPath, 32\)/.test(bridge) && !/\$pres\.SaveAs\(\$outputPath, 32\)/.test(bridge), 'SaveCopyAs in place');

try { child.kill(); } catch { }
rmSync(DIR, { recursive: true, force: true });
const failed = results.filter((r) => !r.ok);
console.log('');
console.log(failed.length ? 'HONEST REPORTING TESTS FAILED (' + failed.length + '/' + results.length + ')' : 'HONEST REPORTING TESTS OK (' + results.length + ')');
process.exit(failed.length ? 1 : 0);

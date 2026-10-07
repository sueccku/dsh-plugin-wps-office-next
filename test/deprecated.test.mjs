// D1 / FIXES 76: the 18 merged-away tool names were kept callable for one release cycle (v0.2.0 →
// v0.4.0). That compatibility window is closed in v0.5.0: the aliases are gone, so an old name is now
// simply an unknown tool. This test needs no WPS - it only exercises dispatch and discovery.
// Run: node test/deprecated.test.mjs
import { spawn } from 'node:child_process';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
// FIXES 95（W1-4）：这个文件的核对源（真实报错 + 文件系统）本身是独立的，但有一处缺口：
// 它只验证规范名 `wps_excel_set_zoom` **能解析**（出现在 wps_help 里），从没验证它**还能干活**。
// 别名删除这类改动的典型风险正是两头断：旧名没了、新名也没接上 —— 名字在帮助里照样漂漂亮亮。
// 正则闸门（要求纯数字）：com() 读失败返回空串，Number('') === 0 会让断言**假通过**。
import { com } from './lib/oracle.mjs';

const OLD_NAMES = [
  'wps_excel_zoom', 'wps_excel_auto_fill', 'wps_excel_insert_row', 'wps_excel_insert_column',
  'wps_excel_delete_row', 'wps_excel_delete_column', 'wps_excel_hide_row',
  'wps_ppt_add_speaker_notes', 'wps_ppt_insert_slide_image', 'wps_ppt_insert_image',
  'wps_ppt_set_animation', 'wps_ppt_set_transition', 'wps_ppt_set_background', 'wps_ppt_add_chart',
  'wps_ppt_duplicate_slide', 'wps_ppt_align_objects',
  'wps_word_generate_doc_toc', 'wps_word_set_font_style',
];

const child = spawn(process.execPath, ['mcp/dist/index.js'], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
let buf = '';
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + '\n'); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); }); }
child.stdout.on('data', (d) => { buf += d.toString(); let i; while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on('data', () => {});
await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'deprecated', version: '1' } });
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

let unknown = 0;
let notFound = 0;
for (const old of OLD_NAMES) {
  const viaCall = await call('wps_call', { tool: old, args: {} });
  if (viaCall.isErr && /未知工具/.test(viaCall.text)) unknown++;
  const viaHelp = await call('wps_help', { tool: old });
  if (viaHelp.isErr && /未找到工具/.test(viaHelp.text)) notFound++;
}
check('every removed alias is an unknown tool through wps_call', unknown === OLD_NAMES.length, unknown + '/' + OLD_NAMES.length);
check('every removed alias is not found through wps_help', notFound === OLD_NAMES.length, notFound + '/' + OLD_NAMES.length);

const canonical = await call('wps_help', { tool: 'wps_excel_set_zoom' });
check('the canonical name still resolves', !canonical.isErr && /wps_excel_set_zoom/.test(canonical.text), canonical.text.slice(0, 70));
// ---- 独立 oracle：规范名不只是"能解析"，还得真的能干原来那个别名的活 ----
const scratch = await call('wps_call', { tool: 'wps_excel_create_workbook', args: {} });
check('a scratch workbook opened for the canonical probe', !scratch.isErr, scratch.text.slice(0, 80));
const zoomSet = await call('wps_excel_set_zoom', { percent: 85 });
check('the canonical tool really runs', !zoomSet.isErr, zoomSet.text.slice(0, 90));
const rawZoom = String(com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application'); [int]$w.ActiveWindow.Zoom")).trim();
check('raw COM says the zoom really changed to 85', /^\d+$/.test(rawZoom) && Number(rawZoom) === 85, 'zoom=' + JSON.stringify(rawZoom));
await call('wps_call', { tool: 'wps_excel_close_workbook', args: { save: false } });

const status = await call('wps_status', {});
check('wps_status no longer reports a deprecated-alias count', !/deprecatedTools/.test(status.text), status.text.slice(0, 90));

// static: the alias table and its consumers are gone for good
check('the alias module is gone from src', !existsSync('mcp/src/tools/deprecated.ts'), 'mcp/src/tools/deprecated.ts');
check('the alias module is gone from dist', !existsSync('mcp/dist/tools/deprecated.js'), 'mcp/dist/tools/deprecated.js');
const srcFiles = [];
const walk = (dir) => { for (const entry of readdirSync(dir, { withFileTypes: true })) { const full = join(dir, entry.name); if (entry.isDirectory()) walk(full); else if (full.endsWith('.ts')) srcFiles.push(full); } };
walk('mcp/src');
const leftovers = srcFiles.filter((f) => /DEPRECATED_TOOLS|DEPRECATED_NAMES|resolveDeprecated/.test(readFileSync(f, 'utf8')));
check('no source file still references the alias table', leftovers.length === 0, leftovers.join(', ') || 'clean');

child.kill();
const failed = results.filter((r) => !r.ok);
console.log('');
console.log(failed.length ? 'DEPRECATION TESTS FAILED (' + failed.length + '/' + results.length + ')' : 'DEPRECATION TESTS OK (' + results.length + ')');
process.exit(failed.length ? 1 : 0);

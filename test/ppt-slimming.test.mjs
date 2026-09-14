// Verifies the P4 PPT slim-down against a real WPS Presentation instance: the four fragmented shape
// setters are now one structured tool, and the dropped families are really gone from the registry.
// Run: node test/ppt-slimming.test.mjs
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

await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'pps', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

check('scratch presentation created', ok(await call('wps_ppt_create_presentation', {})), '');
check('a blank slide is added', ok(await call('wps_ppt_add_slide', { layout: 'blank' })), '');
check('a rectangle is added', ok(await call('wps_ppt_add_shape', { slideIndex: 1, shapeType: 'rectangle' })), '');

const full = await text('wps_ppt_set_shape_effect', {
  slideIndex: 1, shapeIndex: 1,
  shadowEnabled: true, shadowColor: '#333333', shadowBlur: 4, shadowOffsetX: 2, shadowOffsetY: 2,
  borderEnabled: true, borderColor: '#1A365D', borderWidth: 2, borderStyle: 'dash',
  gradientColor1: '#FFFFFF', gradientColor2: '#1A365D', transparency: 0.1,
});
check('the merged tool applies every effect family', full.includes('已更新') && full.includes('shadowColor') && full.includes('borderStyle') && full.includes('gradientColor2') && full.includes('transparency'), full.slice(0, 180));

const partial = await text('wps_ppt_set_shape_effect', { slideIndex: 1, shapeIndex: 1, borderColor: '#FF0000' });
check('a single property only reports that property', partial.includes('borderColor') && !partial.includes('shadow') && !partial.includes('gradient'), partial.slice(0, 140));

const nothing = await call('wps_ppt_set_shape_effect', { slideIndex: 1, shapeIndex: 1 });
check('calling it with no effect is rejected, not a silent no-op', !ok(nothing) && textOf(nothing).includes('nothing to apply'), textOf(nothing).slice(0, 120));

const badStyle = await call('wps_ppt_set_shape_effect', { slideIndex: 1, shapeIndex: 1, borderStyle: 'wavy' });
check('an unknown border style is rejected', !ok(badStyle) && textOf(badStyle).includes('unknown borderStyle'), textOf(badStyle).slice(0, 120));

const removed = ['wps_ppt_set_shape_shadow', 'wps_ppt_set_shape_gradient', 'wps_ppt_set_shape_border', 'wps_ppt_set_shape_transparency', 'wps_ppt_set_3d_rotation', 'wps_ppt_set_3d_depth', 'wps_ppt_set_3d_material'];
for (const name of removed) {
  const res = await call('wps_call', { tool: name, args: {} });
  check(name.replace('wps_ppt_', '') + ' no longer resolves', textOf(res).includes('未知工具'), textOf(res).slice(0, 70));
}

check('presentation closed', ok(await call('wps_ppt_close_presentation', { save: false })), '');
const open = await text('wps_ppt_get_open_presentations', {});
check('no presentation left open', /共?0个|没有打开/.test(open), open.replace(/\n/g, ' | ').slice(0, 70));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'PPT SLIMMING TESTS OK (' + results.length + ')' : 'PPT SLIMMING TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

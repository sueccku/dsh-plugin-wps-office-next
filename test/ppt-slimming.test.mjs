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

const removed = ['wps_ppt_set_shape_shadow', 'wps_ppt_set_shape_gradient', 'wps_ppt_set_shape_border', 'wps_ppt_set_shape_transparency', 'wps_ppt_set_3d_rotation', 'wps_ppt_set_3d_depth', 'wps_ppt_set_3d_material', 'wps_ppt_add_animation_preset', 'wps_ppt_add_emphasis_animation', 'wps_ppt_set_table_style', 'wps_ppt_set_table_cell_style', 'wps_ppt_set_table_row_style', 'wps_ppt_set_slide_number', 'wps_ppt_set_ppt_footer', 'wps_ppt_set_ppt_date_time'];
for (const name of removed) {
  const res = await call('wps_call', { tool: name, args: {} });
  check(name.replace('wps_ppt_', '') + ' no longer resolves', textOf(res).includes('未知工具'), textOf(res).slice(0, 70));
}

// P4 second wave: the merged table / animation / footer tools.
check('a table is inserted (existing tool)', ok(await call('wps_ppt_insert_table', { slideIndex: 1, rows: 2, cols: 3 })), '');
const cellScope = await text('wps_ppt_set_table_format', { slideIndex: 1, tableIndex: 1, row: 1, col: 1, backgroundColor: '#D9E2F3', bold: true });
check('table format styles a single cell', cellScope.includes('作用域 单元格') && cellScope.includes('backgroundColor') && cellScope.includes('bold'), cellScope.slice(0, 150));
const rowScope = await text('wps_ppt_set_table_format', { slideIndex: 1, tableIndex: 1, row: 2, fontColor: '#1A365D' });
check('table format styles a whole row when only row is given', rowScope.includes('作用域 整行') && rowScope.includes('fontColor'), rowScope.slice(0, 150));
const geometry = await text('wps_ppt_set_table_format', { slideIndex: 1, tableIndex: 1, left: 120, top: 90 });
check('table format moves the table', geometry.includes('left') && geometry.includes('top') && geometry.includes('仅位置尺寸'), geometry.slice(0, 150));
const tableNothing = await call('wps_ppt_set_table_format', { slideIndex: 1, tableIndex: 1 });
check('table format with nothing to change is rejected', !ok(tableNothing) && textOf(tableNothing).includes('nothing to apply'), textOf(tableNothing).slice(0, 120));
const footer = await text('wps_ppt_set_slide_footer', { showSlideNumber: true, footerText: '内部资料', showDate: true });
check('footer tool sets all three parts at once', footer.includes('slideNumber') && footer.includes('footer') && footer.includes('dateTime') && footer.includes('内部资料'), footer.replace(/\n/g, ' | ').slice(0, 170));
const footerNothing = await call('wps_ppt_set_slide_footer', {});
check('footer tool with nothing to change is rejected', !ok(footerNothing) && textOf(footerNothing).includes('nothing to apply'), textOf(footerNothing).slice(0, 120));
const animNothing = await call('wps_ppt_add_animation', { slideIndex: 1 });
check('animation without shape or preset is rejected', !ok(animNothing) && textOf(animNothing).includes('shapeIndex/shapeName is required'), textOf(animNothing).slice(0, 120));

check('presentation closed', ok(await call('wps_ppt_close_presentation', { save: false })), '');
const open = await text('wps_ppt_get_open_presentations', {});
check('no presentation left open', /共?0个|没有打开/.test(open), open.replace(/\n/g, ' | ').slice(0, 70));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'PPT SLIMMING TESTS OK (' + results.length + ')' : 'PPT SLIMMING TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

// 回归（S3 走查 O-3 / O-4）：演示文稿的**页码越界**必须给明确中文，不能把 WPS 的 E_FAIL 漏给用户；
// 形状定位的 `shapeName` 必须真的能从工具面走通（此前 schema 只认必填 shapeIndex，桥里的名称分支是死代码）。
//
// 背景：
//   O-3 真机实测（attempt 2 第 118 步）11 页的文稿问 get_slide_info{slideIndex:12} →
//       `获取幻灯片信息失败: 对 COM 组件的调用返回了错误 HRESULT E_FAIL。`
//       文案还引导去查"对象/名称/序号是否正确"，而根因只是页码越界。
//   O-4 `wps_ppt_set_shape_position` 的 schema required 里有 shapeIndex，桥里却有 name 分支 →
//       按技能文档说的"用名称定位形状"根本走不通（参数在工具层被丢掉，报缺少 shapeIndex）。
//
// Run: node test/ppt-slide-index.test.mjs   (needs WPS)
import { spawn } from 'node:child_process';

const child = spawn(process.execPath, ['mcp/dist/index.js'], {
  stdio: ['pipe', 'pipe', 'pipe'],
  windowsHide: true,
  env: { ...process.env, WPS_OFFICE_HOST_SCRIPT: 'host/wps-com-host.ps1' },
});
let buf = '';
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + '\n'); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); }); }
child.stdout.on('data', (d) => { buf += d.toString(); let i; while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on('data', () => {});
await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'ppt-slide-index', version: '1' } });
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

let id = 10;
const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }
async function call(name, args) {
  const r = await req(id++, 'tools/call', { name, arguments: args || {} });
  const isErr = !!(r.result && r.result.isError);
  const text = String((r.result && r.result.content && r.result.content[0].text) || JSON.stringify((r && r.error) || {})).replace(/\s+/g, ' ');
  return { isErr, text };
}

// scratch deck: 1 title + 1 titleContent = 2 pages
await call('wps_ppt_create_presentation', {});
await call('wps_ppt_add_slide', { layout: 'title', title: '越界回归' });
await call('wps_ppt_add_slide', { layout: 'titleContent', title: '第二页' });

// ---- O-3：页码越界 ----
const over = await call('wps_ppt_get_slide_info', { slideIndex: 99 });
check('out-of-range slide index is refused with a clear message',
  over.isErr && /没有第 99 页/.test(over.text) && /共 2 页/.test(over.text), over.text.slice(0, 140));
check('the refusal is not a raw COM E_FAIL', !/E_FAIL|HRESULT|0x800/i.test(over.text), over.text.slice(0, 140));

const zero = await call('wps_ppt_get_slide_info', { slideIndex: 0 });
check('slide index 0 is refused as "starts at 1"', zero.isErr && /从 1 开始/.test(zero.text), zero.text.slice(0, 120));

const nonNum = await call('wps_ppt_get_slide_info', { slideIndex: 'two' });
check('a non-numeric slide index is refused as such', nonNum.isErr && /必须是整数/.test(nonNum.text), nonNum.text.slice(0, 120));

const valid = await call('wps_ppt_get_slide_info', { slideIndex: 2 });
check('a valid slide index still works', !valid.isErr && /第二页|第 2 页/.test(valid.text), valid.text.slice(0, 100));

const moveOver = await call('wps_ppt_move_slide', { fromIndex: 99, toIndex: 1 });
check('out-of-range fromIndex is refused before touching the deck',
  moveOver.isErr && /没有第 99 页/.test(moveOver.text) && !/E_FAIL|HRESULT/.test(moveOver.text), moveOver.text.slice(0, 140));

const still = await call('wps_ppt_get_slide_count', {});
check('the deck is untouched by the refused move (still 2 pages)', !still.isErr && /=?\s*2/.test(still.text), still.text.slice(0, 80));

// ---- O-4：按名称定位形状 ----
const shapes = await call('wps_ppt_get_shapes', { slideIndex: 2 });
// 形状名可能带空格（"标题 1"），列表格式是 `[1] 标题 1 (14) 位置: ...`
const nameMatch = /\[1\]\s*(.+?)\s*\(\d+\)/.exec(shapes.text);
const shapeName = nameMatch ? nameMatch[1] : null;
check('get_shapes reports a usable shape name', !!shapeName, 'name=' + shapeName + ' :: ' + shapes.text.slice(0, 90));

const byName = await call('wps_ppt_set_shape_position', { slideIndex: 2, shapeName, left: 40, top: 60 });
check('set_shape_position accepts shapeName (no shapeIndex)', !byName.isErr, byName.text.slice(0, 140));

const neither = await call('wps_ppt_set_shape_position', { slideIndex: 2, left: 40, top: 60 });
check('missing both shapeIndex and shapeName is a clear error',
  neither.isErr && /shapeIndex|shapeName/.test(neither.text), neither.text.slice(0, 120));

const byIndex = await call('wps_ppt_set_shape_position', { slideIndex: 2, shapeIndex: 1, left: 40, top: 60 });
check('set_shape_position still accepts shapeIndex', !byIndex.isErr, byIndex.text.slice(0, 120));

// ---- dead geometry params on set_image_style must not come back ----
const imgStyle = await call('wps_ppt_set_image_style', { slideIndex: 2, imageIndex: 1, style: { rotation: 0 } });
check('set_image_style still accepts its documented style object',
  !imgStyle.isErr || /没有|no picture|imageIndex/.test(imgStyle.text), imgStyle.text.slice(0, 120));

await call('wps_ppt_close_presentation', { save: false });
child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'PPT SLIDE INDEX TESTS OK (' + results.length + ')' : 'PPT SLIDE INDEX TESTS FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

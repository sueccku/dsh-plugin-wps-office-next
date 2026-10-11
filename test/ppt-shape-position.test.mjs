// 回归：PPT 侧的**几何写入必须走 COM 边界助手**，不能把参数直接赋给 COM 属性。
//
// 背景（S3 场景真机走查抓到的 P0）：
//   宿主收到请求后会把参数 `$req.params | ConvertTo-Json -Compress` 再在动作里
//   `ConvertFrom-Json`。PS 5.1 的这条往返会把**带小数的 JSON 数字解析成 System.Decimal**
//   （整数仍是 Int32）。而 setShapePosition 写的是
//       if ($null -ne $p.left) { $shape.Left = $p.left }
//   —— Decimal 赋给 COM 的 Single 属性直接抛 `InvalidCastException: 指定的转换无效`，
//   于是"任何带小数的坐标/尺寸都设不进去"，还报成"确认操作对象存在、名称或序号正确"。
//   同一文件早就把这条写成约定（COM boundary helpers 注释：每个被写的 COM 成员都必须走
//   Set-ComValue，它每次重新绑定并归一化值），setShapePosition / setImageStyle /
//   setTableFormat / insertImage 是漏网的存量。
//
// 判据以**裸 COM 读回**为准（不读插件自述）：修复前小数用例失败，修复后必须真的落盘。
// Run: node test/ppt-shape-position.test.mjs   (needs WPS)
import { spawn, spawnSync } from "node:child_process";
import { mkdirSync, rmSync, writeFileSync } from "node:fs";
import { resolve as resolvePath } from "node:path";

const ART = resolvePath("test/.artifacts");
mkdirSync(ART, { recursive: true });
// 1x1 PNG：给 setImageStyle 用例一张真图（避免依赖任何场景产物）。
const PNG = resolvePath("test/.artifacts/ppt-shape-position-1x1.png");
writeFileSync(PNG, Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==", "base64"));

const child = spawn(process.execPath, ["mcp/dist/index.js"], {
  stdio: ["pipe", "pipe", "pipe"],
  windowsHide: true,
  env: { ...process.env, WPS_OFFICE_HOST_SCRIPT: resolvePath("host/wps-com-host.ps1") },
});
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

const results = [];
function check(name, okFlag, detail) { results.push({ name, ok: !!okFlag }); console.log((okFlag ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "ppt-shape-position", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });

// 裸 COM 读回（只读）。base64 进出，避免中文/编码链被改坏。
function com(script) {
  const b64 = Buffer.from(script, "utf8").toString("base64");
  const r = spawnSync("powershell", ["-NoProfile", "-STA", "-Command",
    "$s = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + b64 + "')); $out = Invoke-Expression $s; [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$out))"],
    { encoding: "utf8", windowsHide: true });
  try { return Buffer.from((r.stdout || "").trim(), "base64").toString("utf8"); } catch { return ""; }
}
const geom = (slideIndex, shapeIndex) => com(
  "$p=[Runtime.InteropServices.Marshal]::GetActiveObject('Kwpp.Application');" +
  "$s=$p.ActivePresentation.Slides.Item(" + slideIndex + ").Shapes.Item(" + shapeIndex + ");" +
  "[string]$s.Left + '|' + [string]$s.Top + '|' + [string]$s.Width + '|' + [string]$s.Height");
const nums = (s) => s.split("|").map((x) => Number(x));

// ---------- scratch deck ----------
check("create_presentation", ok(await call("wps_ppt_create_presentation", {})), "");
check("add_slide(titleContent)", ok(await call("wps_ppt_add_slide", { layout: "titleContent" })), "");

// ---------- 1) 整数坐标：修复前后都应通过（先跑它，确保对照干净） ----------
const intCase = await call("wps_ppt_set_shape_position", { slideIndex: 1, shapeIndex: 1, left: 100, top: 100 });
check("integer left/top is accepted", ok(intCase), text(intCase).replace(/\s+/g, " ").slice(0, 90));

// ---------- 2) 小数坐标：宿主 JSON 往返后会变成 Decimal —— 这就是回归点 ----------
const fracCase = await call("wps_ppt_set_shape_position", { slideIndex: 1, shapeIndex: 1, left: 100.5, top: 120.25 });
check("fractional left/top is accepted", ok(fracCase), text(fracCase).replace(/\s+/g, " ").slice(0, 90));

// 立刻裸 COM 读回：必须真的落盘成小数，而不是"回 success 但没动"。
const back = nums(geom(1, 1));
check("raw COM left is 100.5", Math.abs(back[0] - 100.5) < 0.05, "left=" + back[0]);
check("raw COM top is 120.25", Math.abs(back[1] - 120.25) < 0.05, "top=" + back[1]);

// ---------- 3) 小数宽高 ----------
const sizeCase = await call("wps_ppt_set_shape_position", { slideIndex: 1, shapeIndex: 1, left: 100.5, top: 120.25, width: 300.5, height: 90.75 });
check("fractional width/height is accepted", ok(sizeCase), text(sizeCase).replace(/\s+/g, " ").slice(0, 90));
const back2 = nums(geom(1, 1));
check("raw COM width is 300.5", Math.abs(back2[2] - 300.5) < 0.05, "width=" + back2[2]);
check("raw COM height is 90.75", Math.abs(back2[3] - 90.75) < 0.05, "height=" + back2[3]);

// ---------- 4) 兄弟动作：同族漏网站点（insertPptImage / insertPptTable / setPptTableFormat） ----------
const img = await call("wps_ppt_insert_ppt_image", { slideIndex: 1, imagePath: PNG, left: 10, top: 10 });
check("insert_ppt_image", ok(img), text(img).replace(/\s+/g, " ").slice(0, 80));
const imgFrac = await call("wps_ppt_insert_ppt_image", { slideIndex: 1, imagePath: PNG, left: 60.5, top: 70.25 });
check("insert_ppt_image accepts fractional left/top", ok(imgFrac), text(imgFrac).replace(/\s+/g, " ").slice(0, 90));

const tbl = await call("wps_ppt_insert_table", { slideIndex: 1, rows: 2, cols: 2, left: 200, top: 200 });
check("insert_table", ok(tbl), text(tbl).replace(/\s+/g, " ").slice(0, 80));
const tblFrac = await call("wps_ppt_insert_table", { slideIndex: 1, rows: 2, cols: 2, left: 220.5, top: 240.25 });
check("insert_table accepts fractional left/top", ok(tblFrac), text(tblFrac).replace(/\s+/g, " ").slice(0, 90));
const tblFmt = await call("wps_ppt_set_table_format", { slideIndex: 1, left: 220.5, top: 240.25, row: 1, col: 1, bold: true });
check("set_table_format accepts fractional left/top", ok(tblFmt), text(tblFmt).replace(/\s+/g, " ").slice(0, 90));

// ---------- cleanup ----------
const closed = await call("wps_ppt_close_presentation", { save: false });
check("scratch presentation closed", ok(closed), text(closed).replace(/\s+/g, " ").slice(0, 80));
try { rmSync(PNG, { force: true }); } catch { /* ignore */ }

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "PPT SHAPE POSITION TESTS OK (" + results.length + ")" : "PPT SHAPE POSITION TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

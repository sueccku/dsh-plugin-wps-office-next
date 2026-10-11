// Live verification of the Excel range I/O fixes:
//  - omitted sheet now resolves to the active sheet instead of index 0
//  - range read is one COM call with explicit 2D marshalling (was cell-by-cell)
// Run: node test/excel-range.test.mjs
import { createRequire } from "node:module";
import { mkdirSync, rmSync } from "node:fs";
import { dirname, resolve } from "node:path";
// P3：范围 I/O 以前只用同一个 comHost 读回（读写同源会一起错）。现在存盘后由 openpyxl 读**文件**。
import { xlsx } from "./lib/oracle.mjs";
const require = createRequire(import.meta.url);
const { comHost } = require("../mcp/dist/client/com-host.js");

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }

await comHost.invoke("createWorkbook", {});
try {
  // write without sheet -> active sheet
  const w = await comHost.invoke("setRangeData", { range: "A1:B2", data: [[11, 22], [33, 44]] });
  check("write_range without sheet", w.success === true, JSON.stringify(w).slice(0, 120));

  const noSheet = await comHost.invoke("getRangeData", { range: "A1:B2" });
  check("read_range without sheet", noSheet.success === true, JSON.stringify(noSheet).slice(0, 140));
  const d1 = noSheet.data && noSheet.data.data;
  // FIXES 93（W1-1）：以前只查 d1[0].length —— 第 2..n 行宽度错了不会红。现在逐行查。
check("2D shape is rows x cols", Array.isArray(d1) && d1.length === 2 && d1.every((r) => Array.isArray(r) && r.length === 2), JSON.stringify(d1));
  check("values round-trip", JSON.stringify(d1) === JSON.stringify([[11, 22], [33, 44]]), JSON.stringify(d1));

  const withSheet = await comHost.invoke("getRangeData", { sheet: 1, range: "A1:B2" });
  check("read_range with sheet", withSheet.success === true && JSON.stringify(withSheet.data.data) === JSON.stringify([[11, 22], [33, 44]]));

  // mixed string/number rows: this used to fail with an Int32 -> String cast error because
  // ConvertFrom-Json yields PSObject-wrapped numbers
  const mixed = [["h1", "h2", "h3"], [1, 2, 3], [4.5, 6, 7]];
  const wm = await comHost.invoke("setRangeData", { range: "F1:H3", data: mixed });
  check("write mixed-type block", wm.success === true, JSON.stringify(wm).slice(0, 140));
  const rm = await comHost.invoke("getRangeData", { range: "F1:H3" });
  check("mixed block round-trips exactly", rm.success === true && JSON.stringify(rm.data.data) === JSON.stringify(mixed), JSON.stringify(rm.data && rm.data.data));

  const one = await comHost.invoke("getRangeData", { range: "A1" });
  check("single cell read stays 1x1", one.success === true && JSON.stringify(one.data.data) === JSON.stringify([[11]]), JSON.stringify(one.data));

  // performance: one COM call for a large block
  const rows = 400, cols = 20;
  const colLetter = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const address = "D1:" + colLetter(3 + cols) + rows;
  const block = [];
  for (let r = 0; r < rows; r++) { const row = []; for (let c = 0; c < cols; c++) row.push(r * cols + c); block.push(row); }
  const wb = await comHost.invoke("setRangeData", { range: address, data: block });
  check("write large block", wb.success === true, address + "  " + JSON.stringify(wb).slice(0, 80));
  const t0 = Date.now();
  const big = await comHost.invoke("getRangeData", { range: address });
  const ms = Date.now() - t0;
  const db = big.data && big.data.data;
  check("read " + (rows * cols) + " cells in one call", big.success === true && db.length === rows && db.every((r) => Array.isArray(r) && r.length === cols), ms + "ms");
  // 墙钟断言放宽到 10s：2s 在负载高的机器上会假红，而这条的本意只是"别退化成逐格往返"。
  check("large read in one call, not cell-by-cell", ms < 10000, ms + "ms  (cell-by-cell would be ~" + (rows * cols) + " COM round trips)");
  // FIXES 93（W1-2）：以前只看最后一个格子 —— 中间任何一格错了都不会红。现在整块逐格比对。
check("large read values intact", JSON.stringify(db) === JSON.stringify(block), "first mismatch at " + (() => { for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) if (db[r][c] !== block[r][c]) return "r" + r + "c" + c + " got " + db[r][c] + " want " + block[r][c]; return "none"; })());
  // ---- 独立 oracle（P3）：以**保存后的文件**为准，而不是用同一个 comHost 读回 ----
  const XL = resolve("test/.artifacts/excel-range.xlsx");
  rmSync(XL, { force: true });                       // 先删旧产物：残留文件不能替失败背书
  mkdirSync(dirname(XL), { recursive: true });
  const savedBook = await comHost.invoke("saveAs", { outputPath: XL, format: "xlsx" });
  check("workbook saved for independent verification", savedBook.success === true, JSON.stringify(savedBook).slice(0, 120));
  const disk = xlsx(XL, "A1:B2");
  check("the file carries the small block the plugin wrote", !!disk && JSON.stringify(disk.values) === JSON.stringify([[11, 22], [33, 44]]), JSON.stringify(disk && disk.values));
  const diskBig = xlsx(XL, "D1:W400");
  const dims = diskBig ? diskBig.values.length + "x" + ((diskBig.values[0] || []).length) : "null";
  check("the file carries the whole 400x20 block", !!diskBig && diskBig.values.length === 400 && diskBig.values.every((r) => r.length === 20), dims);
  check("the file's first and last big-block cells survived", !!diskBig && Number(diskBig.values[0][0]) === 0 && Number(diskBig.values[399][19]) === 7999, diskBig ? JSON.stringify([diskBig.values[0][0], diskBig.values[399][19]]) : "null");
} finally {
  await comHost.invoke("closeWorkbook", { saveChanges: false });
}
await comHost.stop();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "EXCEL RANGE TESTS OK (" + results.length + ")" : "EXCEL RANGE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

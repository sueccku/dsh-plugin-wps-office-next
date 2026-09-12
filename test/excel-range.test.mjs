// Live verification of the Excel range I/O fixes:
//  - omitted sheet now resolves to the active sheet instead of index 0
//  - range read is one COM call with explicit 2D marshalling (was cell-by-cell)
// Run: node test/excel-range.test.mjs
import { createRequire } from "node:module";
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
  check("2D shape is rows x cols", Array.isArray(d1) && d1.length === 2 && Array.isArray(d1[0]) && d1[0].length === 2, JSON.stringify(d1));
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
  check("read " + (rows * cols) + " cells in one call", big.success === true && db.length === rows && db[0].length === cols, ms + "ms");
  check("large read under 2s", ms < 2000, ms + "ms  (cell-by-cell would be ~" + (rows * cols) + " COM round trips)");
  check("large read values intact", db[rows - 1][cols - 1] === rows * cols - 1, "last=" + db[rows - 1][cols - 1]);
} finally {
  await comHost.invoke("closeWorkbook", { saveChanges: false });
}
await comHost.stop();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "EXCEL RANGE TESTS OK (" + results.length + ")" : "EXCEL RANGE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

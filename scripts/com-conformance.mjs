// Machine COM conformance self-check for dsh-plugin-wps-office-next.
// Exercises every PowerShell COM adapter behaviour the action layer depends on, on THIS machine.
// A machine that passes this file is expected to run the plugin correctly; a failure means the
// environment differs and should be reported rather than silently corrupting documents.
//
// Run: node scripts/com-conformance.mjs
import { createRequire } from "node:module";
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";

const require = createRequire(import.meta.url);
const { comHost } = require("../mcp/dist/client/com-host.js");
const { POWERSHELL_EXE } = require("../mcp/dist/client/com-host.js");

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
const val = (r) => (r && r.data ? r.data.value : undefined);

console.log("powershell: " + POWERSHELL_EXE + (existsSync(POWERSHELL_EXE) ? "" : "  (MISSING)"));

const env = await comHost.invoke("__env", {});
check("host reports Windows PowerShell 5.1", env.success === true && Number(env.data.psMajor) === 5, JSON.stringify(env.data));
check("host runs the pinned interpreter", String(env.data && env.data.exe || "").toLowerCase().includes("windowspowershell"), String(env.data && env.data.exe || ""));

// binder-cache matrix: alternate value types through the same cell
await comHost.invoke("createWorkbook", {});
const seq = [["number", 42, 42], ["text", "hello", "hello"], ["decimal", 3.5, 3.5], ["text2", "world", "world"], ["bool", true, true], ["number2", 7, 7]];
let seqOk = true;
let seqDetail = "";
for (const [label, v, expected] of seq) {
  const set = await comHost.invoke("setCellValue", { sheet: 1, row: 1, col: 1, value: v });
  const got = val(await comHost.invoke("getCellValue", { sheet: 1, row: 1, col: 1 }));
  const same = got === expected || String(got) === String(expected);
  if (!set.success || !same) { seqOk = false; seqDetail += label + " set=" + set.success + " read=" + JSON.stringify(got) + "  "; }
}
check("alternating value types through one cell", seqOk, seqDetail || "6 types round-tripped");

const mixed = [["h1", "h2", "h3"], [1, 2, 3], [4.5, 6, 7], ["x", "y", "z"]];
const w = await comHost.invoke("setRangeData", { range: "A1:C4", data: mixed });
const back = await comHost.invoke("getRangeData", { range: "A1:C4" });
check("mixed 2D matrix write/read", w.success === true && JSON.stringify(back.data.data) === JSON.stringify(mixed), JSON.stringify(back.data && back.data.data));

const one = await comHost.invoke("getRangeData", { range: "A1:B1" });
check("2D read keeps row/column shape", one.success === true && one.data.data.length === 1 && one.data.data[0].length === 2, JSON.stringify(one.data && one.data.data));

const fr = await comHost.invoke("findReplaceExcel", { findText: "h1", replaceText: "H1" });
check("excel find/replace with mixed argument types", fr.success === true && Number(fr.data.cells) >= 1, JSON.stringify(fr.data));

await comHost.invoke("createDocument", {});
await comHost.invoke("insertText", { text: "conformance probe", position: "start" });
const ls = await comHost.invoke("setLineSpacing", { lineSpacing: 1.5 });
check("word numeric argument", ls.success === true && Number(ls.data.applied) >= 1, JSON.stringify(ls.data));
const tc = await comHost.invoke("setTextColor", { color: "#FF0000", range: "all" });
check("word mixed value assignment", tc.success === true, JSON.stringify(tc).slice(0, 100));

await comHost.invoke("createPresentation", {});
await comHost.invoke("addSlide", { layout: "title_content", title: "probe", content: "probe" });
const size = await comHost.invoke("setSlideSize", { width: 1280, height: 720 });
check("ppt numeric argument", size.success === true && Math.round(size.data.slideWidth) === 960, JSON.stringify(size.data));
const shape = await comHost.invoke("addShape", { slideIndex: 1, type: "rectangle", left: 80, top: 80, width: 200, height: 100, text: "probe" });
check("ppt shape creation", shape.success === true, JSON.stringify(shape).slice(0, 100));
const fill = await comHost.invoke("setShapeFill", { slideIndex: 1, shapeIndex: shape.data ? 1 : 1, color: "#00AA00" });
check("ppt colour assignment", fill.success === true, JSON.stringify(fill).slice(0, 100));

await comHost.stop();

// leave the machine clean: close scratch documents without saving
const sweep = [
  "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8",
  "foreach ($progid in 'Ket.Application','Kwps.Application','Kwpp.Application') {",
  "  try {",
  "    $a=[System.Runtime.InteropServices.Marshal]::GetActiveObject($progid)",
  "    try { $a.DisplayAlerts=$false } catch {}",
  "    if ($progid -eq 'Ket.Application') { while ($a.Workbooks.Count -gt 0) { try { $a.Workbooks.Item(1).Close($false) } catch { break } } }",
  "    elseif ($progid -eq 'Kwps.Application') { while ($a.Documents.Count -gt 0) { try { $a.Documents.Item(1).Close($false) } catch { break } } }",
  "    else { while ($a.Presentations.Count -gt 0) { try { $a.Presentations.Item(1).Close() } catch { break } } }",
  "  } catch {}",
  "}",
  "'swept'"
].join("\n");
spawnSync(POWERSHELL_EXE, ["-NoProfile", "-STA", "-EncodedCommand", Buffer.from(sweep, "utf16le").toString("base64")], { windowsHide: true });

const failed = results.filter((r) => !r.ok).length;
console.log("");
console.log(failed === 0 ? "COM CONFORMANCE OK (" + results.length + " checks) on this machine" : "COM CONFORMANCE FAILED (" + failed + "/" + results.length + ") on this machine");
process.exit(failed === 0 ? 0 : 1);

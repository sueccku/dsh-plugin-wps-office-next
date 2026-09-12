// Live verification of the ten actions that upstream tools referenced but the bridge never
// implemented, plus the closePresentation saveChanges fix.
// Creates throwaway documents, exercises each action, then closes them without saving.
// Run: node test/new-actions.test.mjs
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const { comHost } = require("../mcp/dist/client/com-host.js");

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
const unknownText = (r) => r && typeof r.error === "string" && r.error.includes("Unknown action");
function verdict(name, r, extra) {
  const ok = !!r && r.success === true && !unknownText(r);
  check(name, ok, (ok ? "" : JSON.stringify(r).slice(0, 160)) + (extra ? "  " + extra : ""));
  return r;
}

await comHost.invoke("ping", {});

// ---------- Excel ----------
try {
  verdict("createWorkbook", await comHost.invoke("createWorkbook", {}));
  verdict("setRangeData", await comHost.invoke("setRangeData", { sheet: 1, range: "A1:B2", data: [[1, 2], [3, 4]] }));
  const sum = verdict("autoSum", await comHost.invoke("autoSum", { sheet: 1, range: "A1:B2", targetCell: "A3" }));
  check("autoSum wrote a SUM formula", sum.success && String(sum.data && sum.data.formula).toUpperCase().includes("SUM"), JSON.stringify(sum.data));
  const ev = await comHost.invoke("evaluateFormula", { formula: "=1+2" });
  verdict("evaluateFormula", ev, "result=" + JSON.stringify(ev.data && ev.data.result) + " via " + (ev.data && ev.data.method));
  check("evaluateFormula returned 3", ev.success && Number(ev.data && ev.data.result) === 3);
  const zoom = verdict("setZoom", await comHost.invoke("setZoom", { percent: 150 }));
  check("setZoom applied 150", zoom.success && Number(zoom.data && zoom.data.zoom) === 150);
  await comHost.invoke("setZoom", { percent: 100 });
} finally {
  verdict("closeWorkbook(discard)", await comHost.invoke("closeWorkbook", { saveChanges: false }));
}

// ---------- Word ----------
try {
  verdict("createDocument", await comHost.invoke("createDocument", {}));
  verdict("insertText", await comHost.invoke("insertText", { text: "probe paragraph", position: "start" }));
  const ls = verdict("setLineSpacing", await comHost.invoke("setLineSpacing", { lineSpacing: 1.5 }));
  check("setLineSpacing reported an applied count", ls.success && Number(ls.data && ls.data.applied) >= 1, JSON.stringify(ls.data));
  verdict("setTextColor(hex)", await comHost.invoke("setTextColor", { color: "#FF0000", range: "all" }));
  verdict("setTextColor(name)", await comHost.invoke("setTextColor", { color: "blue", range: "all" }));
  const bad = await comHost.invoke("setTextColor", { color: "definitely-not-a-color" });
  check("setTextColor rejects a bad color loudly", bad.success === false && !unknownText(bad), JSON.stringify(bad).slice(0, 100));
  const sb = verdict("insertSectionBreak", await comHost.invoke("insertSectionBreak", { breakType: "nextPage" }));
  check("insertSectionBreak created a section", sb.success && Number(sb.data && sb.data.sections) >= 2, "sections=" + (sb.data && sb.data.sections));
} finally {
  verdict("closeDocument(discard)", await comHost.invoke("closeDocument", { saveChanges: false }));
}

// ---------- PowerPoint ----------
try {
  verdict("createPresentation", await comHost.invoke("createPresentation", {}));
  verdict("addSlide(seed)", await comHost.invoke("addSlide", { layout: "title_content", title: "probe", content: "probe" }));
  const counts = await comHost.invoke("getSlideCount", {});
  console.log("      (slides after addSlide: " + JSON.stringify(counts.data).slice(0, 120) + ")");
  const size = verdict("setSlideSize", await comHost.invoke("setSlideSize", { width: 1280, height: 720 }));
  check("setSlideSize converted px to points", size.success && Math.round(size.data.slideWidth) === 960 && Math.round(size.data.slideHeight) === 540, JSON.stringify(size.data));
  verdict("addShape", await comHost.invoke("addShape", { slideIndex: 1, type: "rectangle", left: 80, top: 80, width: 240, height: 120, text: "probe" }));
  const shapes = await comHost.invoke("getShapes", { slideIndex: 1 });
  const hasShape = shapes.success && Array.isArray(shapes.data && shapes.data.shapes) && shapes.data.shapes.length > 0;
  console.log("      (slide 1 shapes: " + (hasShape ? shapes.data.shapes.length : 0) + ")");
  if (hasShape) {
    verdict("setShapeFill", await comHost.invoke("setShapeFill", { slideIndex: 1, shapeIndex: 1, color: "#FF8800" }));
    verdict("setFontColor", await comHost.invoke("setFontColor", { slideIndex: 1, shapeIndex: 1, color: "#FFFFFF" }));
  } else {
    check("setShapeFill (needs a shape)", false, "no shapes on slide 1 of the new presentation");
    check("setFontColor (needs a shape)", false, "no shapes on slide 1 of the new presentation");
  }
  const theme = await comHost.invoke("setSlideTheme", { theme: "C:/definitely/not/a/theme.thmx" });
  check("setSlideTheme exists and fails honestly", theme.success === false && !unknownText(theme) && String(theme.error).includes("not found"), JSON.stringify(theme).slice(0, 120));
} finally {
  verdict("closePresentation(discard)", await comHost.invoke("closePresentation", { saveChanges: false }));
}

await comHost.stop();
const failed = results.filter((r) => !r.ok);
console.log(failed.length === 0 ? "NEW ACTION TESTS OK (" + results.length + ")" : "NEW ACTION TESTS FAILED (" + failed.length + "/" + results.length + ")");
process.exit(failed.length === 0 ? 0 : 1);

// Excel find/replace regression: correct action, correct count, and no cross-application damage.
// The upstream Excel tool called the Word-only findReplace action, so "find and replace in Excel"
// silently edited the Word document instead.
// Run: node test/find-replace.test.mjs
import { spawn } from "node:child_process";

const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function ok(res) { return !!(res && res.result && !res.result.isError); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "fr", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params, appType) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {}, appType } });

// --- Excel side ---
await viaAction("createWorkbook", {});
const seed = await call("wps_excel_write_range", { range: "A1:C2", data: [["alpha", "beta", "alpha"], ["gamma", "alpha", "delta"]] });
check("seed excel data", ok(seed), text(seed).slice(0, 80));

const r1 = await call("wps_excel_find_replace", { find: "alpha", replace: "ALPHA" });
const t1 = text(r1);
check("excel find/replace reports success", ok(r1), t1.slice(0, 90));
check("excel find/replace reports the real count", /3/.test(t1), t1.slice(0, 90));

const after1 = text(await call("wps_excel_read_range", { range: "A1:C2" }));
check("excel cells actually changed", after1.includes("ALPHA") && !after1.includes("alpha"), after1.replace(/\s+/g, " ").slice(0, 110));

// --- Word side: must be untouched by the Excel call ---
await viaAction("createDocument", {}, "wps");
await call("wps_word_insert_text", { text: "alpha word alpha", position: "start" });
const wordBefore = text(await call("wps_word_get_document_text", {}));
check("word document seeded", wordBefore.includes("alpha"), wordBefore.replace(/\s+/g, " ").slice(0, 80));

const r2 = await call("wps_excel_find_replace", { find: "ALPHA", replace: "ZZZ" });
check("excel second replace runs", ok(r2), text(r2).slice(0, 80));
const wordAfter = text(await call("wps_word_get_document_text", {}));
check("CROSS-APP: word document untouched by excel replace", wordAfter.includes("alpha") && !wordAfter.includes("ZZZ"), wordAfter.replace(/\s+/g, " ").slice(0, 90));
const excelAfter = text(await call("wps_excel_read_range", { range: "A1:C2" }));
check("excel second replace applied", excelAfter.includes("ZZZ"), excelAfter.replace(/\s+/g, " ").slice(0, 90));

// --- Word find/replace still works ---
const r3 = await call("wps_word_find_replace", { find_text: "alpha", replace_text: "omega", replace_all: true });
check("word find/replace runs", ok(r3), text(r3).slice(0, 90));
const wordFinal = text(await call("wps_word_get_document_text", {}));
check("word text actually replaced", wordFinal.includes("omega") && !wordFinal.includes("alpha"), wordFinal.replace(/\s+/g, " ").slice(0, 90));

// --- Search-only must never erase what it found ---
// The bridge used to apply the tool's empty replaceText as a replacement, so a plain search
// deleted every match and then reported a fabricated count.
await call("wps_word_find_replace", { find_text: "omega", replace_text: "alpha", replace_all: true });
const beforeSearch = text(await call("wps_word_get_document_text", {}));
check("re-seeded for the search-only check", (beforeSearch.match(/alpha/g) || []).length === 2, beforeSearch.replace(/\s+/g, " ").slice(0, 80));

const r4 = await call("wps_word_find_replace", { find_text: "alpha" });
const t4 = text(r4);
check("search-only returns a real count", ok(r4) && /2\s*次/.test(t4), t4.replace(/\s+/g, " ").slice(0, 90));
const afterSearch = text(await call("wps_word_get_document_text", {}));
check("SEARCH-ONLY DID NOT DELETE THE MATCHES", (afterSearch.match(/alpha/g) || []).length === 2, afterSearch.replace(/\s+/g, " ").slice(0, 90));

// A search that finds nothing must say so rather than silently "succeed".
const r5 = await call("wps_word_find_replace", { find_text: "no-such-token-xyz" });
check("search with no match reports zero", ok(r5) && /0\s*次/.test(text(r5)), text(r5).replace(/\s+/g, " ").slice(0, 90));

// Close everything this test opened. The close actions are dialog-safe now (see
// test/close-safety.test.mjs), so a scratch run leaves no documents behind for the next one.
// Word has no close tool, so all three go through the wps_call facade.
for (const [method, appType] of [["closeWorkbook", "et"], ["closeDocument", "wps"], ["closePresentation", "wpp"]]) {
  for (let i = 0; i < 6; i++) {
    const res = await call("wps_call", { tool: "wps_execute_method", args: { method, params: { save: false }, appType } });
    if (!ok(res)) break;
  }
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "FIND/REPLACE TESTS OK (" + results.length + ")" : "FIND/REPLACE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

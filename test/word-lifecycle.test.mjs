// Verifies the fixes the first real end-to-end DSH run exposed: free-text wps_help search,
// open-workbook list formatting, the Word create/close pair, and honest formula/page reporting.
// Run: node test/word-lifecycle.test.mjs
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
function textOf(res) { try { return res.result.content[0].text; } catch { return ""; } }
function payload(res) { try { return JSON.parse(textOf(res)); } catch { return {}; } }
function ok(res) { return !!(res && res.result && !res.result.isError); }
let id = 10;
async function callTool(name, args) { return req(id++, "tools/call", { name, arguments: args }); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "wl", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });

const namesFor = async (query) => (payload(await callTool("wps_help", { query })).tools || []).map((t) => t.name);

// --- wps_help free-text search: used to test the whole query as one substring ---
const closeWord = await namesFor("close workbook 关闭");
check("multi-word query finds close_workbook", closeWord.includes("wps_excel_close_workbook"), closeWord.slice(0, 4).join(", "));
const newDoc = await namesFor("新建 文档 create new");
check("mixed-language query finds create_document", newDoc.includes("wps_word_create_document"), newDoc.slice(0, 4).join(", "));
const cjk = await namesFor("关闭工作簿");
check("space-less CJK query finds close_workbook", cjk.includes("wps_excel_close_workbook"), cjk.slice(0, 4).join(", "));
const exact = await namesFor("wps_word_create_document");
check("exact name ranks first", exact[0] === "wps_word_create_document", exact.slice(0, 3).join(", "));
const none = payload(await callTool("wps_help", { query: "zzqqxyznope" }));
check("unmatched query explains the next step", none.matched === 0 && typeof none.hint === "string", JSON.stringify(none).slice(0, 110));

// --- open-workbook list formatting ---
const wb = await callTool("wps_call", { tool: "wps_excel_create_workbook", args: {} });
check("scratch workbook created", ok(wb), textOf(wb).slice(0, 70));
const list = textOf(await callTool("wps_excel_get_open_workbooks", {}));
check("open-workbook list has no [object Object]", !list.includes("[object Object]"), list.replace(/\n/g, " | ").slice(0, 110));
check("open-workbook list names and describes the workbook", /工作簿 \(1个\)/.test(list) && list.includes("|"), list.replace(/\n/g, " | ").slice(0, 110));

// --- formula result reporting ---
await callTool("wps_call", { tool: "wps_excel_write_range", args: { range: "A1:B2", data: [[1, 2], [3, 4]] } });
const single = textOf(await callTool("wps_excel_set_formula", { range: "D1", formula: "=SUM(A1:B2)" }));
check("single-cell formula reports its value", /计算结果: 10/.test(single), single.replace(/\n/g, " | "));
const multi = textOf(await callTool("wps_excel_set_formula", { range: "D1:D2", formula: "=SUM(A1:B2)" }));
check("multi-cell formula broadcasts and reports the first cell", /计算结果（区域首格）: 10/.test(multi) && !/null/.test(multi), multi.replace(/\n/g, " | "));
check("multi-cell formula warns about the identical-formula broadcast", /同一个公式/.test(multi), multi.replace(/\n/g, " | ").slice(-90));

// --- Word create / inspect / close ---
const created = await callTool("wps_word_create_document", {});
check("wps_word_create_document succeeds", ok(created), textOf(created).slice(0, 70));
const info = textOf(await callTool("wps_word_get_active_document", {}));
check("document info has no undefined field", !/undefined/.test(info), info.replace(/\n/g, " | "));
check("document info reports paragraph and character counts", /段落数: \d+/.test(info) && /字符数: \d+/.test(info), info.replace(/\n/g, " | "));
check("unsaved document is not given a file path", info.includes("(尚未保存到磁盘)"), info.replace(/\n/g, " | "));
check("document info reports a page count", /页数: \d+/.test(info), info.replace(/\n/g, " | "));
const closedDoc = await callTool("wps_word_close_document", { save: false });
check("wps_word_close_document succeeds", ok(closedDoc) && textOf(closedDoc).includes("文档已关闭"), textOf(closedDoc).replace(/\n/g, " | ").slice(0, 90));
const openDocs = textOf(await callTool("wps_call", { tool: "wps_word_get_open_documents", args: {} }));
check("Word document list is empty again", /没有打开任何文档|0个/.test(openDocs), openDocs.replace(/\n/g, " | ").slice(0, 90));

// --- cleanup ---
await callTool("wps_call", { tool: "wps_excel_close_workbook", args: { save: false } });
child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "WORD LIFECYCLE TESTS OK (" + results.length + ")" : "WORD LIFECYCLE TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

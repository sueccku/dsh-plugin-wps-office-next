// Best-effort failures must not vanish: every swallowed error in the bridge is collected and
// attached to the action's result as "warnings" (see Add-WpsWarning / Output-Json).
// Run: node test/warnings.test.mjs
import { spawn } from "node:child_process";
// FIXES 95（W1-4）：这里验证的语义是"样式没设上、但文字照样插进去了"（部分成功）。
// 文件只断言了 success 与警告内容，**从没验证文字真的在文档里** —— 部分成功最容易退化成
// "其实什么都没做，只是警告照报"。补一条裸 COM 读正文。
import { com } from "./lib/oracle.mjs";

const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }
function payload(res) { try { return JSON.parse(res.result.content[0].text); } catch { return {}; } }
function ok(res) { return !!(res && res.result && !res.result.isError); }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "warn", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });
const viaAction = (method, params) => call("wps_call", { tool: "wps_execute_method", args: { method, params: params || {} } });
const raw = async (method, params) => payload(await viaAction(method, params));

const created = await viaAction("createDocument", {});
check("scratch document created", ok(created), "");

// R11 不变量：`success === true` ⟺ 调用方要求的事**全部生效**。
// 一个不存在的样式名会让「设置样式」这步没做到 —— 文字照样插进去了，但要求的没全生效，
// 所以这里 success 必须是 false，并且短欠条目要单独报出来（这是 shortfall，不是 warning）。
const seeded = await raw("insertText", { text: "warning probe", position: "start", style: "NoSuchStyleName" });
check("a request that did not fully apply is NOT success", seeded.success === false, JSON.stringify(seeded).slice(0, 100));
check("it is marked partial, not a plain failure", seeded.partial === true, JSON.stringify(seeded).slice(0, 100));
// FIXES 93（W1-2）：以前只断言"至少一条"和"长度 > 8" —— 这两句证明不了内容对不对。
check("exactly one shortfall is reported for the one unfulfilled step", Array.isArray(seeded.shortfalls) && seeded.shortfalls.length === 1, JSON.stringify(seeded.shortfalls || []).slice(0, 110));
check("the shortfall names the style that actually failed", typeof seeded.shortfalls?.[0] === "string" && seeded.shortfalls[0].includes("NoSuchStyleName"), String(seeded.shortfalls?.[0]).slice(0, 120));
check("shortfalls are also inside data for pass-through tools", Array.isArray(seeded.data?.shortfalls), JSON.stringify(seeded.data || {}).slice(0, 90));

// 工具层必须把它变成**看得见的失败**（MCP 的 isError），并且正文里保留已经生效的部分 ——
// 这正是不变量的意义：只看 success 的调用方也不会误以为一切照做了。
const toolLayer = await call("wps_word_insert_text", { text: "tool layer probe", position: "end", style: "NoSuchStyleName" });
check("the tool layer reports it as an error, not a silent success", !!(toolLayer.result && toolLayer.result.isError), String(toolLayer.result?.content?.[0]?.text || "").slice(0, 100));
// R11 第一方工具的**已知边界**（记在 FIXES 103）：handler 走失败分支时只打印 error 文案，
// 所以「已生效部分」的结构化细节在**直通路径**（桥层 data）里完整，第一方工具的正文里只有短欠清单。
// 这里断言当前契约下确实成立的两件事：报成 error、且短欠条目写在文案里。
check("the tool layer names the unfulfilled step in its error", /未生效|不一致|partial/.test(String(toolLayer.result?.content?.[0]?.text || "")), String(toolLayer.result?.content?.[0]?.text || "").slice(0, 130));

// ---- 独立 oracle：部分成功的**实际效果**必须落地 ----
const docText = String(com("$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application'); [string]$w.ActiveDocument.Content.Text")).trim();
check("the insert really happened despite the failed style", docText.includes("warning probe"), JSON.stringify(docText).slice(0, 130));

// Warnings are per action: a clean call must not inherit the previous one.
const clean = await raw("ping", {});
check("a clean action reports no warnings", clean.warnings === undefined, JSON.stringify(clean).slice(0, 80));
check("a clean action reports no shortfalls either", clean.shortfalls === undefined && clean.partial === undefined, JSON.stringify(clean).slice(0, 80));

// And a first-class tool whose optional step succeeds reports nothing either.
const text = await raw("getDocumentText", {});
check("a successful action after a warned one stays clean", text.warnings === undefined, JSON.stringify(text).slice(0, 90));

for (let i = 0; i < 3; i++) {
  const res = await call("wps_call", { tool: "wps_execute_method", args: { method: "closeDocument", params: { save: false } } });
  if (!ok(res)) break;
}

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "WARNING TESTS OK (" + results.length + ")" : "WARNING TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

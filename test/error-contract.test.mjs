// S6: the failure and timeout contract, asserted. Companion to docs/error-contract.md.
// The timeout half is driven by test/watchdog.test.mjs (stub host); this file asserts the batch
// contract and the facade errors against a real server, plus the documented phrases in the sources.
// Run: node test/error-contract.test.mjs
import { spawn } from "node:child_process";
import { readFileSync } from "node:fs";

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }

// ---- static: the contract text must stay where the doc says it is -------------------------
const comHost = readFileSync("mcp/src/client/com-host.ts", "utf8");
check("timeout text keeps the \"state unknown\" wording", comHost.includes("状态未知"), "");
check("timeout tiers are the documented env-overridable ones", comHost.includes("WPS_OFFICE_TIMEOUT_MS") && comHost.includes("WPS_OFFICE_LONG_TIMEOUT_MS") && comHost.includes("WPS_OFFICE_SUSPECT_TIMEOUT_MS"), "");
check("a timeout never force-closes WPS", comHost.includes("不会自动关闭 WPS"), "");
const doc = readFileSync("docs/error-contract.md", "utf8");
check("the contract doc records the batch limit", doc.includes("单次批量最多 50 项"), "");
check("the contract doc records the batch continue-on-failure rule", doc.includes("继续执行后面的项"), "");

// ---- behavioral: batch and facade --------------------------------------------------------
const child = spawn(process.execPath, ["mcp/dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
let buf = "";
const pending = new Map();
function send(o) { child.stdin.write(JSON.stringify(o) + "\n"); }
function req(id, method, params) { return new Promise((r) => { pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); }); }
child.stdout.on("data", (d) => { buf += d.toString(); let i; while ((i = buf.indexOf("\n")) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
child.stderr.on("data", () => {});
function text(res) { return res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}); }
function isErr(res) { return !!(res && res.result && res.result.isError); }
function payload(res) { try { return JSON.parse(text(res)); } catch { return {}; } }

await req(1, "initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "errctr", version: "1" } });
send({ jsonrpc: "2.0", method: "notifications/initialized" });
let id = 10;
const call = (name, args) => req(id++, "tools/call", { name, arguments: args });

const empty = await call("wps_batch", { calls: [] });
check("empty batch is rejected", isErr(empty) && text(empty).includes("calls 不能为空"), text(empty).slice(0, 70));

const tooMany = await call("wps_batch", { calls: Array.from({ length: 51 }, () => ({ tool: "wps_common_ping" })) });
check("a batch over 50 is rejected as a whole", isErr(tooMany) && text(tooMany).includes("50"), text(tooMany).slice(0, 70));

const mixed = await call("wps_batch", { calls: [
  { tool: "wps_common_ping" },
  { tool: "wps_definitely_not_a_tool" },
  { tool: "wps_common_ping" },
] });
const m = payload(mixed);
check("a partial failure does not fail the whole batch", !isErr(mixed) && m.count === 3, JSON.stringify(m).slice(0, 120));
check("the first item succeeded", !!(m.results && m.results[0] && m.results[0].success === true), "");
check("the unknown tool is reported as its own failed item", !!(m.results && m.results[1] && m.results[1].success === false && String(m.results[1].error || "").includes("无效")), JSON.stringify(m.results && m.results[1]));
check("execution continued after the failed item", !!(m.results && m.results[2] && m.results[2].success === true), "");
check("every item result is truncated to 2000 chars", !!(m.results && m.results.every((r) => !r.result || r.result.length <= 2000)), "");

const unknown = await call("wps_call", { tool: "wps_definitely_not_a_tool", args: {} });
check("wps_call rejects an unknown tool", isErr(unknown) && text(unknown).includes("未知工具"), text(unknown).slice(0, 70));
const recursion = await call("wps_call", { tool: "wps_call", args: { tool: "wps_common_ping", args: {} } });
check("wps_call refuses facade recursion", isErr(recursion) && text(recursion).includes("门面"), text(recursion).slice(0, 70));

child.kill();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "ERROR CONTRACT TESTS OK (" + results.length + ")" : "ERROR CONTRACT TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

// Regression test for the resident COM host: recovery, serialization, protocol framing.
// Run: node test/com-host.test.mjs
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { comHost, Utf8LineReader } = require("../mcp/dist/client/com-host.js");

const results = [];
function check(name, ok, detail) {
  results.push({ name, ok });
  console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : ""));
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---- 按字节收帧（改进指南 §9 P1：正式客户端不能用 chunk.toString()）---------------------------
// 一个 chunk 边界可能落在多字节中文字符中间；toString() 会把半个字符解成 U+FFFD，整帧 JSON 里
// 的中文就坏了 —— 而模型不该为了绕开编码缺陷去回避中文。这里把同一帧按各种切法喂进去。
const MESSAGE = "WPS MCP Bridge 已连接 — 中文 test";
const FRAME = JSON.stringify({ id: 7, ok: true, result: { success: true, data: { message: MESSAGE } } }) + "\n";
const bytes = Buffer.from(FRAME, "utf8");
check("the probe frame really contains multi-byte characters", bytes.length > FRAME.length, "bytes=" + bytes.length + " chars=" + FRAME.length);
for (const split of [1, 2, 3, 5, 11, 17]) {
  const reader = new Utf8LineReader();
  const lines = [...reader.push(bytes.subarray(0, split)), ...reader.push(bytes.subarray(split))];
  let message = null;
  try { message = JSON.parse(lines[0]).result.data.message; } catch { message = null; }
  check("a Chinese frame survives a byte-boundary split at " + split, lines.length === 1 && message === MESSAGE, lines.length + " line(s), message=" + String(message));
}
{
  const reader = new Utf8LineReader();
  const lines = [];
  for (let i = 0; i < bytes.length; i++) lines.push(...reader.push(bytes.subarray(i, i + 1)));
  let message = null;
  try { message = JSON.parse(lines[0]).result.data.message; } catch { message = null; }
  check("byte-by-byte framing reassembles the frame", lines.length === 1 && message === MESSAGE, lines.length + " line(s)");
}
{
  // 半帧不能被当成一帧；补完才吐出来，且 reset() 必须清掉残留。
  const reader = new Utf8LineReader();
  const partial = reader.push(bytes.subarray(0, 12));
  check("an incomplete frame is not emitted", partial.length === 0, JSON.stringify(partial));
  reader.reset();
  const afterReset = reader.push(bytes.subarray(12));
  // reset 之后，上一次留下的半帧（`{"id":7,"ok`）**不能**被拼到新进程的首帧上 ——
  // 所以这里吐出来的是"后半段"，而不是一条以 `{` 开头的完整帧。
  check("reset drops the half frame (no cross-process mixing)", afterReset.length === 1 && !afterReset[0].startsWith("{"), JSON.stringify(afterReset).slice(0, 80));
}

const first = await comHost.invoke("ping", {});
check("cold ping succeeds", first.success === true, JSON.stringify(first).slice(0, 80));

// 端到端：桥回的中文（"WPS MCP Bridge 已连接"）必须原样到客户端，不能被编码链改坏。
const wire = await comHost.invoke("wireCheck", {});
check("Chinese from the bridge reaches the client intact", wire.success === true && String(wire.data?.message || "").includes("已连接"), JSON.stringify(wire).slice(0, 120));

const warmStart = Date.now();
await comHost.invoke("ping", {});
const warmMs = Date.now() - warmStart;
// 100ms 在负载高的机器上会假红；这里只要求"热路径没有退化到秒级"。
check("warm ping stays in the sub-second range", warmMs < 500, "warm=" + warmMs + "ms");

// Concurrent calls must be serialized by the client and all resolve correctly.
const concurrent = await Promise.all([
  comHost.invoke("ping", {}),
  comHost.invoke("wireCheck", {}),
  comHost.invoke("ping", {}),
  comHost.invoke("wireCheck", {}),
  comHost.invoke("ping", {})
]);
check("5 concurrent calls all succeed", concurrent.every((r) => r.success === true), "ok=" + concurrent.filter((r) => r.success).length + "/5");

// Kill the host process and confirm the next call transparently respawns it.
const child = comHost.child;
check("host process is attached", !!child, child ? "pid=" + child.pid : "no child");
try { child.kill(); } catch { /* already gone */ }
await sleep(400);
let recovered = null;
try {
  recovered = await comHost.invoke("ping", {});
} catch (error) {
  recovered = { success: false, error: String(error.message) };
}
check("host recovers after being killed", recovered.success === true, JSON.stringify(recovered).slice(0, 100));

// An unknown action must come back as a structured failure, never as a silent success.
const unknown = await comHost.invoke("__definitely_not_an_action", {});
check("unknown action fails loudly", unknown.success === false && typeof unknown.error === "string", JSON.stringify(unknown).slice(0, 100));

await comHost.stop();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "COM HOST TESTS OK (" + results.length + ")" : "COM HOST TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

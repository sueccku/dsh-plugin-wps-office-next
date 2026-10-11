// 测试文件共用的最小 harness：JSON-RPC stdio 客户端 + check()/收尾。
//
// 为什么要有它：本仓库 38 个测试文件各自抄了一遍同一段 stdio/JSON-RPC 代码，47 个各自抄了一遍
// check()/results/process.exit。抄件一多，行为就开始分叉（超时、env 钉法、PASS 文案）。新测试请用它，
// 老文件可以按需增量迁移。
//
// 与 scripts/run-tests.ps1 对齐的两件事（漏了会得到"假失败"）：
//   1. spawn 前把 WPS_OFFICE_HOST_SCRIPT 钉到**仓库内**路径 —— 从「已装本插件」的 DSH 会话里起终端时，
//      环境里的那份指向 profile 里的旧副本，测的就不是仓库代码了。
//   2. 每个调用都带超时：挂住的 WPS 不能把整个文件拖死。
import { spawn } from "node:child_process";
import { resolve } from "node:path";

/** 起一个 MCP server 客户端。返回的 call() 已带默认 30s 超时；stop() 结束子进程。 */
export function startServer({ entry = "mcp/dist/index.js", clientName = "harness", env = {} } = {}) {
  const child = spawn(process.execPath, [entry], {
    stdio: ["pipe", "pipe", "pipe"],
    windowsHide: true,
    env: { ...process.env, WPS_OFFICE_HOST_SCRIPT: resolve("host/wps-com-host.ps1"), ...env },
  });
  let nextId = 1;
  const pending = new Map();
  // **按字节收帧**：不要写成 `buf += d.toString()`。chunk 边界切开多字节 UTF-8 字符时，那种写法会把
  // 中文变成 U+FFFD（2026-10-11 实测：`指`→2 个替换字符、`明`→3 个），于是"断言里含中文"的用例
  // 会偶发假红，而错误信息本身也被改坏、难以定位。定位依据：MCP server 的 stdout 原始字节里
  // `efbfbd` 出现 0 次 —— 数据是干净的，坏的是读法。
  let raw = Buffer.alloc(0);
  child.stdout.on("data", (d) => {
    raw = Buffer.concat([raw, d]);
    let i;
    while ((i = raw.indexOf(0x0a)) >= 0) {
      const line = raw.subarray(0, i).toString("utf8").trim();
      raw = raw.subarray(i + 1);
      if (!line) continue;
      let m;
      try { m = JSON.parse(line); } catch { continue; }
      if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); }
    }
  });
  child.stderr.on("data", () => {});
  const send = (o) => child.stdin.write(JSON.stringify(o) + "\n");
  const req = (method, params) => new Promise((r) => { const id = nextId++; pending.set(id, r); send({ jsonrpc: "2.0", id, method, params }); });
  const ready = (async () => {
    await req("initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: clientName, version: "1" } });
    send({ jsonrpc: "2.0", method: "notifications/initialized" });
  })();
  const call = async (name, args, ms = 30000) => {
    await ready;
    const res = await Promise.race([req("tools/call", { name, arguments: args || {} }), new Promise((r) => setTimeout(() => r({ __timeout: true }), ms))]);
    if (res && res.__timeout) console.log("      TIMEOUT " + name);
    return res;
  };
  const text = (res) => (res && res.result && res.result.content ? String(res.result.content[0].text) : JSON.stringify((res && res.error) || {}));
  const ok = (res) => !!(res && res.result && !res.result.isError);
  const payload = (res) => { try { return JSON.parse(text(res)); } catch { return {}; } };
  return { child, call, req, text, ok, payload, stop: () => child.kill() };
}

/** 统一的结果收集器：check() 打印 PASS/FAIL，finish() 打印小结并按失败数决定退出码。 */
export function createReport(title) {
  const results = [];
  const check = (name, okFlag, detail) => {
    results.push({ name, ok: !!okFlag });
    console.log((okFlag ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : ""));
  };
  const finish = () => {
    const failed = results.filter((r) => !r.ok).length;
    console.log(failed === 0 ? title + " OK (" + results.length + ")" : title + " FAILED (" + failed + "/" + results.length + ")");
    process.exit(failed === 0 ? 0 : 1);
  };
  return { check, finish, results };
}

// 对一个**装好的**插件副本做真实 MCP 握手，并调用一次 wps_status。
//
// 与 scripts/verify-package.mjs 的分工：
//   verify-package.mjs —— 发布前，对 `npm pack` 的产物做（自己装临时目录、不需要 WPS 也能验证握手）。
//   本脚本            —— 发布后 / 排错时，对任意一个已安装的副本做，可指定路径，会真的调 wps_status（需要 WPS）。
//
// Usage: node scripts/probe-installed.mjs [插件根目录]
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

const arg = process.argv[2];
const pluginRoot = arg || join(homedir(), '.dsh', 'profiles', 'desktop', 'node_modules', 'dsh-plugin-wps-office-next');
if (!existsSync(pluginRoot)) { console.log("plugin root not found: " + pluginRoot); process.exit(1); }
const entry = join(pluginRoot, 'mcp', 'dist', 'index.js');
const hostScript = join(pluginRoot, 'host', 'wps-com-host.ps1');
console.log('plugin root: ' + pluginRoot);

const child = spawn(process.execPath, [entry], {
  stdio: ["pipe", "pipe", "pipe"],
  env: { ...process.env, WPS_OFFICE_HOST_SCRIPT: hostScript, WPS_OFFICE_TOOLSET: "standard" },
});
const stderr = [];
child.stderr.on("data", (d) => { for (const l of String(d).split(/\r?\n/)) if (l.trim()) stderr.push(l.trim()); });
let buf = "";
const pending = new Map();
child.stdout.on("data", (d) => {
  buf += d.toString();
  let i;
  while ((i = buf.indexOf("\n")) >= 0) {
    const line = buf.slice(0, i).trim();
    buf = buf.slice(i + 1);
    if (!line) continue;
    try { const m = JSON.parse(line); if (m.id !== undefined && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } }
    catch { stderr.push("NON-JSON STDOUT: " + line.slice(0, 160)); }
  }
});
let seq = 0;
const send = (method, params) => new Promise((resolve, reject) => {
  const id = ++seq;
  pending.set(id, resolve);
  child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n");
  setTimeout(() => { if (pending.has(id)) { pending.delete(id); reject(new Error("timeout waiting for " + method)); } }, 120000);
});

let exitCode = 0;
try {
  const init = await send("initialize", { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "probe-installed", version: "1.0.0" } });
  console.log("INITIALIZE OK   serverInfo = " + JSON.stringify(init.result && init.result.serverInfo));
  child.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized", params: {} }) + "\n");
  const tools = await send("tools/list", {});
  const names = ((tools.result && tools.result.tools) || []).map((t) => t.name);
  console.log("TOOLS ADVERTISED = " + names.length);
  const status = names.find((n) => /status/i.test(n)) || names[0];
  const call = await send("tools/call", { name: status, arguments: {} });
  const text = ((call.result && call.result.content) || []).map((c) => c.text).join("\n");
  console.log("tools/call " + status + "  isError = " + (call.result && call.result.isError));
  console.log(text.slice(0, 1200));
} catch (e) {
  console.log("HANDSHAKE FAILED: " + e.message);
  exitCode = 1;
}
if (stderr.length) { console.log("--- server stderr (last 25) ---"); console.log(stderr.slice(-25).join("\n")); }
try { child.kill(); } catch { /* already gone */ }
setTimeout(() => process.exit(exitCode), 300);

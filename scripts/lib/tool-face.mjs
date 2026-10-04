// 从**真实 MCP server** 量三档工具面的载荷：tools/list 返回的 JSON 字节数。
// 为什么不用 spec 推：spec 是「应该长什么样」，tools/list 是「模型实际收到什么」——
// 预算门禁与 README 的数字都必须以后者为准（FIXES 91）。
import { spawn } from 'node:child_process';
import path from 'node:path';

function measure(mode, entry, env) {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, [entry], {
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
      env: { ...process.env, ...env, WPS_OFFICE_TOOLSET: mode },
    });
    let buf = '';
    const pending = new Map();
    const send = (o) => child.stdin.write(JSON.stringify(o) + '\n');
    const req = (id, method, params) => new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); });
    let done = false;
    const finish = (v) => { if (!done) { done = true; try { child.kill(); } catch {} resolve(v); } };
    const timer = setTimeout(() => finish({ mode, error: 'timeout' }), 30000);
    child.stdout.on('data', (d) => {
      buf += d.toString();
      let i;
      while ((i = buf.indexOf('\n')) >= 0) {
        const line = buf.slice(0, i).trim();
        buf = buf.slice(i + 1);
        if (!line) continue;
        let j; try { j = JSON.parse(line); } catch { continue; }
        if (j.id && pending.has(j.id)) { pending.get(j.id)(j); pending.delete(j.id); }
      }
    });
    child.stderr.on('data', () => {});
    child.on('error', (e) => { clearTimeout(timer); finish({ mode, error: String(e.message) }); });
    (async () => {
      await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'gen-numbers', version: '1' } });
      send({ jsonrpc: '2.0', method: 'notifications/initialized' });
      const r = await req(2, 'tools/list', {});
      clearTimeout(timer);
      if (!r || !r.result || !Array.isArray(r.result.tools)) return finish({ mode, error: 'no tools/list' });
      const payload = JSON.stringify(r.result.tools);
      const bytes = Buffer.byteLength(payload, 'utf8');
      finish({ mode, tools: r.result.tools.length, bytes, approxTokens: Math.round(bytes / 3.5) });
    })().catch((e) => { clearTimeout(timer); finish({ mode, error: String(e && e.message) }); });
  });
}

export async function measureTiers() {
  const ROOT = process.cwd();
  const entry = path.join(ROOT, 'mcp', 'dist', 'index.js');
  const env = {
    WPS_OFFICE_MCP_ENTRY: entry,
    WPS_OFFICE_HOST_SCRIPT: path.join(ROOT, 'host', 'wps-com-host.ps1'),
  };
  const out = {};
  for (const mode of ['minimal', 'standard', 'full']) out[mode] = await measure(mode, entry, env);
  return out;
}

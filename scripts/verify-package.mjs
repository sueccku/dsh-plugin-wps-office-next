// 打包产物的运行时冒烟测试：真的把打包出来的 MCP server 拉起来，做一次 JSON-RPC 握手。
//
// 为什么需要它（FIXES 84）：0.6.0 把 `files` 从整个 `mcp` 目录改成 `mcp/dist` + `mcp/scripts`，漏掉了
// `mcp/package.json`。那份文件是 `mcp/` 的模块边界声明，没有它，`mcp/dist/*.js`（CommonJS 编译产物）
// 会继承根 `package.json` 的 `"type": "module"`，一启动就抛 `exports is not defined in ES module scope`。
// 当时的验证只看「文件在不在」，所以全绿 —— 这个脚本补的就是那一关：真的把 server 拉起来握手。
//
// 不需要 WPS：握手与 tools/list 只用 MCP 协议本身。
// Usage: node scripts/verify-package.mjs [--plugin-root <解开的包目录>]
import { spawn, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = dirname(dirname(fileURLToPath(import.meta.url)));
const argv = process.argv.slice(2);
const rootArg = argv.indexOf("--plugin-root");
const KEEP = argv.includes('--keep');

let checks = 0;
let failed = 0;
function check(name, ok, detail) {
  checks++;
  if (!ok) failed++;
  console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : ''));
}

/** Resolve (or build) the plugin root that should be smoke-tested. */
function resolvePluginRoot() {
  if (rootArg >= 0 && argv[rootArg + 1]) return argv[rootArg + 1];
  // shell: true because npm on Windows is a .cmd shim; --no-deprecation keeps DEP0190 out of the output.
  const npmOpts = { cwd: ROOT, encoding: 'utf8', shell: true, env: { ...process.env, NODE_OPTIONS: '--no-deprecation' } };
  const packed = spawnSync('npm', ['pack', '--dry-run', '--json'], npmOpts);
  if (packed.status !== 0) { console.log('FAIL npm pack --dry-run'); console.log(packed.stderr); process.exit(1); }
  const manifest = JSON.parse(packed.stdout)[0];
  check("packed manifest lists mcp/package.json (the module boundary)", manifest.files.some((f) => f.path === "mcp/package.json"),
    "entries=" + manifest.entryCount);
  const staging = mkdtempSync(join(tmpdir(), "wps-verify-package-"));
  const tarball = spawnSync("npm", ["pack", "--pack-destination", staging], npmOpts);
  if (tarball.status !== 0) { console.log("FAIL npm pack"); console.log(tarball.stderr); process.exit(1); }
  const tgz = readdirSync(staging).find((f) => f.endsWith(".tgz"));
  // Install the tarball the way a consumer does: the pack output has no node_modules, and the
  // runtime dependencies come from the root package.json, so an unpacked tree alone cannot boot.
  const consumer = join(staging, "consumer");
  mkdirSync(consumer, { recursive: true });
  writeFileSync(join(consumer, "package.json"), JSON.stringify({ name: "verify-package-consumer", version: "1.0.0", private: true }));
  const install = spawnSync("npm", ["install", "--no-audit", "--no-fund", "--ignore-scripts", join(staging, tgz)],
    { ...npmOpts, cwd: consumer, timeout: 300000 });
  check("npm install of the packed tarball succeeds", install.status === 0, install.status === 0 ? "" : String(install.stderr).slice(-200));
  const installed = join(consumer, "node_modules", "dsh-plugin-wps-office-next");
  if (KEEP) console.log("kept staging dir: " + staging);
  return { root: installed, cleanup: () => { if (!KEEP) rmSync(staging, { recursive: true, force: true }); } };
}

/** Drive one JSON-RPC session against the pack output. Never needs WPS. */
function handshake(pluginRoot) {
  return new Promise((resolve) => {
    const entry = join(pluginRoot, "mcp", "dist", "index.js");
    const hostScript = join(pluginRoot, "host", "wps-com-host.ps1");
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
      let idx;
      while ((idx = buf.indexOf("\n")) >= 0) {
        const line = buf.slice(0, idx).trim();
        buf = buf.slice(idx + 1);
        if (!line) continue;
        try { const msg = JSON.parse(line); if (msg.id !== undefined && pending.has(msg.id)) { pending.get(msg.id)(msg); pending.delete(msg.id); } }
        catch { stderr.push("NON-JSON STDOUT: " + line.slice(0, 160)); }
      }
    });
    let id = 0;
    const send = (method, params) => new Promise((res, rej) => {
      const n = ++id;
      pending.set(n, res);
      child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: n, method, params }) + "\n");
      setTimeout(() => { if (pending.has(n)) { pending.delete(n); rej(new Error("timeout waiting for " + method)); } }, 60000);
    });
    const done = (value) => { try { child.kill(); } catch { /* already gone */ } resolve(value); };
    setTimeout(() => done({ timeout: true, stderr }), 90000);
    (async () => {
      try {
        const init = await send("initialize", { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "verify-package", version: "1.0.0" } });
        child.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized", params: {} }) + "\n");
        const tools = await send("tools/list", {});
        const names = ((tools.result && tools.result.tools) || []).map((t) => t.name);
        done({ ok: true, serverInfo: init.result && init.result.serverInfo, toolCount: names.length, hasStatus: names.includes("wps_status"), stderr });
      } catch (e) { done({ ok: false, error: e.message, stderr }); }
    })();
  });
}

const resolved = resolvePluginRoot();
const pluginRoot = typeof resolved === "string" ? resolved : resolved.root;
const cleanup = typeof resolved === "string" ? () => {} : resolved.cleanup;
console.log("plugin root: " + pluginRoot);

for (const rel of ["mcp/package.json", "mcp/dist/index.js", "host/wps-com-host.ps1", "plugin.js", "cordis.patch.yml"]) {
  check("pack output carries " + rel, existsSync(join(pluginRoot, rel)));
}
if (existsSync(join(pluginRoot, "mcp", "package.json"))) {
  const type = JSON.parse(readFileSync(join(pluginRoot, "mcp", "package.json"), "utf8")).type;
  check("mcp/package.json declares a CommonJS boundary", type === undefined || type === "commonjs", "type=" + String(type));
}

const result = await handshake(pluginRoot);
check("MCP server answers initialize", Boolean(result.ok), result.ok ? "" : String(result.error || "timeout"));
if (result.ok) {
  check("tools/list advertises the standard toolset", result.toolCount === 69, "tools=" + result.toolCount);
  check("wps_status is advertised", result.hasStatus === true);
}
if (!result.ok || result.stderr.length) {
  console.log("--- server stderr ---");
  console.log(result.stderr.slice(-25).join("\n"));
}
cleanup();
console.log(failed === 0 ? "PACKAGE RUNTIME OK (" + checks + ")" : "PACKAGE RUNTIME FAILED (" + failed + "/" + checks + ")");
process.exit(failed === 0 ? 0 : 1);


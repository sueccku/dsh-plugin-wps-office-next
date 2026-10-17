// 找到 dsh 的可执行入口，并给出**不经过 shell**的启动方式。
//
// 为什么不能直接 spawn('dsh')：Windows 上 dsh 是 .cmd shim，Node 的 spawn 不带 shell 时
// 解析不到它（ENOENT），而带 shell 又要处理引号转义。这里按 scripts/e2e.mjs 的办法
// 找到真正的 lib/bin.js，用当前 node 直接跑 —— 路径里有没有空格都不需要担心。
import { existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { spawnSync } from 'node:child_process';

const REL = join('node_modules', '@deepseek-ai', 'dsh', 'lib', 'bin.js');

/** 返回 { label, command, args, env }；找不到返回 null。 */
export function resolveDshLauncher(explicit) {
  const candidates = [];
  if (explicit) candidates.push(explicit);
  if (process.env.DSH_BIN) candidates.push(process.env.DSH_BIN);
  const exeDir = dirname(process.execPath);
  candidates.push(join(exeDir, REL));
  candidates.push(join(exeDir, '..', 'lib', REL));
  if (process.env.APPDATA) candidates.push(join(process.env.APPDATA, 'npm', REL));

  // PATH 上的 dsh.cmd 旁边的 node_modules 也是候选（run 目录布局）
  const where = spawnSync('where.exe', ['dsh'], { encoding: 'utf8', windowsHide: true });
  if (!where.error && where.status === 0) {
    for (const line of String(where.stdout).split(/\r?\n/)) {
      const shim = line.trim();
      if (!shim) continue;
      const shimDir = dirname(shim);
      candidates.push(join(shimDir, REL));
      candidates.push(join(shimDir, '..', REL));
    }
  }

  const bin = candidates.find((candidate) => candidate && existsSync(candidate));
  if (bin) return { label: bin, command: process.execPath, args: [bin], env: {} };

  // 打包成桌面应用时走 Electron：app.asar 里的 cli.js 只能由 exe 自己加载
  const roots = [];
  if (process.env.LOCALAPPDATA) roots.push(join(process.env.LOCALAPPDATA, 'Programs', 'DeepSeek Harness'));
  if (process.env.ProgramFiles) roots.push(join(process.env.ProgramFiles, 'DeepSeek Harness'));
  for (const root of roots) {
    const exe = join(root, 'DeepSeek Harness.exe');
    const asar = join(root, 'resources', 'app.asar');
    const cli = join(asar, 'dsh', 'node_modules', '@deepseek-ai', 'dsh-desktop-host', 'lib', 'cli.js');
    if (existsSync(exe) && existsSync(asar)) {
      return { label: exe + ' -> ' + cli, command: exe, args: ['--expose-internals', cli], env: { ELECTRON_RUN_AS_NODE: '1' } };
    }
  }
  return null;
}

// 找得到一个能读 Office 文件的 Python。
//
// 为什么需要它：夹具生成与产物检查都要 openpyxl / python-docx / python-pptx。这台机器上 PATH 里的
// python 未必带这些库，而 DSH 自带运行时里带（版本见 docs/skill-refactor-report.md 的实测记录）。
// 这里**不写死用户路径**：从 $DSH_HOME/dsh-runtimes/* 推，找不到再退回 PATH。
import { existsSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';

const HELP =
  '找不到带 openpyxl / python-docx / python-pptx 的 Python。\n' +
  '  * 显式指定：--python <python.exe> 或环境变量 SKILLEVAL_PYTHON\n' +
  '  * DSH 自带运行时：<DSH_HOME>/dsh-runtimes/<runtime>/dependencies/python/python.exe';

function probe(exe) {
  const run = spawnSync(exe, ['-c', 'import openpyxl, docx, pptx; print("ok")'], {
    encoding: 'utf8',
    windowsHide: true,
    env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
  });
  return run.status === 0 && String(run.stdout || '').includes('ok');
}

/** 按优先级返回一个可用的 Python 可执行文件；找不到就抛错（带怎么修）。 */
export function resolvePython(explicit) {
  const candidates = [];
  if (explicit) candidates.push(explicit);
  if (process.env.SKILLEVAL_PYTHON) candidates.push(process.env.SKILLEVAL_PYTHON);
  const home = process.env.DSH_HOME;
  if (home) {
    const runtimes = join(home, 'dsh-runtimes');
    try {
      for (const name of readdirSync(runtimes)) candidates.push(join(runtimes, name, 'dependencies', 'python', 'python.exe'));
    } catch {
      // 没有自带运行时目录就继续用 PATH
    }
  }
  candidates.push('python', 'py');

  const tried = [];
  for (const exe of candidates) {
    if (exe.includes('\\') || exe.includes('/')) {
      if (!existsSync(exe)) continue;
    }
    tried.push(exe);
    if (probe(exe)) return exe;
  }
  throw new Error(HELP + '\n已尝试：' + (tried.length ? tried.join(', ') : '(无)'));
}

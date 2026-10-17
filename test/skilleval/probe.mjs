// 跑一次 headless 会话并把轨迹落盘（自由提示词版，用于机制验证与 D1 的指标采集）。
//
// 用法：
//   node test/skilleval/probe.mjs --prompt <file> --dir <工作目录> --profile <name> [--timeout 420] [--label 名字]
//   node test/skilleval/probe.mjs --task "…"         （内联任务，等价于 --prompt）
//
// 它做四件事：
//   1) 用指定 profile 跑一次 `dsh --profile <name> <task>`（cwd = --dir），捕获 stdout/stderr/退出码/耗时；
//   2) 从 $DSH_HOME/sessions 找这次运行的会话日志（按"第一条用户消息 == 任务"认，不按时间猜）；
//   3) 解 zstd 帧还原轨迹，写出 transcript.txt / metrics.json；
//   4) 打印一行摘要，方便脚本化使用。
//
// ⚠️ 跑模型会话要花真钱。本脚本**不检查**你是否获得同意：调用方自己负责（本仓库 AGENTS.md §1）。
// 需要环境变量 DEEPSEEK_API_KEY（headless profile 从环境读，web 端存在 credentials 服务里）。
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { findSession, summarize } from './session.mjs';
import { resolveDshLauncher } from './dsh-bin.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));

function opt(argv, flag) {
  const i = argv.indexOf(flag);
  return i >= 0 ? argv[i + 1] : undefined;
}

const argv = process.argv.slice(2);
const profile = opt(argv, '--profile');
const workdir = resolve(opt(argv, '--dir') || process.cwd());
const label = opt(argv, '--label') || 'probe';
// 单次会话默认 60s：正常一次 10–60s；超时就让进程整棵树退出并如实报错，不无限等。
const timeoutMs = Number(opt(argv, '--timeout') || 60) * 1000;
const promptFile = opt(argv, '--prompt');
const inlineTask = opt(argv, '--task');

if (!profile) {
  console.error('用法: node test/skilleval/probe.mjs --prompt <file>|--task "<task>" --dir <dir> --profile <name>');
  process.exit(2);
}
if (!promptFile && !inlineTask) {
  console.error('必须给 --prompt <file> 或 --task "<task>"');
  process.exit(2);
}
const task = inlineTask || readFileSync(promptFile, 'utf8').trim();
if (!task) {
  console.error('任务为空');
  process.exit(2);
}
if (!process.env.DEEPSEEK_API_KEY) {
  console.error('缺少 DEEPSEEK_API_KEY：headless profile 从环境变量读模型凭据。');
  console.error('  web/desktop 端的凭据存在 $DSH_HOME/.credentials.yaml 的 refs.DEEPSEEK_API_KEY；');
  console.error('  跑一次：$env:DEEPSEEK_API_KEY = <该值>');
  process.exit(2);
}

mkdirSync(workdir, { recursive: true });
const outDir = join(workdir, 'probe-' + label);
mkdirSync(outDir, { recursive: true });
writeFileSync(join(outDir, 'prompt.txt'), task + '\n', 'utf8');

// 与 scripts/e2e.mjs 同一个坑：从装了插件的会话里起子进程时，继承来的
// WPS_OFFICE_MCP_ENTRY / WPS_OFFICE_HOST_SCRIPT 会指向**另一份**插件副本，让本次运行的
// MCP 工具面整个变空。这里必须剔掉，让 plugin.js 自己算路径。
const env = { ...process.env };
delete env.WPS_OFFICE_MCP_ENTRY;
delete env.WPS_OFFICE_HOST_SCRIPT;
delete env.WPS_OFFICE_MCP_NODE;

const startedAt = Date.now();
// 不能直接 spawn('dsh')：Windows 上它是 .cmd shim，Node 不带 shell 时解析不到（ENOENT）。
const launcher = resolveDshLauncher(opt(argv, '--dsh-bin'));
if (!launcher) {
  console.error('找不到 dsh 启动器（可用 --dsh-bin 或 DSH_BIN 指定 lib/bin.js）');
  process.exit(2);
}
const run = spawnSync(launcher.command, launcher.args.concat(['--profile', profile, task]), {
  cwd: workdir,
  encoding: 'utf8',
  windowsHide: true,
  timeout: timeoutMs,
  env: { ...env, ...launcher.env },
});
const elapsedMs = Date.now() - startedAt;
const stdout = String(run.stdout || '');
const stderr = String(run.stderr || '');
if (run.error && run.pid) {
  // 超时只结束直接子进程，孙进程（MCP server / COM 宿主）可能活下来占着 WPS 租约 —— 整棵树收掉。
  spawnSync('taskkill', ['/PID', String(run.pid), '/T', '/F'], { encoding: 'utf8', windowsHide: true });
}

writeFileSync(join(outDir, 'stdout.txt'), stdout, 'utf8');
writeFileSync(join(outDir, 'stderr.txt'), stderr, 'utf8');

const dshHome = process.env.DSH_HOME || join(process.env.USERPROFILE || '', '.dsh');
const session = findSession({ dshHome, task, sinceMs: startedAt });
let metrics = null;
let sessionFile = null;
if (session) {
  sessionFile = session.file;
  metrics = summarize(session.records);
  const lines = [];
  for (const call of metrics.toolCalls) lines.push('CALL   ' + call.name + ' ' + String(call.args).replace(/\s+/g, ' ').slice(0, 300));
  for (const result of metrics.toolResults) lines.push('RESULT ' + result.name + ' ' + String(result.text).replace(/\s+/g, ' ').slice(0, 300));
  for (const text of metrics.assistantTexts) lines.push('ASST   ' + String(text).replace(/\s+/g, ' ').slice(0, 300));
  writeFileSync(join(outDir, 'transcript.txt'), lines.join('\n') + '\n', 'utf8');
  const { toolCalls, toolResults, ...rest } = metrics;
  writeFileSync(join(outDir, 'metrics.json'), JSON.stringify({ label, profile, elapsedMs, exitCode: run.status, sessionFile, ...rest, toolCalls, toolResults }, null, 1) + '\n', 'utf8');
  writeFileSync(join(outDir, 'session.txt'), session.file + '\n', 'utf8');
}

console.log('PROBE ' + label + ' exit=' + run.status + ' elapsed=' + Math.round(elapsedMs / 1000) + 's session=' + (sessionFile ? 'found' : 'MISSING'));
if (metrics) {
  console.log('  tool calls: ' + metrics.toolCalls.length + '  skill loads: ' + metrics.skillLoads + '  wps_help: ' + metrics.helpQueries);
  console.log('  tools: ' + Object.entries(metrics.toolCounts).map(([k, v]) => k + 'x' + v).join(', '));
  console.log('  tokens: in=' + metrics.usage.inputTokens + ' out=' + metrics.usage.outputTokens +
    ' cacheRead=' + metrics.usage.cacheReadTokens + ' total=' + metrics.usage.totalTokens + ' (' + metrics.usage.turns + ' turns)');
  console.log('  loaded chars: skill=' + metrics.loadedChars.skill + ' read=' + metrics.loadedChars.read);
  console.log('  final answer (' + metrics.finalAnswer.length + ' chars) 见 ' + join(outDir, 'metrics.json'));
} else {
  console.log('  ! 没找到这次运行的会话日志；stdout/stderr 已保存到 ' + outDir);
  if (stderr) console.log('  stderr: ' + stderr.replace(/\s+/g, ' ').slice(0, 300));
}
if (run.error) console.log('  spawn error: ' + run.error.message);

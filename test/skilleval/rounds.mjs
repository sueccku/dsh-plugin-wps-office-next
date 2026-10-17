// 跑固定场景的 N 轮（D1 的执行器）：每轮 = 复制夹具 → headless 跑一次 → 解轨迹 → 确定性判定。
//
// 用法：
//   node test/skilleval/rounds.mjs --profile <name> --label <new|old> [--rounds 3] [--scenario <key|all>] [--resume]
//
// 约定：
//   * 夹具 `input/` 只读；每轮复制成 `run-<label>-r<N>/`，多轮/多臂互不覆盖；
//   * 需要 DEEPSEEK_API_KEY（headless profile 从环境读）；跑模型会话要花真钱，按 AGENTS.md §1 先问用户；
//   * 结果写 `<scenario>/run-<label>-r<N>/result.json`，汇总写 `_summary/rounds-<label>.{json,md}`；
//   * --resume 跳过已有 result.json 的轮次（中断后接着跑）；
//   * 单轮默认超时 60s：正常一轮 20–80s，超时的轮次记为失败并标 timedOut（不静默重试）。
import { cpSync, existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { SCENARIOS, scenarioByKey } from './scenarios.mjs';
import { findSession, summarize } from './session.mjs';
import { resolveDshLauncher } from './dsh-bin.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '..', '..');
const EVAL_ROOT = join(ROOT, 'test', '.artifacts', 'skill-eval');
const VERIFY = join(HERE, 'verify.mjs');

const argv = process.argv.slice(2);
const opt = (flag) => {
  const i = argv.indexOf(flag);
  return i >= 0 ? argv[i + 1] : undefined;
};
const profile = opt('--profile');
const label = opt('--label') || 'run';
const rounds = Number(opt('--rounds') || 3);
const resume = argv.includes('--resume');
// --reap-wps：每轮开始前用 scripts/reap-wps.ps1 -Since <本臂最早一轮的开始时刻> 清掉评测自己留下的 WPS。
// 不这么做时，第二、三轮跑的是"WPS 已经开着"的状态 —— 场景 1 要的"WPS 未启动"前提就没了。
// -Since 用**最早一轮**的时刻（不是本轮）：否则上一轮留下的残留进程（比本轮早）杀不到。
// 它只结束评测开始之后启动的 WPS，用户更早开的实例不受影响；但测试机上它真的会杀进程。
const reapWps = argv.includes('--reap-wps');
const only = opt('--scenario');
// 正常一轮 20–80s；60s 是刻意收紧的上限（用户要求，别让卡死的轮次拖时间），超时记为失败。
const timeoutSec = Number(opt('--timeout') || 60);

// --summarize-only：只从已落盘的 result.json 重建汇总（不跑模型、不需要 profile/API key）。
// 改了判定口径或想重算时用；函数声明是提升的，这里可以直接调。
if (argv.includes('--summarize-only')) {
  mkdirSync(join(EVAL_ROOT, '_summary'), { recursive: true });
  writeSummary(loadCompleted(label));
  process.exit(0);
}

if (!profile) {
  console.error('用法: node test/skilleval/rounds.mjs --profile <name> --label <new|old> [--rounds 3] [--scenario <key|all>] [--resume]');
  process.exit(2);
}
if (!process.env.DEEPSEEK_API_KEY) {
  console.error('缺少 DEEPSEEK_API_KEY（headless profile 从环境读模型凭据）');
  process.exit(2);
}
const scenarios = (!only || only === 'all')
  ? SCENARIOS.filter((s) => s.ready)
  : [scenarioByKey(only)].filter(Boolean);
if (!scenarios.length) {
  console.error('没有匹配的 ready 场景：' + only);
  process.exit(2);
}
const launcher = resolveDshLauncher();
if (!launcher) {
  console.error('找不到 dsh 启动器');
  process.exit(2);
}

const env = { ...process.env };
delete env.WPS_OFFICE_MCP_ENTRY;
delete env.WPS_OFFICE_HOST_SCRIPT;
delete env.WPS_OFFICE_MCP_NODE;

function scenarioDir(scenario) {
  return join(EVAL_ROOT, String(scenario.id).padStart(2, '0') + '-' + scenario.key);
}

/** 只结束"某个时刻之后启动的" WPS 进程（scripts/reap-wps.ps1 -Since）。 */
function reapSince(date) {
  if (!reapWps) return;
  const reaper = join(ROOT, 'scripts', 'reap-wps.ps1');
  if (!existsSync(reaper)) return;
  // 必须带 -ExecutionPolicy Bypass：这台机器上直接 -File 会报
  // "AuthorizationManager check failed"（本仓库其它脚本也是这么调的）。
  const result = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', reaper, '-Since', date.toISOString()], {
    encoding: 'utf8',
    windowsHide: true,
  });
  if (result.status !== 0) {
    const detail = String(result.stderr || result.stdout || '').replace(/\s+/g, ' ').slice(0, 160);
    console.log('warn  清场失败（继续跑）：' + detail);
  }
}

/** 从 verify.mjs 的输出里抽 PASS/FAIL/SKIP 行与计数。 */
function runVerify(scenario, runDir) {
  const baseline = join(scenarioDir(scenario), 'fingerprint.json');
  const args = [VERIFY, scenario.key, '--dir', runDir];
  if (existsSync(baseline)) args.push('--baseline', baseline);
  const run = spawnSync(process.execPath, args, { encoding: 'utf8', windowsHide: true });
  const output = String(run.stdout || '') + String(run.stderr || '');
  const lines = output.split(/\r?\n/).filter((line) => /^(PASS|FAIL|SKIP) /.test(line));
  return {
    exitCode: run.status,
    pass: lines.filter((l) => l.startsWith('PASS ')).length,
    fail: lines.filter((l) => l.startsWith('FAIL ')).length,
    skip: lines.filter((l) => l.startsWith('SKIP ')).length,
    lines,
  };
}

function isWpsCall(name) {
  return /^mcp__wps-office-next__/.test(name) || /^wps_/.test(name);
}

const results = [];
// 本臂的起点要按**已有轮次目录**推：--resume 时"现在"比早先那几轮晚，用 now 会漏清早先留下的 WPS。
let runStartedAt = Date.now();
for (const scenario of SCENARIOS) {
  const dir = scenarioDir(scenario);
  if (!existsSync(dir)) continue;
  for (const entry of readdirSync(dir)) {
    if (!entry.startsWith('run-' + label + '-r')) continue;
    try {
      const created = statSync(join(dir, entry)).mtimeMs;
      if (created < runStartedAt) runStartedAt = created;
    } catch {
      // 目录刚好被删掉就忽略
    }
  }
}
mkdirSync(join(EVAL_ROOT, '_summary'), { recursive: true });

for (const scenario of scenarios) {
  for (let round = 1; round <= rounds; round++) {
    const runDir = join(scenarioDir(scenario), 'run-' + label + '-r' + round);
    const resultFile = join(runDir, 'result.json');
    if (resume && existsSync(resultFile)) {
      const previous = JSON.parse(readFileSync(resultFile, 'utf8'));
      results.push(previous);
      console.log('skip  ' + scenario.key + ' r' + round + '（已有结果）');
      continue;
    }
    // 清场必须在动文件**之前**：上一轮超时被杀的 WPS 可能还开着这个目录里的 xlsx，
    // 文件被锁时 rmSync 会 EPERM（实测踩过）。同时让"WPS 未启动"这个前提每轮都成立。
    reapSince(new Date(runStartedAt));
    rmSync(runDir, { recursive: true, force: true });
    mkdirSync(runDir, { recursive: true });
    cpSync(join(scenarioDir(scenario), 'input'), runDir, { recursive: true });

    const startedAt = Date.now();
    const run = spawnSync(launcher.command, launcher.args.concat(['--profile', profile, scenario.prompt]), {
      cwd: runDir,
      encoding: 'utf8',
      windowsHide: true,
      timeout: timeoutSec * 1000,
      env: { ...env, ...launcher.env },
    });
    const elapsedMs = Date.now() - startedAt;
    const timedOut = !!(run.error && (run.error.code === 'ETIMEDOUT' || /timed? ?out/i.test(run.error.message || '')))
      || run.signal === 'SIGTERM';
    if (timedOut && run.pid) {
      // spawnSync 的超时只结束直接子进程，孙进程（MCP server / COM 宿主）可能活下来并**占着 WPS 租约**，
      // 让下一轮直接失败。所以超时后整棵进程树一起收掉。
      spawnSync('taskkill', ['/PID', String(run.pid), '/T', '/F'], { encoding: 'utf8', windowsHide: true });
    }

    let metrics = null;
    let sessionFile = null;
    try {
      const session = findSession({ dshHome: process.env.DSH_HOME || '', task: scenario.prompt, sinceMs: startedAt });
      if (session) {
        sessionFile = session.file;
        metrics = summarize(session.records);
      }
    } catch {
      metrics = null;
    }

    // verify.mjs 的 transcript-* 断言在同目录找 transcript.txt，这里把轨迹落一份进去。
    if (metrics) {
      const lines = [];
      for (const call of metrics.toolCalls) lines.push('CALL   ' + call.name + ' ' + call.args.replace(/\s+/g, ' ').slice(0, 300));
      for (const result of metrics.toolResults) lines.push('RESULT ' + result.name + ' ' + result.text.replace(/\s+/g, ' ').slice(0, 600));
      for (const text of metrics.assistantTexts) lines.push('ASST   ' + text.replace(/\s+/g, ' '));
      writeFileSync(join(runDir, 'transcript.txt'), lines.join('\n') + '\n', 'utf8');
      writeFileSync(join(runDir, 'stdout.txt'), String(run.stdout || ''), 'utf8');
      writeFileSync(join(runDir, 'stderr.txt'), String(run.stderr || ''), 'utf8');
    }

    const verification = runVerify(scenario, runDir);
    const usedWps = metrics ? metrics.toolCalls.filter((c) => isWpsCall(c.name)).length : 0;
    const record = {
      scenario: scenario.key,
      id: scenario.id,
      label,
      round,
      exitCode: run.status,
      timedOut,
      elapsedMs,
      sessionFile,
      toolCounts: metrics ? metrics.toolCounts : {},
      toolCalls: metrics ? metrics.toolCalls.map((c) => c.name) : [],
      skillLoads: metrics ? metrics.skillLoads : 0,
      helpQueries: metrics ? metrics.helpQueries : 0,
      usedWpsTools: usedWps,
      usage: metrics ? metrics.usage : null,
      loadedChars: metrics ? metrics.loadedChars : null,
      finalAnswer: metrics ? metrics.finalAnswer : '',
      verify: verification,
      // 判定只看**产物正确性**（verify 全绿）。超时是延迟问题，单独记 timedOut / 超时率 ——
      // 把"文件其实改对了但 agent 超过 60s"算成任务失败，会污染成功率这个最重要的指标。
      ok: verification.fail === 0 && verification.exitCode === 0,
    };
    writeFileSync(resultFile, JSON.stringify(record, null, 1) + '\n', 'utf8');
    results.push(record);
    console.log(
      (record.ok ? 'PASS ' : 'FAIL ') + scenario.key.padEnd(20) + ' ' + label + ' r' + round +
      (timedOut ? ' TIMEOUT(' + timeoutSec + 's)' : '') +
      '  verify ' + verification.pass + '/' + (verification.pass + verification.fail) +
      (verification.skip ? ' (+' + verification.skip + ' skip)' : '') +
      '  wps=' + usedWps + '/tools=' + (metrics ? metrics.toolCalls.length : 0) +
      '  tokens=' + (metrics ? metrics.usage.totalTokens : '?') +
      '  ' + Math.round(elapsedMs / 1000) + 's',
    );
  }
}

// ---- 汇总 ----
/** 从逐轮结果生成 `_summary/rounds-<label>.{json,md}`。成功率和超时率分开报。 */
function writeSummary(records) {
  // 判定一律**从 verify 重算**，不信任 result.json 里的旧 ok 字段 —— 口径改过一次（超时不再算失败），
  // 旧文件里存的 ok 是旧口径，直接读会得到自相矛盾的汇总。
  const isOk = (record) => !!(record.verify && record.verify.fail === 0 && record.verify.exitCode === 0);
  const byScenario = new Map();
  for (const record of records) {
    if (!byScenario.has(record.scenario)) byScenario.set(record.scenario, []);
    byScenario.get(record.scenario).push(record);
  }
  const summary = {
    label,
    profile: profile || '(summarize-only)',
    rounds,
    generatedAt: new Date().toISOString(),
    scenarios: [...byScenario.entries()].map(([key, list]) => {
      const ok = list.filter(isOk).length;
      const timeouts = list.filter((r) => r.timedOut).length;
      const tokens = list.map((r) => (r.usage ? r.usage.totalTokens : 0));
      const wps = list.map((r) => r.usedWpsTools || 0);
      return {
        scenario: key,
        rounds: list.length,
        ok,
        timeouts,
        successRate: list.length ? Math.round((ok / list.length) * 100) : 0,
        avgTokens: Math.round(tokens.reduce((a, b) => a + b, 0) / Math.max(1, list.length)),
        avgWpsToolCalls: (wps.reduce((a, b) => a + b, 0) / Math.max(1, list.length)).toFixed(1),
        failedChecks: list
          .filter((r) => !isOk(r))
          .flatMap((r) => r.verify.lines.filter((l) => l.startsWith('FAIL ')))
          .slice(0, 6),
        perRound: list.map((r) => ({ round: r.round, ok: isOk(r), timedOut: !!r.timedOut, exitCode: r.exitCode, elapsedMs: r.elapsedMs })),
      };
    }),
  };
  writeFileSync(join(EVAL_ROOT, '_summary', 'rounds-' + label + '.json'), JSON.stringify(summary, null, 1) + '\n', 'utf8');

  const md = [
    '# D1 轮次汇总（' + label + '）',
    '',
    '- profile: ' + profile,
    '- 每场景轮数: ' + rounds,
    '- 生成时间: ' + summary.generatedAt,
    '- 判定：`通过` = `verify.mjs` 的确定性断言全绿（不看模型自述）；超时单独计，不算判定失败。',
    '',
    '| 场景 | 轮数 | 通过 | 成功率 | 超时 | 平均 tokens | 平均 WPS 工具调用 | 主要失败断言 |',
    '|---|---|---|---|---|---|---|---|',
    ...summary.scenarios.map((s) => '| ' + s.scenario + ' | ' + s.rounds + ' | ' + s.ok + ' | ' + s.successRate + '% | ' + s.timeouts + ' | ' + s.avgTokens + ' | ' + s.avgWpsToolCalls + ' | ' + (s.failedChecks.length ? s.failedChecks.map((l) => l.replace(/\|/g, '/')).join('<br>') : '—') + ' |'),
    '',
  ];
  writeFileSync(join(EVAL_ROOT, '_summary', 'rounds-' + label + '.md'), md.join('\n'), 'utf8');
  console.log('');
  console.log('汇总: ' + join(EVAL_ROOT, '_summary', 'rounds-' + label + '.md'));
  console.log('总轮次=' + records.length + '  通过=' + records.filter(isOk).length + '  超时=' + records.filter((r) => r.timedOut).length);
}

function loadCompleted(labelName) {
  const loaded = [];
  for (const scenario of SCENARIOS.filter((s) => s.ready)) {
    for (let round = 1; round <= rounds; round++) {
      const file = join(scenarioDir(scenario), 'run-' + labelName + '-r' + round, 'result.json');
      if (!existsSync(file)) continue;
      try {
        loaded.push(JSON.parse(readFileSync(file, 'utf8')));
      } catch {
        // 坏文件跳过：汇总不替它编数
      }
    }
  }
  return loaded;
}

writeSummary(results);
// 收尾：把本次评测自己启动的 WPS 清掉（只清 runStartedAt 之后启动的）。
reapSince(new Date(runStartedAt));
process.exit(results.every((r) => r.ok) ? 0 : 1);

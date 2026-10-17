// 把两条臂（new / old）的轮次汇总并成一张对比表（改进指南 §10.3 的口径）。
//
// 用法：node test/skilleval/compare.mjs [--new new] [--old old]
// 输入：test/.artifacts/skill-eval/_summary/rounds-<label>.json（由 rounds.mjs 生成）
// 输出：test/.artifacts/skill-eval/_summary/compare.md / compare.json（本机产物，.gitignore 忽略）
//
// 统计口径都是可复算的：成功率按 verify 全绿计；token 取会话日志里的真实 usage；
// "无依据的成功声明"是**启发式**代理指标（判定失败、但最终回答里出现"已完成/已保存/已修改"之类措辞），
// 报告里必须标明它是启发式，不能当成严格审计结果。
import { existsSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const EVAL_ROOT = join(resolve(HERE, '..', '..'), 'test', '.artifacts', 'skill-eval');
const SUMMARY = join(EVAL_ROOT, '_summary');

const argv = process.argv.slice(2);
const opt = (flag) => {
  const i = argv.indexOf(flag);
  return i >= 0 ? argv[i + 1] : undefined;
};
const newLabel = opt('--new') || 'new';
const oldLabel = opt('--old') || 'old';

function loadArm(label) {
  const file = join(SUMMARY, 'rounds-' + label + '.json');
  if (!existsSync(file)) {
    console.error('缺少汇总：' + file + '（先跑 rounds.mjs --label ' + label + '）');
    process.exit(2);
  }
  const data = JSON.parse(readFileSync(file, 'utf8'));
  const byScenario = new Map(data.scenarios.map((s) => [s.scenario, s]));
  // 逐轮记录在 scenarios 的 perRound 里；无依据成功声明要从每轮的 result.json 读最终回答。
  const rounds = [];
  for (const scenario of data.scenarios) {
    for (const round of scenario.perRound) {
      const resultFile = join(EVAL_ROOT, findScenarioDir(scenario.scenario), 'run-' + label + '-r' + round.round, 'result.json');
      let answer = '';
      if (existsSync(resultFile)) {
        try {
          answer = JSON.parse(readFileSync(resultFile, 'utf8')).finalAnswer || '';
        } catch {
          answer = '';
        }
      }
      rounds.push({ scenario: scenario.scenario, round: round.round, ok: round.ok, elapsedMs: round.elapsedMs, answer });
    }
  }
  return { data, byScenario, rounds };
}

function findScenarioDir(key) {
  // 目录名是 <NN>-<key>；按 key 找一次，避免在 compare 里重复维护编号。
  return readdirSync(EVAL_ROOT).find((name) => name.endsWith('-' + key)) || '';
}

const SUCCESS_CLAIM = /(已(完成|保存|修改|改为|写入|保存到)|完成[:：]|成功保存)/;
function falseSuccessCount(arm) {
  return arm.rounds.filter((r) => !r.ok && SUCCESS_CLAIM.test(r.answer || '')).length;
}

function totals(arm) {
  const rounds = arm.rounds.length;
  const ok = arm.rounds.filter((r) => r.ok).length;
  const tokens = arm.data.scenarios.reduce((sum, s) => sum + s.avgTokens * s.rounds, 0);
  const wpsCalls = arm.data.scenarios.reduce((sum, s) => sum + Number(s.avgWpsToolCalls) * s.rounds, 0);
  const elapsed = arm.rounds.reduce((sum, r) => sum + r.elapsedMs, 0);
  return {
    rounds,
    ok,
    successRate: rounds ? Math.round((ok / rounds) * 100) : 0,
    avgTokens: rounds ? Math.round(tokens / rounds) : 0,
    avgWpsCalls: rounds ? (wpsCalls / rounds).toFixed(1) : '0',
    avgElapsedSec: rounds ? Math.round(elapsed / rounds / 1000) : 0,
    falseSuccess: falseSuccessCount(arm),
  };
}

const fresh = loadArm(newLabel);
const old = loadArm(oldLabel);
const newTotals = totals(fresh);
const oldTotals = totals(old);

const keys = [...new Set([...fresh.byScenario.keys(), ...old.byScenario.keys()])];
const rows = keys.map((key) => {
  const a = fresh.byScenario.get(key);
  const b = old.byScenario.get(key);
  return {
    scenario: key,
    [newLabel]: a ? { ok: a.ok, rounds: a.rounds, rate: a.successRate, tokens: a.avgTokens, wps: a.avgWpsToolCalls, failed: a.failedChecks } : null,
    [oldLabel]: b ? { ok: b.ok, rounds: b.rounds, rate: b.successRate, tokens: b.avgTokens, wps: b.avgWpsToolCalls, failed: b.failedChecks } : null,
    deltaRate: a && b ? a.successRate - b.successRate : null,
    deltaTokens: a && b ? a.avgTokens - b.avgTokens : null,
  };
});

const md = [
  '# D1 A/B 对比（' + newLabel + ' vs ' + oldLabel + '）',
  '',
  '> 判定口径：每轮由 `verify.mjs` 的确定性断言判定（不看模型自述）。token 是会话日志里的真实 usage。',
  '',
  '## 总览',
  '',
  '| 臂 | 轮次 | 通过 | 成功率 | 平均 tokens/轮 | 平均 WPS 工具调用/轮 | 平均耗时 | 无依据成功声明（启发式） |',
  '|---|---|---|---|---|---|---|---|',
  '| ' + newLabel + ' | ' + newTotals.rounds + ' | ' + newTotals.ok + ' | ' + newTotals.successRate + '% | ' + newTotals.avgTokens + ' | ' + newTotals.avgWpsCalls + ' | ' + newTotals.avgElapsedSec + 's | ' + newTotals.falseSuccess + ' |',
  '| ' + oldLabel + ' | ' + oldTotals.rounds + ' | ' + oldTotals.ok + ' | ' + oldTotals.successRate + '% | ' + oldTotals.avgTokens + ' | ' + oldTotals.avgWpsCalls + ' | ' + oldTotals.avgElapsedSec + 's | ' + oldTotals.falseSuccess + ' |',
  '',
  '## 逐场景',
  '',
  '| 场景 | ' + newLabel + ' 通过 | ' + oldLabel + ' 通过 | 成功率差 | 平均 tokens 差 | ' + newLabel + ' 失败断言 |',
  '|---|---|---|---|---|---|',
  ...rows.map((row) => {
    const a = row[newLabel];
    const b = row[oldLabel];
    const cell = (v) => (v ? v.ok + '/' + v.rounds + ' (' + v.rate + '%)' : '—');
    return '| ' + row.scenario + ' | ' + cell(a) + ' | ' + cell(b) + ' | ' +
      (row.deltaRate === null ? '—' : (row.deltaRate > 0 ? '+' : '') + row.deltaRate + 'pp') + ' | ' +
      (row.deltaTokens === null ? '—' : (row.deltaTokens > 0 ? '+' : '') + row.deltaTokens) + ' | ' +
      (a && a.failed.length ? a.failed.map((l) => l.replace(/\|/g, '/')).join('<br>') : '—') + ' |';
  }),
  '',
  '> 「无依据成功声明」= 判定失败、但最终回答出现"已完成/已保存/已修改"等措辞的轮次数。这是**启发式**代理，',
  '> 只能提示"模型可能过度自信"，不等于逐条审计。',
  '',
].join('\n');

writeFileSync(join(SUMMARY, 'compare.json'), JSON.stringify({ newLabel, oldLabel, newTotals, oldTotals, rows }, null, 1) + '\n', 'utf8');
writeFileSync(join(SUMMARY, 'compare.md'), md, 'utf8');
console.log(md);
process.exit(0);

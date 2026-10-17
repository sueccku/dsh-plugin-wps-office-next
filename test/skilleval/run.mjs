// 场景评估的公共入口：**只准备与判定，不自己决定要不要跑模型**。
//
// 用法：
//   node test/skilleval/run.mjs --dry-run [--scenario <key|all>] [--python <exe>]
//       生成（或重建）每个场景的输入夹具，算出输入指纹，写出 prompt.txt / checks.txt。
//       这一步不需要 WPS、不调模型，用于验证"夹具 + 断言清单"本身是正确的。
//   node test/skilleval/run.mjs --print-command --profile <name> [--scenario <key|all>]
//       只打印"要跑这个场景该执行什么命令"，**不执行**。真正跑模型会话属于 P2，
//       按 AGENTS.md §1 需要用户明确同意。
//   node test/skilleval/run.mjs --verify <key> [--dir <工作目录>] [--python <exe>]
//       对已有产物做确定性判定（转调 verify.mjs）。
//
// 产物目录 test/.artifacts/skill-eval/ 已被 .gitignore 忽略：夹具是机器本地产物，不入库；
// 这里的**代码**入库，保证同一份场景可以重跑。
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { FIXTURE_KEYS, SCENARIOS, scenarioByKey } from './scenarios.mjs';
import { resolvePython } from './python.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '..', '..');
const EVAL_ROOT = join(ROOT, 'test', '.artifacts', 'skill-eval');
const FIXTURES = join(HERE, 'fixtures.py');
const VERIFY = join(HERE, 'verify.mjs');

function listFiles(dir, prefix = '') {
  const out = [];
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    const rel = prefix ? prefix + '/' + entry : entry;
    if (statSync(full).isDirectory()) out.push(...listFiles(full, rel));
    else out.push(rel);
  }
  return out;
}

function sha256(file) {
  return createHash('sha256').update(readFileSync(file)).digest('hex');
}

function scenarioDir(scenario) {
  return join(EVAL_ROOT, String(scenario.id).padStart(2, '0') + '-' + scenario.key);
}

/** 表与 fixtures.py 必须一一对应，防止"写了场景但夹具脚本没实现"这种静默漂移。 */
function validateTable() {
  const used = SCENARIOS.filter((s) => s.fixture).map((s) => s.fixture);
  const errors = [];
  for (const key of used) if (!FIXTURE_KEYS.includes(key)) errors.push('场景引用了未声明的夹具 key：' + key);
  for (const key of FIXTURE_KEYS) if (!used.includes(key)) errors.push('FIXTURE_KEYS 里多了一个没人用的 key：' + key);
  const keys = SCENARIOS.map((s) => s.key);
  for (const key of keys) if (keys.indexOf(key) !== keys.lastIndexOf(key)) errors.push('场景 key 重复：' + key);
  for (const s of SCENARIOS) {
    if (s.ready && !s.checks.length) errors.push('场景声明 ready 却没有断言：' + s.key);
    if (!s.ready && !s.coreCheck) errors.push('未 ready 的场景必须写 coreCheck（人读的成功标准）：' + s.key);
  }
  return errors;
}

function selected(argv) {
  const i = argv.indexOf('--scenario');
  const wanted = i >= 0 ? argv[i + 1] : 'all';
  if (!wanted || wanted === 'all') return SCENARIOS;
  const one = scenarioByKey(wanted);
  if (!one) {
    console.error('未知场景：' + wanted);
    console.error('可用场景：' + SCENARIOS.map((s) => s.id + ':' + s.key).join(', '));
    process.exit(2);
  }
  return [one];
}

function optValue(argv, flag) {
  const i = argv.indexOf(flag);
  return i >= 0 ? argv[i + 1] : undefined;
}

function dryRun(python, scenarios, force) {
  mkdirSync(EVAL_ROOT, { recursive: true });
  let readyCount = 0;
  let todoCount = 0;
  for (const scenario of scenarios) {
    const dir = scenarioDir(scenario);
    const input = join(dir, 'input');
    if (!scenario.ready) {
      mkdirSync(dir, { recursive: true });
      writeFileSync(join(dir, 'prompt.txt'), scenario.prompt + '\n', 'utf8');
      writeFileSync(join(dir, 'checks.txt'), '# 待补：' + (scenario.note || '') + '\n# 成功标准：' + scenario.coreCheck + '\n', 'utf8');
      todoCount++;
      console.log('todo  ' + String(scenario.id).padStart(2, '0') + ' ' + scenario.key.padEnd(20) + ' 待补夹具/判定：' + (scenario.note || ''));
      continue;
    }

    // 默认**不重建**已存在的 input/：那是"夹具目录"，跑过一轮之后里面会有产物，
    // 直接 rm -rf 会把上一轮的结果悄悄清掉（多轮对比时最致命）。要重建就显式 --force。
    if (existsSync(input) && !force) {
      const files = existsSync(join(dir, 'fingerprint.json'))
        ? Object.keys(JSON.parse(readFileSync(join(dir, 'fingerprint.json'), 'utf8')).files || {})
        : [];
      readyCount++;
      console.log('keep  ' + String(scenario.id).padStart(2, '0') + ' ' + scenario.key.padEnd(20) + ' 已存在（要重建加 --force）| 夹具=' + (files.join(',') || '无'));
      continue;
    }

    rmSync(input, { recursive: true, force: true });
    mkdirSync(input, { recursive: true });
    if (scenario.fixture) {
      const run = spawnSync(python, [FIXTURES, scenario.fixture, input], {
        encoding: 'utf8',
        windowsHide: true,
        env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
      });
      if (run.status !== 0) {
        console.error('FAIL  夹具生成失败（' + scenario.key + ' / ' + scenario.fixture + '）');
        console.error(String(run.stderr || run.stdout || '').trim());
        process.exitCode = 1;
        continue;
      }
    }

    const files = {};
    for (const rel of listFiles(input)) {
      const file = join(input, rel);
      files[rel] = { size: statSync(file).size, sha256: sha256(file) };
    }
    writeFileSync(join(dir, 'fingerprint.json'), JSON.stringify({ scenario: scenario.key, id: scenario.id, files }, null, 1) + '\n', 'utf8');
    writeFileSync(join(dir, 'prompt.txt'), scenario.prompt + '\n', 'utf8');
    writeFileSync(
      join(dir, 'checks.txt'),
      ['# 场景 ' + scenario.id + ' ' + scenario.title, '# 提示：' + scenario.prompt, '# 断言：', ...scenario.checks.map((c) => '  - ' + c), '# 成功标准：' + scenario.coreCheck, ''].join('\n'),
      'utf8',
    );
    readyCount++;
    const fixtureText = scenario.fixture ? '夹具=' + Object.keys(files).join(',') : '无夹具';
    console.log('ready ' + String(scenario.id).padStart(2, '0') + ' ' + scenario.key.padEnd(20) + ' ' + fixtureText + ' | 断言 ' + scenario.checks.length + ' 条');
  }
  console.log('');
  console.log('dry-run 完成：ready=' + readyCount + ' todo=' + todoCount + '  产物目录=' + EVAL_ROOT);
  console.log('产物目录被 .gitignore 忽略；夹具与断言清单可直接用 node test/skilleval/verify.mjs <key> 复核。');
}

function printCommands(profile, scenarios) {
  if (!profile) {
    console.error('--print-command 需要 --profile <name>（例如安装本仓库的 headless profile）');
    process.exit(2);
  }
  console.log('# 以下命令**不会被执行**，只是把"跑一轮"的确切动作写清楚。');
  console.log('# 真正跑模型会话属于 P2，按 AGENTS.md §1 需要用户明确同意；执行前先跑 --dry-run 生成夹具。');
  console.log('# 每一轮都从 input/ 复制一份 run-rN/ 再跑：夹具保持纯净，多轮对比互不覆盖。');
  console.log('');
  for (const scenario of scenarios) {
    if (!scenario.ready) continue;
    const dir = scenarioDir(scenario);
    const prompt = scenario.prompt.replace(/"/g, '\\"');
    const round = join(dir, 'run-r1');
    console.log('# ' + scenario.id + ' ' + scenario.title);
    console.log('Copy-Item -Recurse -Force "' + join(dir, 'input') + '" "' + round + '"');
    console.log('cd "' + round + '"');
    console.log('dsh --profile ' + profile + ' "' + prompt + '"');
    console.log('node test\\skilleval\\verify.mjs ' + scenario.key + ' --dir "' + round + '"');
    console.log('');
  }
}

const argv = process.argv.slice(2);
const wantsVerify = argv.includes('--verify');
const wantsPrint = argv.includes('--print-command');
const wantsForce = argv.includes('--force');
const wantsDryRun = argv.includes('--dry-run') || (!wantsVerify && !wantsPrint);

const tableErrors = validateTable();
if (tableErrors.length) {
  console.error('scenarios.mjs / fixtures.py 表不一致：');
  for (const error of tableErrors) console.error('  x ' + error);
  process.exit(1);
}

let python;
try {
  python = resolvePython(optValue(argv, '--python'));
} catch (error) {
  console.error(error.message);
  process.exit(2);
}

if (wantsVerify) {
  const run = spawnSync(process.execPath, [VERIFY, ...argv.filter((a) => a !== '--verify')], { stdio: 'inherit', windowsHide: true });
  process.exit(run.status === null ? 2 : run.status);
}

const scenarios = selected(argv);
if (wantsDryRun) dryRun(python, scenarios, wantsForce);
if (wantsPrint) printCommands(optValue(argv, '--profile'), scenarios);
if (!existsSync(EVAL_ROOT)) {
  console.error('产物目录不存在：' + EVAL_ROOT);
  process.exit(1);
}
process.exit(process.exitCode === undefined ? 0 : process.exitCode);

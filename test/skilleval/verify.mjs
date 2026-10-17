// 对一个场景的产物做**确定性**判定（不读模型自述）。
//
// 用法：
//   node test/skilleval/verify.mjs <scenario-key> [--dir <工作目录>] [--baseline <fingerprint.json>] [--python <exe>]
//
// 工作目录默认是 test/.artifacts/skill-eval/<NN>-<key>/input/（就是被测 Agent 的 cwd）。
// baseline 是 run.mjs --dry-run 落盘的输入指纹，用来判 "unchanged:<rel>"（原文件没被改）。
//
// 断言语法（scenarios.mjs 的 checks）：
//   file-exists:<rel>                       文件存在
//   file-absent:<rel>                        文件不存在
//   docx-text:<rel>|<子串>                   docx 正文/表格文本里包含该子串
//   xlsx-value:<rel>|<sheet>|<cell>|<期望>   单元格值等于期望（数字按数值比）
//   xlsx-formula:<rel>|<sheet>|<cell>|<期望> 单元格**公式文本**等于期望（去空白、去前导 =、忽略大小写）
//   unchanged:<rel>                          与 baseline 指纹一致（原文件未被修改）
//   transcript-matches:<正则>               会话/运行轨迹文本匹配该正则
//   transcript-not-matches:<正则>           会话/运行轨迹文本不匹配
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { SCENARIOS, scenarioByKey } from './scenarios.mjs';
import { resolvePython } from './python.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '..', '..');
const EVAL_ROOT = join(ROOT, 'test', '.artifacts', 'skill-eval');
const INSPECT = join(HERE, 'ooxml_probe.py');

const results = [];
function record(kind, name, detail) {
  results.push({ kind, name, detail });
  console.log(kind.padEnd(4) + ' ' + name + (detail ? '  ' + detail : ''));
}

function scenarioDir(key) {
  const scenario = scenarioByKey(key);
  if (!scenario) throw new Error('unknown scenario key: ' + key);
  return join(EVAL_ROOT, String(scenario.id).padStart(2, '0') + '-' + scenario.key);
}

function sha256(file) {
  return createHash('sha256').update(readFileSync(file)).digest('hex');
}

function pythonJson(python, args) {
  const stdout = execFileSync(python, [INSPECT, ...args], {
    encoding: 'utf8',
    windowsHide: true,
    env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
  });
  return JSON.parse(String(stdout).trim());
}

function normFormula(value) {
  return String(value === null || value === undefined ? '' : value)
    .trim()
    .replace(/^=/, '')
    .replace(/\s+/g, '')
    .toUpperCase();
}

function sameValue(actual, expected) {
  const a = actual === null || actual === undefined ? '' : String(actual).trim();
  const e = String(expected).trim();
  if (a === e) return true;
  const na = Number(a);
  const ne = Number(e);
  return Number.isFinite(na) && Number.isFinite(ne) && na === ne;
}

/** 轨迹文本：优先 transcript.txt / run.stdout.txt / trace.txt，都没有就返回 null（调用方给 SKIP）。 */
function transcriptText(dir) {
  const candidates = ['transcript.txt', 'run.stdout.txt', 'trace.txt'];
  const texts = [];
  for (const name of candidates) {
    const file = join(dir, name);
    if (existsSync(file)) texts.push(readFileSync(file, 'utf8'));
  }
  if (!texts.length) {
    const found = existsSync(dir) ? readdirSync(dir).filter((f) => f.endsWith('.stdout.txt') || f.endsWith('.txt')) : [];
    for (const name of found) texts.push(readFileSync(join(dir, name), 'utf8'));
  }
  return texts.length ? texts.join('\n') : null;
}

const checks = [];
function parseCheck(spec) {
  const i = spec.indexOf(':');
  if (i < 0) return { kind: spec, parts: [] };
  const kind = spec.slice(0, i);
  const rest = spec.slice(i + 1);
  // transcript-* 的值本身可能含 |（正则），所以只有它们不按 | 切。
  if (kind.startsWith('transcript-')) return { kind, parts: [rest] };
  return { kind, parts: rest.split('|') };
}

function runCheck(spec, { dir, baseline, python, scenario }) {
  let parsed;
  try {
    parsed = parseCheck(spec);
  } catch (error) {
    record('FAIL', spec, '断言语法错误：' + error.message);
    return;
  }
  const { kind, parts } = parsed;
  const rel = parts[0];
  const abs = rel ? join(dir, rel) : '';

  switch (kind) {
    case 'file-exists':
      record(existsSync(abs) ? 'PASS' : 'FAIL', spec);
      return;
    case 'file-absent':
      record(!existsSync(abs) ? 'PASS' : 'FAIL', spec, existsSync(abs) ? '文件不应存在' : '');
      return;
    case 'unchanged': {
      if (!baseline) { record('SKIP', spec, '没有 baseline（先跑 run.mjs --dry-run）'); return; }
      const before = baseline.files && baseline.files[rel];
      if (!before) { record('FAIL', spec, 'baseline 里没有这个文件'); return; }
      if (!existsSync(abs)) { record('FAIL', spec, '文件不见了'); return; }
      const now = sha256(abs);
      record(now === before.sha256 ? 'PASS' : 'FAIL', spec, now === before.sha256 ? '' : '内容被改动');
      return;
    }
    case 'docx-text': {
      if (!existsSync(abs)) { record('FAIL', spec, '文件不存在'); return; }
      const text = pythonJson(python, ['docx-text', abs]).text || '';
      const ok = text.includes(parts[1]);
      record(ok ? 'PASS' : 'FAIL', spec, ok ? '' : '实际文本：' + text.replace(/\s+/g, ' ').slice(0, 120));
      return;
    }
    case 'xlsx-value': {
      if (!existsSync(abs)) { record('FAIL', spec, '文件不存在'); return; }
      const cell = pythonJson(python, ['xlsx-cell', abs, parts[1], parts[2]]);
      const ok = sameValue(cell.value, parts[3]);
      record(ok ? 'PASS' : 'FAIL', spec, ok ? '' : '实际值：' + JSON.stringify(cell.value));
      return;
    }
    case 'xlsx-formula': {
      if (!existsSync(abs)) { record('FAIL', spec, '文件不存在'); return; }
      const cell = pythonJson(python, ['xlsx-cell', abs, parts[1], parts[2]]);
      const actual = normFormula(cell.formula);
      const expected = normFormula(parts[3]);
      const ok = actual.length > 0 && actual === expected;
      record(ok ? 'PASS' : 'FAIL', spec, ok ? '' : '实际公式：' + JSON.stringify(cell.formula) + '（期望 ' + parts[3] + '）');
      return;
    }
    case 'transcript-matches':
    case 'transcript-not-matches': {
      const text = transcriptText(dir);
      if (text === null) { record('SKIP', spec, '找不到会话轨迹文本'); return; }
      const hit = new RegExp(parts[0], 'm').test(text);
      const want = kind === 'transcript-matches';
      record(hit === want ? 'PASS' : 'FAIL', spec, hit ? '命中' : '未命中');
      return;
    }
    default:
      record('FAIL', spec, '未知断言类型：' + kind + '（scenario=' + (scenario ? scenario.key : '?') + '）');
  }
}

const argv = process.argv.slice(2);
const opts = {};
let key = null;
for (let i = 0; i < argv.length; i++) {
  const arg = argv[i];
  if (arg === '--dir' || arg === '--baseline' || arg === '--python') {
    opts[arg.slice(2)] = argv[++i];
    continue;
  }
  if (arg.startsWith('--')) {
    console.error('未知参数：' + arg);
    process.exit(2);
  }
  key = key || arg;
}
const scenario = key ? scenarioByKey(key) : null;
if (!scenario) {
  console.error('用法: node test/skilleval/verify.mjs <scenario-key> [--dir <dir>] [--baseline <fingerprint.json>] [--python <exe>]');
  console.error('可用场景：' + SCENARIOS.map((s) => s.key).join(', '));
  process.exit(2);
}
const dir = resolve(opts.dir || join(scenarioDir(key), 'input'));
const baselinePath = opts.baseline || join(scenarioDir(key), 'fingerprint.json');
let baseline = null;
if (existsSync(baselinePath)) baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));

let python;
try {
  python = resolvePython(opts.python);
} catch (error) {
  console.error(error.message);
  process.exit(2);
}

if (!existsSync(dir)) {
  console.error('工作目录不存在：' + dir + '（先跑 node test/skilleval/run.mjs --dry-run）');
  process.exit(2);
}

console.log('# 场景 ' + scenario.id + ' ' + scenario.title + '（' + key + '）');
console.log('# 工作目录 ' + dir);
if (!scenario.checks.length) {
  console.log('SKIP  该场景还没有确定性断言（见 scenarios.mjs 的 note）');
  process.exit(0);
}
for (const spec of scenario.checks) runCheck(spec, { dir, baseline, python, scenario });

const failed = results.filter((r) => r.kind === 'FAIL').length;
const passed = results.filter((r) => r.kind === 'PASS').length;
const skipped = results.filter((r) => r.kind === 'SKIP').length;
console.log('SCENARIO ' + key + (failed ? ' FAILED' : ' OK') + ' (' + passed + ' pass / ' + failed + ' fail / ' + skipped + ' skip)');
process.exit(failed ? 1 : 0);

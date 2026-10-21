// 广告面换血计划器（P2-2，只读、不跑模型、不改工具面）。
// 输入：test/.artifacts/skill-eval/**/metrics.json（probe.mjs 的产物，已被 .gitignore 忽略）
// 输出：按 R14 阈值（跨 >=2 轮真实轨迹出现，或单轮被调用 >=5 次）排序的隐藏工具候选。
// 没有轨迹时明确输出 "no evidence"，不允许无证据扩容。
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';

const ROOT = resolve('.');
const ART = join(ROOT, 'test', '.artifacts', 'skill-eval');
const catalog = JSON.parse(readFileSync(join(ROOT, 'spec', 'tool-definitions.json'), 'utf8'));
const advertised = new Set(JSON.parse(readFileSync(join(ROOT, 'spec', 'advertised.json'), 'utf8')));
const allNames = catalog.map((t) => t.name);
const hidden = allNames.filter((n) => !advertised.has(n));

function walk(dir, out = []) {
  if (!existsSync(dir)) return out;
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) walk(full, out);
    else if (entry === 'metrics.json') out.push(full);
  }
  return out;
}

const metrics = walk(ART);
const calls = new Map();
const runs = new Map();
for (const file of metrics) {
  let m;
  try { m = JSON.parse(readFileSync(file, 'utf8')); } catch { continue; }
  const list = m.toolCalls || m.calls || [];
  const seen = new Set();
  for (const c of list) {
    const name = typeof c === 'string' ? c : (c.tool || c.name);
    if (!name || !hidden.includes(name)) continue;
    calls.set(name, (calls.get(name) || 0) + 1);
    if (!seen.has(name)) { runs.set(name, (runs.get(name) || 0) + 1); seen.add(name); }
  }
}
const ranked = [...calls.entries()]
  .map(([tool, count]) => ({ tool, count, runs: runs.get(tool) || 0 }))
  .sort((a, b) => (b.runs - a.runs) || (b.count - a.count));
const eligible = ranked.filter((r) => r.runs >= 2 || r.count >= 5);
console.log('metrics files=' + metrics.length + ' hidden=' + hidden.length + ' advertised=' + advertised.size);
if (!metrics.length) {
  if (advertised.size >= 100) {
    console.log('NO EVIDENCE: 没有真实轨迹 metrics.json；广告面已到 100 工具上限，后续只能换血（本脚本不改工具面）。');
  } else {
    console.log('NO EVIDENCE: 没有真实轨迹 metrics.json；按 R14 不换血、不扩容。');
  }
  process.exit(0);
}
console.log('eligible candidates (R14 threshold):');
for (const r of eligible.slice(0, 20)) console.log('  ' + r.tool + '  runs=' + r.runs + ' calls=' + r.count);
if (!eligible.length) console.log('  (none)');
console.log('NOTE: 换入还需要换出等量工具并保持 <=100 工具 / <=80000 字节；本脚本只给候选，不改广告面。');

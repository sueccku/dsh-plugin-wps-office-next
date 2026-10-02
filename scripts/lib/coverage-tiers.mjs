/**
 * 覆盖率分层的唯一实现。scripts/smoke-tools.mjs（打印与 --check）与 test/spec-reproduction.test.mjs
 * （CI 棘轮）共用同一份口径，避免两处各算一遍然后慢慢漂移（FIXES 74）。
 *
 * 三层口径，从强到弱：
 *   bespoke   —— 至少有一条**非矩阵条目**的代码点到它（矩阵文件里真正的断言代码也算）
 *   matrixOk  —— 只出现在矩阵条目里，期望 ok / error（至少断言了成功或明确失败）
 *   matrixAny —— 只出现在矩阵条目里，期望 any（只断言「没挂住」，最弱的一层）
 *   notDriven —— 哪里都没点到（注释里的名字不算）
 *
 * 旧的单一口径是「名字在测试源码里出现过」，注释里提一句也算 —— 那是必要条件，不是证据。
 */
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

export const MATRIX_FILES = ['ppt-coverage.test.mjs', 'excel-coverage.test.mjs', 'word-common-coverage.test.mjs'];

// 账本：涨即失败、缩要显式改（和 ALIAS_DEBT / UNTOOLED_ACTIONS 一个玩法）。
export const TIER_LEDGER = { bespoke: 168, matrixAny: 77 };

function stripComments(source) {
  return source.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:])\/\/.*$/gm, '$1');
}

export function computeCoverageTiers(root = process.cwd()) {
  const defs = JSON.parse(readFileSync(join(root, 'spec/tool-definitions.json'), 'utf8'));
  const registered = defs.map((d) => d.name);

  const files = readdirSync(join(root, 'test')).filter((f) => f.endsWith('.test.mjs'));
  const sources = files.map((f) => [f, stripComments(readFileSync(join(root, 'test', f), 'utf8'))]);
  sources.push(['scripts/e2e.mjs', stripComments(readFileSync(join(root, 'scripts/e2e.mjs'), 'utf8'))]);

  // 矩阵条目形如 ["wps_x", { ... }, "any"], —— 逐行解析，嵌套花括号不会骗到它。条目行要从
  // 「有没有专门场景」的判断里剔除：矩阵文件里也有真正的断言代码。
  const matrixEntry = /^\s*\["(wps_[a-z0-9_]+)",\s*.+?,\s*"(ok|any|error)"\],?\s*$/;
  const expectation = new Map();
  let matrixEntries = 0;
  const bespokeLines = [];
  for (const [, text] of sources) {
    for (const line of text.split(/\r?\n/)) {
      const m = matrixEntry.exec(line);
      if (m) {
        matrixEntries++;
        if (!expectation.has(m[1])) expectation.set(m[1], m[2]);
        continue;
      }
      bespokeLines.push(line);
    }
  }

  const inBespokeCode = new Set([...bespokeLines.join('\n').matchAll(/\bwps_[a-z0-9_]+/g)].map((m) => m[0]));

  const tiers = { registered, bespoke: [], matrixOk: [], matrixAny: [], notDriven: [], matrixEntries };
  for (const name of registered) {
    if (inBespokeCode.has(name)) tiers.bespoke.push(name);
    else if (expectation.get(name) === 'any') tiers.matrixAny.push(name);
    else if (expectation.has(name)) tiers.matrixOk.push(name);
    else tiers.notDriven.push(name);
  }
  return tiers;
}

if (process.argv[1] && process.argv[1].endsWith('coverage-tiers.mjs')) {
  const t = computeCoverageTiers();
  console.log('registered=' + t.registered.length);
  console.log('bespoke=' + t.bespoke.length);
  console.log('matrixOk=' + t.matrixOk.length);
  console.log('matrixAny=' + t.matrixAny.length);
  console.log('notDriven=' + t.notDriven.length + (t.notDriven.length ? ' ' + t.notDriven.join(',') : ''));
  console.log('matrixEntries=' + t.matrixEntries);
}

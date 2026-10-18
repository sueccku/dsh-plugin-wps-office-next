// 生成 docs/FIXES.md 顶部的索引（编号 + 标题）。只重写标记之间的块，正文一个字都不动。
//
// 为什么需要它：FIXES 已经 300+ KB，HANDOFF 让新会话"读 103–115 号"，实际约 60 KB。有了索引，
// 新会话先看标题定位、再点开需要的那几条；索引由本脚本生成，gates 只读核对，不会漂。
//
// 用法：
//   node scripts/gen-fixes-index.mjs          # 写入索引块
//   node scripts/gen-fixes-index.mjs --check  # 只比对，不写盘（gates 用）
import { readFileSync, writeFileSync } from 'node:fs';

const FILE = 'docs/FIXES.md';
const BEGIN = '<!-- FIXES-INDEX:BEGIN（由 scripts/gen-fixes-index.mjs 生成，勿手改） -->';
const END = '<!-- FIXES-INDEX:END -->';
const CHECK = process.argv.includes('--check');

function buildBlock(body) {
  const entries = [...body.matchAll(/^#{2,3} (\d+)\. (.+)$/gm)].map((m) => ({ n: Number(m[1]), title: m[2].trim() }));
  entries.sort((a, b) => a.n - b.n);
  const seen = new Set();
  const rows = [];
  for (const entry of entries) {
    if (seen.has(entry.n)) continue;
    seen.add(entry.n);
    rows.push('| ' + entry.n + ' | ' + entry.title.replace(/\|/g, '\\|') + ' |');
  }
  return [
    BEGIN,
    '## 索引（编号 + 标题）',
    '',
    '> 由 `node scripts/gen-fixes-index.mjs` 生成；`npm run gates:*` 会核对它是否最新。',
    '> 版本归属见 [release-records.md](release-records.md)，细节点开对应条目。',
    '',
    '| # | 标题 |',
    '| --- | --- |',
    ...rows,
    END,
  ].join('\n');
}

const original = readFileSync(FILE, 'utf8');
const beginAt = original.indexOf(BEGIN);
const endAt = original.indexOf(END);
let body;
if (beginAt >= 0 && endAt > beginAt) {
  body = original.slice(0, beginAt) + '__FIXES_INDEX__' + original.slice(endAt + END.length);
} else {
  const anchor = original.indexOf('\n## 已修复');
  if (anchor < 0) {
    console.error('gen-fixes-index: 找不到插入点 "\\n## 已修复"');
    process.exit(2);
  }
  body = original.slice(0, anchor) + '\n__FIXES_INDEX__\n' + original.slice(anchor);
}
const expected = body.replace('__FIXES_INDEX__', buildBlock(body));

if (CHECK) {
  if (expected !== original) {
    console.error('fixes-index: docs/FIXES.md 的索引不是最新 —— 跑 `node scripts/gen-fixes-index.mjs`');
    process.exit(1);
  }
  console.log('fixes-index: up to date');
} else {
  writeFileSync(FILE, expected, 'utf8');
  console.log('fixes-index: written ("## 索引（编号 + 标题）")');
}

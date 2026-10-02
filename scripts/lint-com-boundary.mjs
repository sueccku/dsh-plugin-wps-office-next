/**
 * COM 边界 lint —— 从「列清单」升级成真门禁（FIXES 74）。
 *
 * 背景：这个仓库踩过三类根因，都能被机器判定，但过去完全靠新写一个真机场景测试才发现：
 *   1) `return $range` 被 PowerShell 枚举展开成数组（FIXES 43：create_pivot_table 从来没成功过）；
 *   2) 裸 `catch { continue }` 把整段逻辑失效伪装成空结果（FIXES 38：多格查找静默 0 命中）；
 *   3) 枚举常量凭记忆写（FIXES 39）—— 这条无法静态判定，仍靠实测，见 docs/FIXES.md。
 *
 * 规则：
 *   A  Range 一律写 `return ,$range`（变量形式与直接表达式形式都查）；
 *   B  裸 `catch { continue }` 必须登记并给理由（账本式：涨即失败，缩要显式改）；
 *   C  （信息）仍在边界层外的 COM 成员读写 —— 只报告，不判红：那是逐步迁移的存量。
 *
 * Run: node scripts/lint-com-boundary.mjs
 */
import { readFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

export const BRIDGE = 'mcp/scripts/wps-com.ps1';

// B 的账本：键是「规范化后的源码行」，值必须写清为什么静默是安全的。
export const BARE_CONTINUE_ALLOWLIST = [
  {
    line: 'try { $pt = $sheet.PivotTables($i) } catch { continue }',
    reason: '枚举候选透视表：越界的 PivotTables($i) 会抛，continue 是循环控制流，不是吞掉真实失败',
  },
];

const RANGE_ASSIGN = /=\s*[^=].*\.(Range|Cells|Rows|Columns|Areas)\s*\(/;
const RANGE_RETURN_EXPR = /^\s*return\s+(?!,)[^\n]*\.(Range|Cells|Rows|Columns|Areas)\b/;

export function collectBoundaryViolations(text, file = BRIDGE) {
  const violations = [];
  const lines = text.split(/\r?\n/);
  const allow = new Map(BARE_CONTINUE_ALLOWLIST.map((e) => [e.line, e]));
  const seenBare = new Map();

  let func = '';
  const rangeVars = new Map();
  lines.forEach((line, index) => {
    const fn = /^function\s+([A-Za-z0-9_-]+)/.exec(line);
    if (fn) func = fn[1];
    if (!rangeVars.has(func)) rangeVars.set(func, new Set());
    const vars = rangeVars.get(func);

    const assign = /^\s*\$([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$/.exec(line);
    if (assign && RANGE_ASSIGN.test(line) && !/^\s*\$\w+\s*=\s*,/.test(line)) vars.add(assign[1]);

    const ret = /^\s*return\s+\$([A-Za-z_][A-Za-z0-9_]*)\s*$/.exec(line);
    if (ret && vars.has(ret[1])) {
      violations.push({ rule: 'range-return', file, line: index + 1, detail: 'Range 必须写成 `return ,$' + ret[1] + '`，否则会被 PowerShell 展开成数组：' + line.trim().slice(0, 70) });
    }
    if (RANGE_RETURN_EXPR.test(line)) {
      violations.push({ rule: 'range-return', file, line: index + 1, detail: 'Range 必须写成 `return ,<表达式>`：' + line.trim().slice(0, 70) });
    }

    if (/catch\s*\{\s*continue\s*\}/.test(line)) {
      const key = line.trim().replace(/\s+/g, ' ');
      seenBare.set(key, (seenBare.get(key) || 0) + 1);
      if (!allow.has(key)) {
        violations.push({ rule: 'bare-continue', file, line: index + 1, detail: '裸 `catch { continue }` 必须登记并给理由（见 BARE_CONTINUE_ALLOWLIST）：' + key.slice(0, 70) });
      }
    }
  });

  for (const entry of BARE_CONTINUE_ALLOWLIST) {
    if (!seenBare.has(entry.line)) {
      violations.push({ rule: 'bare-continue-stale', file, line: 1, detail: '账本里的这条已经不存在了，删掉：' + entry.line.slice(0, 60) });
    }
  }
  return violations;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const text = readFileSync(BRIDGE, 'utf8');
  const violations = collectBoundaryViolations(text);

  // 信息性统计：仍在边界层外的 COM 成员读写（逐步迁移的存量，不判红）。
  const patterns = [
    { name: 'value-write', re: /\.(Value2|Formula|NumberFormat|Text)\s*=(?!=)/, risk: 'HIGH - 值类型会变' },
    { name: 'com-call', re: /\$[A-Za-z_][A-Za-z0-9_.]*\.(Replace|Find|FindNext)\s*\(/, risk: 'HIGH - 参数类型会变' },
  ];
  for (const { name, re, risk } of patterns) {
    let n = 0;
    for (const line of text.split(/\r?\n/)) { if (!line.trim().startsWith('#') && re.test(line)) n++; }
    console.log('info  ' + name.padEnd(12) + String(n).padStart(4) + '   ' + risk);
  }

  for (const v of violations) console.log('LINT ' + v.rule + '  ' + v.file + ':' + v.line + '  ' + v.detail);
  console.log('boundary gate: ' + violations.length + ' violation(s)');
  process.exit(violations.length === 0 ? 0 : 1);
}

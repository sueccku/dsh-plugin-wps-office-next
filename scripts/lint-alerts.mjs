// 弹出式对话框的静态门禁（FIXES 70）。
//
// 为什么需要它：WPS 的模态框弹在 WPS 进程里，常驻宿主是单线程 STA，一个框就能把整个会话钉死
// （宿主被超时杀掉、框还在，之后每次调用只能短超时失败）。S1/S3 是「按动作清单」加固的，所以
// SaveAs / Close / 刷新 这些同源站点成片漏网。这里把「哪些调用必须先关 DisplayAlerts」变成
// 机器可判定的不变量：新加调用点不写守卫就红，写了守卫漏还原也红。
//
// 规则
//   一 覆盖：包含「可能弹框的调用」的 action 必须有 Set-WpsAlertsSuppressed；
//   二 配对：同一 action 里 Set-WpsAlertsSuppressed 与 Restore-WpsAlerts 数量必须相等；
//   三 不绕过：DisplayAlerts 的直接读写只允许出现在两个共用助手里（含注释行不算）；
//   四 例外有账：ALERTS_ACTION_EXCEPTIONS 里的每一项都必须对得上真实 action，过期即失败。
//
// 注意：`.Open(` 不在这条门禁里 —— 密码框不吃 DisplayAlerts，它由文件头预检
// （Test-WpsOoxmlEncrypted，FIXES 69）挡，断言在 test/encrypted-preflight.test.mjs。
// Run: node scripts/lint-alerts.mjs
import { readFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

export const BRIDGE_PATH = 'mcp/scripts/wps-com.ps1';

// 每条都是「没关弹窗就可能弹框」的调用，理由写在这里，免得后人再靠记忆。
const RISKY = [
  { re: /\.SaveAs\(/, why: 'SaveAs 到已存在的文件会弹「是否替换」' },
  { re: /\.Save\(\)/, why: '未落盘的文档 Save 会弹「另存为」' },
  { re: /\.Close\(/, why: '脏文档 / 只读文档 Close 会弹保存框' },
  { re: /\.ExportAsFixedFormat\(/, why: '导出到同名文件会弹「是否替换」' },
  { re: /\.InsertFromFile\(/, why: '加密来源会弹密码框' },
  { re: /\.UpdateLink\(/, why: '外部链接失效会弹提示' },
  { re: /\.RefreshAll\(/, why: '带参数的查询会弹参数框' },
  { re: /\.PrintOut\(/, why: '没有打印机时会弹框' },
  { re: /\.Quit\(/, why: '有未保存内容时会弹保存框' },
];

// 例外必须写明理由；空数组就是「一条例外都没有」——这是当前的期望值。
export const ALERTS_ACTION_EXCEPTIONS = [];

function functionBody(lines, name) {
  const start = lines.findIndex((l) => new RegExp('^function ' + name + '\\(').test(l));
  if (start < 0) return null;
  let depth = 0;
  for (let i = start; i < lines.length; i++) {
    for (const ch of lines[i]) {
      if (ch === '{') depth++;
      else if (ch === '}') depth--;
    }
    if (depth === 0 && i > start) return { start, end: i };
  }
  return { start, end: lines.length - 1 };
}

export function collectAlertsViolations(text, file = BRIDGE_PATH) {
  const violations = [];
  const lines = text.split(/\r?\n/);
  const actionRe = /^    "([A-Za-z0-9_.]+)" \{$/;
  const starts = [];
  lines.forEach((line, i) => { const m = actionRe.exec(line); if (m) starts.push({ action: m[1], line: i }); });

  starts.forEach((s, k) => {
    const end = k + 1 < starts.length ? starts[k + 1].line : lines.length;
    const body = lines.slice(s.line, end).join('\n');
    const hit = RISKY.filter((r) => r.re.test(body));
    if (!hit.length) return;
    if (ALERTS_ACTION_EXCEPTIONS.some((e) => e.action === s.action)) return;
    const sets = (body.match(/Set-WpsAlertsSuppressed/g) || []).length;
    const restores = (body.match(/Restore-WpsAlerts/g) || []).length;
    const call = hit.map((h) => h.re.source.replace(/\\/g, '')).join(' / ');
    if (sets === 0) {
      violations.push({ rule: 'alerts-missing', file, line: s.line + 1, detail: 'action "' + s.action + '" 调用了 ' + call + ' 却没有 Set-WpsAlertsSuppressed：' + hit[0].why });
    } else if (sets !== restores) {
      violations.push({ rule: 'alerts-unbalanced', file, line: s.line + 1, detail: 'action "' + s.action + '" 关了 ' + sets + ' 次弹窗却还原 ' + restores + ' 次' });
    }
  });

  // 三：DisplayAlerts 的直接读写只属于两个共用助手；注释行不算。
  const allowed = [functionBody(lines, 'Set-WpsAlertsSuppressed'), functionBody(lines, 'Restore-WpsAlerts')].filter(Boolean);
  const inHelper = (i) => allowed.some((r) => i >= r.start && i <= r.end);
  lines.forEach((line, i) => {
    if (!/\.DisplayAlerts/.test(line)) return;
    if (line.trim().startsWith('#')) return;
    if (inHelper(i)) return;
    violations.push({ rule: 'alerts-bypassed', file, line: i + 1, detail: '绕过共用助手直接读写 DisplayAlerts：' + line.trim().slice(0, 70) });
  });

  // 四：例外必须对得上真实 action。
  const names = new Set(starts.map((s) => s.action));
  for (const e of ALERTS_ACTION_EXCEPTIONS) {
    if (!names.has(e.action)) violations.push({ rule: 'alerts-stale-exception', file, line: 1, detail: '例外 "' + e.action + '" 已经没有对应 action，删掉它' });
  }

  return violations;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const text = readFileSync(BRIDGE_PATH, 'utf8');
  const violations = collectAlertsViolations(text);
  for (const v of violations) console.log('LINT ' + v.rule + '  ' + v.file + ':' + v.line + '  ' + v.detail);
  console.log('alerts gate: ' + violations.length + ' violation(s)');
  process.exit(violations.length === 0 ? 0 : 1);
}

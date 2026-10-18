// 项目化 lint：只查「这个仓库真的在乎、而且机器能一眼判定」的事情，不做通用风格警察。
// 规则来源都是本项目踩过的坑（BOM 被编辑器吃掉、CRLF 漂移、测试退出码不传播、遗留 console.log）。
// Run: node scripts/lint.mjs [--quiet]
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { collectAlertsViolations } from './lint-alerts.mjs';
import { collectBoundaryViolations } from './lint-com-boundary.mjs';

const ROOT = process.cwd();
const QUIET = process.argv.includes('--quiet');
const violations = [];
function bad(rule, file, detail) { violations.push({ rule, file: relative(ROOT, file).replace(/\\/g, '/'), detail }); }

function walk(dir, out = []) {
  for (const entry of readdirSync(dir)) {
    if (entry === 'node_modules' || entry === '.git' || entry === 'dist' || entry === '.artifacts') continue;
    const full = join(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) walk(full, out);
    // test/summary.json 是跑测试才产生的机器快照（不入库），不必遵守文本约定
    else if (!full.endsWith('summary.json')) out.push(full);
  }
  return out;
}

function textFiles() {
  const dirs = ['mcp/src', 'scripts', 'test', 'skills', 'docs', 'baseline'];
  const files = [];
  for (const d of dirs) {
    const full = join(ROOT, d);
    try { if (statSync(full).isDirectory()) files.push(...walk(full)); } catch { }
  }
  for (const f of ['plugin.js', 'cordis.patch.yml', 'README.md', 'CHANGELOG.md', 'AGENTS.md', 'THIRD_PARTY_NOTICES.md', 'package.json']) {
    const full = join(ROOT, f);
    try { if (statSync(full).isFile()) files.push(full); } catch { }
  }
  return files.filter((f) => /\.(ts|mjs|js|ps1|json|md|yml|yaml)$/.test(f));
}

const files = textFiles();

// 1) 手写 PowerShell 的字节约定（生成物 host/wps-actions.ps1 由 CI 对账，不在此列）
// FIXES 92（P8）：scripts/run-tests.ps1 原先**含中文却不带 BOM**，而 Windows PowerShell 5.1 会把无 BOM 的
// 文件按 ANSI/GBK 读 —— 实测中文注释把紧随其后的那条语句（$summaryRows = @()）整个吞掉：于是第一次
// `$null += [ordered]@{...}` 把变量变成了字典，第二次 `+=` 就抛「Item has already been added」，
// test/summary.json 的 perFile 只剩第一个文件那一行（48 行丢 47 行），而 totals 是另一条累加所以看着正常。
// 补上 BOM 后立刻恢复。教训：**含非 ASCII 的手写 .ps1 必须带 BOM**，这份清单不能只列两个文件。
// 每条 = [路径, 是否要 BOM, 是否要 CRLF]。\`scripts/*.ps1\` 按本目录既有约定是 **LF**
// （\`build-host-actions.ps1\` 也是 LF，且 \`.gitattributes\` 写着 \`* -text\`「never rewrite line endings」），
// 所以对它**只强制 BOM**，不动行尾；host/ 与 mcp/scripts/ 下的是 CRLF。
for (const [rel, wantBom, wantCrlf] of [['mcp/scripts/wps-com.ps1', false, true], ['host/wps-com-host.ps1', true, true], ['scripts/run-tests.ps1', true, false], ['scripts/reap-wps.ps1', false, false]]) {
  const buf = readFileSync(join(ROOT, rel));
  const hasBom = buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf;
  if (hasBom !== wantBom) bad('ps1-bom', join(ROOT, rel), 'BOM=' + hasBom + ' 期望 ' + wantBom);
  if (wantCrlf && /[^\r]\n/.test(buf.toString('utf8'))) bad('ps1-crlf', join(ROOT, rel), '存在裸 LF');
}

// 2) 文本文件里不许有制表符、不许行尾空白、必须以换行结尾
for (const file of files) {
  const text = readFileSync(file, 'utf8');
  if (file.endsWith('.md')) continue; // Markdown 允许行尾两个空格表示换行
  const lines = text.split(/\r?\n/);
  lines.forEach((line, i) => {
    if (line.includes('\t')) bad('no-tabs', file, '第 ' + (i + 1) + ' 行有制表符');
    if (/[ \t]+$/.test(line)) bad('no-trailing-space', file, '第 ' + (i + 1) + ' 行行尾有空白');
  });
  if (text.length > 0 && !text.endsWith('\n')) bad('final-newline', file, '文件没有以换行结尾');
}

// 3) mcp/src 里不许留 console.*（要走 logger）
for (const file of walk(join(ROOT, 'mcp/src'))) {
  if (!file.endsWith('.ts')) continue;
  readFileSync(file, 'utf8').split(/\r?\n/).forEach((line, i) => {
    if (/^\s*console\.(log|info|warn|error|debug)\(/.test(line)) bad('no-console', file, '第 ' + (i + 1) + ' 行: ' + line.trim().slice(0, 60));
  });
}

// 4) 每个测试文件必须自己决定退出码，否则失败不会传播。
//    用共用 harness 的 finish() 也算 —— 它内部就是 process.exit（见 test/lib/harness.mjs），
//    而且这样新测试就不必再抄一遍 results/收尾（P4）。
for (const name of readdirSync(join(ROOT, 'test'))) {
  if (!name.endsWith('.test.mjs')) continue;
  const src = readFileSync(join(ROOT, 'test', name), 'utf8');
  if (!/process\.exit\(/.test(src) && !/from ['"]\.\/lib\/harness\.mjs['"]/.test(src)) {
    bad('test-exit-code', join(ROOT, 'test', name), '既没有 process.exit(...) 也没有用 harness 的 finish()');
  }
}

// 5) 弹出式对话框的静态门禁（FIXES 70）：可能弹框的调用必须先关 DisplayAlerts 并在 finally 还原，
//    且不许绕过两个共用助手直接读写 DisplayAlerts。规则与理由见 scripts/lint-alerts.mjs。
const bridge = join(ROOT, 'mcp/scripts/wps-com.ps1');
for (const v of collectAlertsViolations(readFileSync(bridge, 'utf8'), bridge)) bad(v.rule, v.file, v.detail);

// 6) COM 边界门禁（FIXES 74）：`return $range` 必须 `return ,$range`；裸 `catch { continue }` 必须登记理由。
//    规则与理由见 scripts/lint-com-boundary.mjs —— 它过去只是一份没人跑的清单。
for (const v of collectBoundaryViolations(readFileSync(bridge, 'utf8'), bridge)) bad(v.rule, v.file, v.detail);

// 7) Markdown 相对链接必须指向存在的文件/目录。
//    2026-10-18 文档清理暴露的问题：悬空链接只能靠人肉找（那次清完还留了一条历史链接）。
//    历史遗留、故意保留的断链登记在这里；新增断链必须修掉或改成纯文字。
const ALLOWED_BROKEN_LINKS = new Set([
  'docs/FIXES.md -> bug-hunt-orders.md', // 2026-10-07 文档清理：按“不改历史”保留，文件头已注明
]);
function stripFencedCode(text) {
  let fence = null;
  return text.split(/\r?\n/).map((line) => {
    const m = /^\s*(`{3,}|~{3,})/.exec(line);
    if (m) {
      if (!fence) fence = m[1][0];
      else if (m[1][0] === fence) fence = null;
      return '';
    }
    return fence ? '' : line;
  }).join('\n');
}
function decodeTarget(raw) { try { return decodeURIComponent(raw); } catch { return raw; } }
for (const file of files.filter((f) => f.endsWith('.md'))) {
  const rel = relative(ROOT, file).replace(/\\/g, '/');
  const body = stripFencedCode(readFileSync(file, 'utf8'));
  for (const m of body.matchAll(/\]\(([^)\s]+)\)/g)) {
    const raw = m[1];
    if (/^(?:https?:|mailto:|data:|#)/i.test(raw)) continue;
    const target = decodeTarget(raw.split('#')[0].split('?')[0]);
    if (!target || target.startsWith('/') || /^[A-Za-z]:[\\/]/.test(target)) continue;
    if (existsSync(resolve(dirname(file), target))) continue;
    const key = rel + ' -> ' + target;
    if (!ALLOWED_BROKEN_LINKS.has(key)) bad('md-link', file, '链接目标不存在: ' + raw);
  }
}

if (!QUIET) {
  for (const v of violations.slice(0, 40)) console.log('LINT ' + v.rule + '  ' + v.file + '  ' + v.detail);
  if (violations.length > 40) console.log('... 还有 ' + (violations.length - 40) + ' 条');
}
const byRule = {};
for (const v of violations) byRule[v.rule] = (byRule[v.rule] || 0) + 1;
console.log('scanned=' + files.length + ' files, violations=' + violations.length + (Object.keys(byRule).length ? '  ' + JSON.stringify(byRule) : ''));
process.exit(violations.length === 0 ? 0 : 1);

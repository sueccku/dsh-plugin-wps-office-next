// 项目化 lint：只查「这个仓库真的在乎、而且机器能一眼判定」的事情，不做通用风格警察。
// 规则来源都是本项目踩过的坑（BOM 被编辑器吃掉、CRLF 漂移、测试退出码不传播、遗留 console.log）。
// Run: node scripts/lint.mjs [--quiet]
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';

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
    else out.push(full);
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
  for (const f of ['plugin.js', 'cordis.patch.yml', 'README.md', 'CHANGELOG.md', 'package.json']) {
    const full = join(ROOT, f);
    try { if (statSync(full).isFile()) files.push(full); } catch { }
  }
  return files.filter((f) => /\.(ts|mjs|js|ps1|json|md|yml|yaml)$/.test(f));
}

const files = textFiles();

// 1) 手写 PowerShell 的字节约定（生成物 host/wps-actions.ps1 由 CI 对账，不在此列）
for (const entry of [['mcp/scripts/wps-com.ps1', false], ['host/wps-com-host.ps1', true]]) {
  const rel = entry[0]; const wantBom = entry[1];
  const buf = readFileSync(join(ROOT, rel));
  const hasBom = buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf;
  if (hasBom !== wantBom) bad('ps1-bom', join(ROOT, rel), 'BOM=' + hasBom + ' 期望 ' + wantBom);
  if (/[^\r]\n/.test(buf.toString('utf8'))) bad('ps1-crlf', join(ROOT, rel), '存在裸 LF');
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

// 4) 每个测试文件必须自己决定退出码，否则失败不会传播
for (const name of readdirSync(join(ROOT, 'test'))) {
  if (!name.endsWith('.test.mjs')) continue;
  if (!/process\.exit\(/.test(readFileSync(join(ROOT, 'test', name), 'utf8'))) bad('test-exit-code', join(ROOT, 'test', name), '没有 process.exit(...)');
}

if (!QUIET) {
  for (const v of violations.slice(0, 40)) console.log('LINT ' + v.rule + '  ' + v.file + '  ' + v.detail);
  if (violations.length > 40) console.log('... 还有 ' + (violations.length - 40) + ' 条');
}
const byRule = {};
for (const v of violations) byRule[v.rule] = (byRule[v.rule] || 0) + 1;
console.log('scanned=' + files.length + ' files, violations=' + violations.length + (Object.keys(byRule).length ? '  ' + JSON.stringify(byRule) : ''));
process.exit(violations.length === 0 ? 0 : 1);

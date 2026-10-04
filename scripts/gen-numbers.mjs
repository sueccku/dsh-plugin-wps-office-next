const CHECK = process.argv.includes('--check');
const argv = process.argv;
/**
 * Input : spec/*.json + mcp/dist/spec/*.js + host/wps-actions.ps1 + test/summary.json
 * Output: docs/current-numbers.md（数字的唯一权威快照，已入库）+ test/numbers-report.json
 * Pos   : 门禁层。文档里的数字以前靠手改，历史上漂了至少 5 次（918/928/949/982/1072…）。
 *         这里把能静态算出来的全部算出来（广告面、schema 字节、注册工具、桥 action、别名债务、
 *         未工具化 action、测试文件数），把必须跑测试才知道的从 test/summary.json 读，
 *         再逐项核对 README / HANDOFF 的声明值 —— 对不上就非零退出（CI 用 --check）。
 *         一旦我被修改，请更新我的头部注释与 docs/HANDOFF.md 的契约管线那一节。
 */
import { readFileSync, writeFileSync, existsSync, readdirSync } from 'node:fs';
import path from 'node:path';

const ROOT = process.cwd();
const read = (p) => readFileSync(path.join(ROOT, p), 'utf8');
const readJson = (p) => JSON.parse(read(p));

const advertised = readJson('spec/advertised.json');
const definitions = readJson('spec/tool-definitions.json');
const actionKeys = readJson('spec/action-keys.json');
const signatures = readJson('spec/signatures.json');

// schema 字节：从**真实 tools/list** 量（模型实际收到的载荷），不是从 spec 推
const { measureTiers } = await import('./lib/tool-face.mjs');
const tiers = await measureTiers();
if (!tiers.standard || tiers.standard.error) {
  console.error('无法测量 tools/list 载荷：' + JSON.stringify(tiers.standard));
  process.exit(2);
}
const schemaBytes = tiers.standard.bytes;
const approxTokensMeasured = tiers.standard.approxTokens;

// 桥 action：从生成的宿主键表读，而不是从 spec 读 —— 键表才是「桥真的认这些动作」的证据
const host = read('host/wps-actions.ps1');
const keysStart = host.indexOf('$script:ActionParamKeys');
// 用带 `$script:` 前缀的**赋值行**做终点：只找 'ActionParamAliases' 会先撞上注释里的提及。
const keysSeg = host.slice(keysStart, host.indexOf('$script:ActionParamAliases = @{', keysStart));
const bridgeActions = new Set();
let bridgeKeys = 0;
// 注意：别名表（`'addAnimation' = @{ ... }`）就在键表下面，形状同样是 `'名字' = `，
// 所以切片终点必须落在别名表**之前**（见上面 keysSeg 的算法）。
for (const m of keysSeg.matchAll(/^\s*'([A-Za-z][A-Za-z0-9_]*)'\s*=\s*@\(([^)]*)\)/gm)) {
  bridgeActions.add(m[1]);
  bridgeKeys += [...m[2].matchAll(/'([^']+)'/g)].length;
}
const untooled = [...bridgeActions].filter((a) => !(a in actionKeys)).sort();

// 桥 action 的权威口径与 verify.mjs 一致：**桥源文件里的 case 块**（键表会漏掉那几个
// 「参数读不出来」的 action —— 它们只有别名条目，没有键表条目）。
const bridgeSrc = read('mcp/scripts/wps-com.ps1');
const caseRe = /^ {4}"([A-Za-z][A-Za-z0-9_.]*)" \{\s*$/gm;
const bridgeActionCases = (bridgeSrc.match(caseRe) || []).length;

const testFiles = readdirSync(path.join(ROOT, 'test')).filter((f) => f.endsWith('.test.mjs'));
const summaryPath = path.join(ROOT, 'test', 'summary.json');
const summary = existsSync(summaryPath) ? JSON.parse(readFileSync(summaryPath, 'utf8')) : null;

const numbers = {
  note: '由 scripts/gen-numbers.mjs 生成。文档里的数字请引用这里，不要手写（FIXES 91）。',
  generatedAt: new Date().toISOString().slice(0, 19) + 'Z',
  advertisedTools: advertised.length,
  registeredTools: definitions.length,
  schemaBytes,
  approxTokens: approxTokensMeasured,
  tiers,
  bridgeActions: bridgeActionCases,
  bridgeActionCases,
  keyTableActions: bridgeActions.size,
  bridgeKeys,
  actionKeyTables: Object.keys(actionKeys).length,
  untooledActions: untooled.length,
  untooledActionNames: untooled,
  signatures: signatures.length,
  testFiles: testFiles.length,
  testAssertions: summary ? summary.assertions : null,
  testRunAt: summary ? summary.generatedAt : null,
  testFailures: summary ? summary.fail : null,
};

// —— 声明值核对：文档里写死的数字必须与上面算出来的一致 ——
const claims = [];
const claim = (label, expected, actual) => claims.push({ label, expected, actual, ok: expected === actual });

const readme = read('README.md');
const badge = /advertised%20tools-(\d+)%20%2F%20(\d+)-blue/.exec(readme);
if (badge) {
  claim('README badge advertised', Number(badge[1]), numbers.advertisedTools);
  claim('README badge registered', Number(badge[2]), numbers.registeredTools);
}
const table = /\| \*\*standard（默认）\*\* \| \*\*(\d+)\*\* \| \*\*([\d,]+)\*\* \| \*\*≈([\d,]+)\*\* \|/.exec(readme);
if (table) {
  claim('README standard tools', Number(table[1]), numbers.advertisedTools);
  claim('README standard bytes', Number(table[2].replace(/,/g, '')), numbers.schemaBytes);
  claim('README standard tokens', Number(table[3].replace(/,/g, '')), numbers.approxTokens);
}
const minRow = /\| minimal \| (\d+) \| ([\d,]+) \| ([\d,]+) \|/.exec(readme);
if (minRow) {
  claim('README minimal tools', Number(minRow[1]), tiers.minimal.tools);
  claim('README minimal bytes', Number(minRow[2].replace(/,/g, '')), tiers.minimal.bytes);
}
const fullRow = /\| full \| (\d+) \| ([\d,]+) \| ([\d,]+) \|/.exec(readme);
if (fullRow) {
  claim('README full tools', Number(fullRow[1]), tiers.full.tools);
  claim('README full bytes', Number(fullRow[2].replace(/,/g, '')), tiers.full.bytes);
}
const detail = /每次请求只广告 (\d+) 个/.exec(readme);
if (detail) claim('README detail summary', Number(detail[1]), numbers.advertisedTools);
const reg = /注册目录 \*\*(\d+)\*\* 个工具/.exec(readme);
if (reg) claim('README registered catalogue', Number(reg[1]), numbers.registeredTools);
const cur = /广告面 (\d+) 工具 \/ ([\d,]+) 字节/.exec(readme);
if (cur) {
  claim('README current advertised', Number(cur[1]), numbers.advertisedTools);
  claim('README current bytes', Number(cur[2].replace(/,/g, '')), numbers.schemaBytes);
}
if (summary) {
  const rt = /当前数字：测试 \*\*(\d+) 项 \/ (\d+) 个文件\*\*/.exec(readme);
  if (rt) {
    claim('README test assertions', Number(rt[1]), numbers.testAssertions);
    claim('README test files', Number(rt[2]), numbers.testFiles);
  }
}

const handoff = read('docs/HANDOFF.md');
const ht = /\| 测试 \| \*\*(\d+) 断言 \/ (\d+) 个测试文件\*\*/.exec(handoff);
if (ht && summary) {
  claim('HANDOFF test assertions', Number(ht[1]), numbers.testAssertions);
  claim('HANDOFF test files', Number(ht[2]), numbers.testFiles);
}
const ha = /广告面 \*\*(\d+) 个\*\*/.exec(handoff);
if (ha) claim('HANDOFF advertised', Number(ha[1]), numbers.advertisedTools);

const failed = claims.filter((c) => !c.ok);
const lines = [];
lines.push('# 当前数字（自动生成，勿手改）');
lines.push('');
lines.push('> 由 `node scripts/gen-numbers.mjs` 生成；CI 跑 `--check` 核对文档里的声明值。');
lines.push('> 需要「跑测试才知道」的数字（断言数）来自 `test/summary.json`（由 `scripts/run-tests.ps1` 每次整轮写出，不入库）。');
lines.push('');
lines.push('| 项 | 值 | 来源 |');
lines.push('| --- | --- | --- |');
const row = (k, v, src) => lines.push('| ' + k + ' | **' + v + '** | ' + src + ' |');
row('广告面工具', numbers.advertisedTools + ' / ' + numbers.registeredTools, 'spec/advertised.json + spec/tool-definitions.json');
row('广告面 schema 字节', numbers.schemaBytes.toLocaleString('en-US') + '（≈' + numbers.approxTokens.toLocaleString('en-US') + ' tokens）', '真实 tools/list 载荷');
row('minimal 档', tiers.minimal.tools + ' 工具 / ' + tiers.minimal.bytes.toLocaleString('en-US') + ' 字节', '同上');
row('full 档', tiers.full.tools + ' 工具 / ' + tiers.full.bytes.toLocaleString('en-US') + ' 字节', '同上');
row('桥 action', String(numbers.bridgeActions), 'mcp/scripts/wps-com.ps1 的 case 块（与 verify.mjs 同口径）');
row('其中带键表', String(numbers.keyTableActions), 'host/wps-actions.ps1 键表');
row('桥参数键', String(numbers.bridgeKeys), '同上');
row('带键表的 action', String(numbers.actionKeyTables), 'spec/action-keys.json');
row('未工具化 action', numbers.untooledActions + (numbers.untooledActionNames.length ? '（' + numbers.untooledActionNames.join(' / ') + '）' : ''), '键表 − spec');
row('测试文件', String(numbers.testFiles), 'test/*.test.mjs');
// 只有真跑过整轮的机器才有 test/summary.json（它不入库）。没有就**不写这一行** ——
// 否则同一份生成物在「跑过测试的机器」和「CI」上会长得不一样，--check 会误报。
if (numbers.testAssertions !== null) {
  row('测试断言', numbers.testAssertions + '（失败 ' + numbers.testFailures + '，' + numbers.testRunAt + '）', 'test/summary.json');
}
lines.push('');
lines.push('## 声明值核对');
lines.push('');
lines.push('| 位置 | 文档里写的 | 实际 | 结果 |');
lines.push('| --- | --- | --- | --- |');
for (const c of claims) lines.push('| ' + c.label + ' | ' + c.expected + ' | ' + c.actual + ' | ' + (c.ok ? '✅' : '❌') + ' |');
lines.push('');
if (!summary) lines.push('> 本机没有 `test/summary.json`，涉及测试断言数的声明这次没有核对（跑一次整轮即可补上）。');
lines.push('');

writeFileSync(path.join(ROOT, 'docs', 'current-numbers.md'), lines.join('\n'), 'utf8');
writeFileSync(path.join(ROOT, 'test', 'numbers-report.json'), JSON.stringify({ numbers, claims }, null, 1) + '\n', 'utf8');

console.log('docs/current-numbers.md written');
console.log('  advertised=' + numbers.advertisedTools + ' registered=' + numbers.registeredTools +
  ' schemaBytes=' + numbers.schemaBytes + ' bridgeActions=' + numbers.bridgeActions +
  ' untooled=' + numbers.untooledActions + ' testFiles=' + numbers.testFiles +
  ' assertions=' + (numbers.testAssertions === null ? 'n/a' : numbers.testAssertions));
if (failed.length) {
  console.log('CLAIM MISMATCH (' + failed.length + '):');
  for (const c of failed) console.log('  x ' + c.label + ': doc=' + c.expected + ' actual=' + c.actual);
  if (CHECK) process.exit(1);
} else {
  console.log('CLAIMS OK (' + claims.length + ' checked)');
}

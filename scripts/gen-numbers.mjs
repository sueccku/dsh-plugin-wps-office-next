const CHECK = process.argv.includes('--check');
const argv = process.argv;
/**
 * Input : spec/*.json + mcp/dist/spec/*.js + host/wps-actions.ps1（+ test/summary.json，只有跑过整轮的机器才有）
 * Output: docs/current-numbers.md（**机器无关**的权威快照，已入库）+ test/numbers-report.json（本机产物，.gitignore）
 * Pos   : 门禁层。文档里的数字以前靠手改，历史上漂了至少 5 次（918/928/949/982/1072…）。
 *         这里把能静态算出来的全部算出来（广告面、schema 字节、注册工具、桥 action、别名债务、
 *         未工具化 action、测试文件数），再逐项核对 README / HANDOFF 的声明值 —— 对不上就非零退出
 *         （CI 用 --check）。
 *
 *         **入库文档必须机器无关**：凡是由 test/summary.json 算出来的东西（断言数，以及涉及它的
 *         4 条声明）一律**不写进 docs/current-numbers.md** —— 否则同一份生成物在「跑过测试的机器」
 *         与「CI」上不一样，CI 的 --check / git diff 永远红（实测：该门禁上线起连续 8 次 CI 全红）。
 *         它们只进 gitignored 的 test/numbers-report.json，并照常参与本机退出码判定。
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
  // 按**天**取，不按秒：这个文件里的每个字节都会进 CI 的生成物对账，秒级时间戳只会制造噪音。
  generatedAt: new Date().toISOString().slice(0, 10),
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
// 安装指引里的 advertisedTools（用户照着这句核对安装是否成功，必须一起对账；
// FIXES 91 发版审计发现这两处写着 69 而 CI 抓不到）。
for (const m of readme.matchAll(/advertisedTools 应为 (\d+)/g)) claim('README install guide advertisedTools', Number(m[1]), numbers.advertisedTools);
for (const m of readme.matchAll(/advertisedTools: (\d+)/g)) claim('README status sample advertisedTools', Number(m[1]), numbers.advertisedTools);
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
    // 比的是**最近一次整轮实际跑到的**文件数（summary.files），不是目录里现有的文件数：
    // 刚新增、还没被跑到的测试文件不该让 README 那句"最新一次整轮回归的读数"变成撒谎。
    claim('README test files', Number(rt[2]), summary.files);
  }
}

const handoff = read('docs/HANDOFF.md');
const ht = /\| 测试 \| \*\*(\d+) 断言 \/ (\d+) 个测试文件\*\*/.exec(handoff);
if (ht && summary) {
  claim('HANDOFF test assertions', Number(ht[1]), numbers.testAssertions);
  claim('HANDOFF test files', Number(ht[2]), summary.files);
}
const ha = /广告面 \*\*(\d+) 个\*\*/.exec(handoff);
if (ha) claim('HANDOFF advertised', Number(ha[1]), numbers.advertisedTools);

// 只有跑过测试的机器才算得出来的那几条声明：**不写进入库文档**（写了就会让文档随机器而变），
// 只进 gitignored 的 numbers-report.json，并照常参与本机的退出码判定。
const LOCAL_ONLY_CLAIMS = new Set(['README test assertions', 'README test files', 'HANDOFF test assertions', 'HANDOFF test files']);

const failed = claims.filter((c) => !c.ok);
const lines = [];
lines.push('# 当前数字（自动生成，勿手改）');
lines.push('');
lines.push('> 由 `node scripts/gen-numbers.mjs` 生成；CI 跑 `--check` 核对文档里的声明值。');
lines.push('> 本文件只放**从干净检出就能复现**的数字。需要「跑测试才知道」的断言数**不在本文件** ——');
lines.push('> 它随本机是否跑过整轮而变，只写进 `test/numbers-report.json`（本机产物，.gitignore，CI 不核对）。');
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
// 这里**刻意没有**「测试断言」行：它只有跑过整轮才算得出来，写进入库快照就等于把"本机状态"刻进
// 生成物 —— CI 上没有 test/summary.json，生成出来的文档必然与入库那份不同，--check 从上线起一直红。
lines.push('');
lines.push('> 「测试文件」数的是目录里现有的 `test/*.test.mjs`；README / HANDOFF 写的「N 个文件」是');
lines.push('> **最近一次整轮实际跑到的**数 —— 新增文件但还没跑整轮时，两者本来就该不同。');
lines.push('');
lines.push('## 声明值核对');
lines.push('');
lines.push('| 位置 | 文档里写的 | 实际 | 结果 |');
lines.push('| --- | --- | --- | --- |');
// 入库文档里只放机器无关的声明；与 summary 有关的那 4 条只在 numbers-report.json 与本机退出码里体现。
for (const c of claims) if (!LOCAL_ONLY_CLAIMS.has(c.label)) lines.push('| ' + c.label + ' | ' + c.expected + ' | ' + c.actual + ' | ' + (c.ok ? '✅' : '❌') + ' |');
lines.push('');
lines.push('> 与测试计数有关的那几条声明只在**本机存在 `test/summary.json`** 时核对，**不属于 CI 核对范围**；');
lines.push('> 现值看 `test/summary.json` / `test/numbers-report.json`（跑完整轮后按 README / HANDOFF 的口径刷新）。');
lines.push('');

// FIXES 91（发版审计发现）：`--check` 以前**照样覆写**生成物，只影响退出码 —— 于是文档把它归进
// 「只读核对」，实际却会先改写已入库的快照再返回非零。现在 --check 是**真只读**：
// 只把"会写成什么"与本机现存的比，报告有没有差异，不落盘。
const nextDoc = lines.join('\n');
const nextReport = JSON.stringify({ numbers, claims }, null, 1) + '\n';
const docPath = path.join(ROOT, 'docs', 'current-numbers.md');
const reportPath = path.join(ROOT, 'test', 'numbers-report.json');
const onDisk = (p) => { try { return readFileSync(p, 'utf8'); } catch { return null; } };
// 只有**入库快照**才参与 drift 判定。test/numbers-report.json 是 .gitignore 的本机产物：每跑一次整轮
// 它就会变，CI 上又根本不存在 —— 拿它当判据会让 --check 在干净检出上必红（这就是该 gate 长期失败的另一半原因）。
const drift = [];
if (onDisk(docPath) !== nextDoc) drift.push('docs/current-numbers.md');
const reportOnDisk = onDisk(reportPath);
if (reportOnDisk !== null && reportOnDisk !== nextReport) {
  console.log('note: test/numbers-report.json 落后于当前生成值（本机产物，不入库，不影响退出码）');
}

if (CHECK) {
  console.log('check 模式：不写文件。与磁盘比对的差异：' + (drift.length ? drift.join(', ') : '(无)'));
} else {
  writeFileSync(docPath, nextDoc, 'utf8');
  writeFileSync(reportPath, nextReport, 'utf8');
  console.log('docs/current-numbers.md written');
}
console.log('  advertised=' + numbers.advertisedTools + ' registered=' + numbers.registeredTools +
  ' schemaBytes=' + numbers.schemaBytes + ' bridgeActions=' + numbers.bridgeActions +
  ' untooled=' + numbers.untooledActions + ' testFiles=' + numbers.testFiles +
  ' assertions=' + (numbers.testAssertions === null ? 'n/a' : numbers.testAssertions));
if (failed.length) {
  console.log('CLAIM MISMATCH (' + failed.length + '):');
  for (const c of failed) console.log('  x ' + c.label + ': doc=' + c.expected + ' actual=' + c.actual);
} else {
  console.log('CLAIMS OK (' + claims.length + ' checked)');
}
if (CHECK && (failed.length || drift.length)) {
  if (drift.length && !failed.length) console.log('GENERATED FILES OUT OF DATE: ' + drift.join(', ') + '（跑不带 --check 的 gen-numbers 刷新）');
  process.exit(1);
}

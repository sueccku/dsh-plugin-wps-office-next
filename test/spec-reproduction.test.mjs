// P1 acceptance: the spec-driven generator must reproduce the pre-P1 model-facing surface.
//
// Compares spec/tool-definitions.json against the live tools/list (full mode), the advertised set
// against toolset.ts, and the generated action key table against the one embedded in the COM host.
// Equality is checked two ways: deep-equal after canonical key sorting, and identical serialized
// byte length (which is what the token budget is actually made of).
//
// Run: node test/spec-reproduction.test.mjs   (needs mcp/dist built)
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { readFileSync, readdirSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { computeCoverageTiers, TIER_LEDGER } from '../scripts/lib/coverage-tiers.mjs';

const { analyseToolSource } = await import(pathToFileURL(resolve('scripts/lib/tool-action-map.mjs')).href);
const { map: TOOL_MAP } = analyseToolSource();

const require = createRequire(import.meta.url);
const generated = JSON.parse(readFileSync('spec/tool-definitions.json', 'utf8'));
const advertisedGenerated = JSON.parse(readFileSync('spec/advertised.json', 'utf8'));
const actionKeysGenerated = JSON.parse(readFileSync('spec/action-keys.json', 'utf8'));

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  ' + detail : '')); }

function listFull() {
  return new Promise((resolvePromise) => {
    const child = spawn(process.execPath, ['mcp/dist/index.js'], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true, env: { ...process.env, WPS_OFFICE_TOOLSET: 'full' } });
    // **按字节收帧**：不要写成 `buf += d.toString()`。那种写法在 chunk 边界切开多字节 UTF-8 字符时，
    // 会把中文变成 U+FFFD（2026-10-11 实测：`指`→2 个替换字符、`明`→3 个），于是这条冻结判据开始
    // 误报"生成物与活面对不上"。定位依据：服务器 stdout 的原始字节里 `efbfbd` 出现 **0 次**，
    // 坏的是读法不是数据；同一份 mcp/dist，字节式分帧读出来 0 个替换字符。
    let buf = Buffer.alloc(0);
    const send = (o) => child.stdin.write(JSON.stringify(o) + '\n');
    child.stdout.on('data', (d) => {
      buf = Buffer.concat([buf, d]);
      let i;
      while ((i = buf.indexOf(0x0a)) >= 0) {
        const frame = buf.subarray(0, i);
        buf = buf.subarray(i + 1);
        const line = frame.toString('utf8').trim();
        if (!line) continue;
        let m; try { m = JSON.parse(line); } catch { continue; }
        if (m.id === 1) send({ jsonrpc: '2.0', method: 'notifications/initialized' });
        if (m.id === 2) { child.kill(); resolvePromise(m.result.tools); }
      }
    });
    child.stderr.on('data', () => {});
    send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'spec-repro', version: '1' } } });
    setTimeout(() => send({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }), 1500);
    setTimeout(() => resolvePromise([]), 40000);
  });
}

const canon = (value) => {
  if (Array.isArray(value)) return value.map(canon);
  if (value && typeof value === 'object') {
    const out = {};
    for (const key of Object.keys(value).sort()) out[key] = canon(value[key]);
    return out;
  }
  return value;
};
const bytesOf = (list) => list.reduce((n, t) => n + Buffer.byteLength(JSON.stringify(t), 'utf8'), 0);

const live = await listFull();
check('live surface is non-empty', live.length > 0, live.length + ' tools');

const generatedSorted = [...generated].sort((a, b) => a.name.localeCompare(b.name));
const liveSorted = [...live].sort((a, b) => a.name.localeCompare(b.name));

check('same tool count', generatedSorted.length === liveSorted.length, 'generated=' + generatedSorted.length + ' live=' + liveSorted.length);

const namesGenerated = generatedSorted.map((t) => t.name).join(',');
const namesLive = liveSorted.map((t) => t.name).join(',');
check('same tool names', namesGenerated === namesLive, namesGenerated === namesLive ? 'identical' : 'first difference at index ' + [...generatedSorted.keys()].find((i) => namesGenerated.split(',')[i] !== namesLive.split(',')[i]));

// 备注（2026-10-11）：这里曾经有一版"传输层坏点白名单"的归一化补丁。根因查清后已撤掉 ——
// 坏点来自**本文件 listFull() 旧的字符串式缓冲**（见那里的注释），服务器侧数据一直是干净的，
// 所以冻结判据保持**逐字节严格比对**，不做任何折叠。
const diff = [];
for (let i = 0; i < Math.min(generatedSorted.length, liveSorted.length); i++) {
  if (JSON.stringify(canon(generatedSorted[i])) !== JSON.stringify(canon(liveSorted[i]))) diff.push(generatedSorted[i].name);
}
check('every schema deep-equals the live one', diff.length === 0, diff.length ? diff.length + ' differ: ' + diff.slice(0, 6).join(', ') : 'all ' + generatedSorted.length + ' match');

const bytesGenerated = bytesOf(generatedSorted);
const bytesLive = bytesOf(liveSorted);
check('identical serialized byte length', bytesGenerated === bytesLive, 'generated=' + bytesGenerated + ' live=' + bytesLive);

const toolset = require(resolve('mcp/dist/server/toolset.js'));
const advertisedLive = [...toolset.STANDARD_TOOLS, ...toolset.FACADE_TOOLS].sort();
check('advertised set is reproduced', advertisedGenerated.join(',') === advertisedLive.join(','), 'generated=' + advertisedGenerated.length + ' live=' + advertisedLive.length);

const host = readFileSync('host/wps-actions.ps1', 'utf8');
const keysStart = host.indexOf('$script:ActionParamKeys');
const keysSeg = host.slice(keysStart, host.indexOf('$script:ActionParamAliases', keysStart));
const hostKeys = new Map();
for (const m of keysSeg.matchAll(/^\s*'([A-Za-z][A-Za-z0-9_]*)'\s*=\s*@\(([^)]*)\)/gm)) {
  hostKeys.set(m[1], [...m[2].matchAll(/'([^']+)'/g)].map((x) => x[1]));
}

// The classification of every parameter is asserted outright: nothing may be left unclassified,
// and every bridge parameter must land on a key the bridge actually reads.
//
// Two debts are ratchets instead: they fail when they GROW and must be updated deliberately when
// they shrink.
//   ALIAS_DEBT       parameters whose public name differs from the bridge key they land on. P1-4
//                    aligned them all (FIXES 77), so this is now a hard 0: a new tool must name its
//                    parameters exactly the way the bridge reads them.
//   UNTOOLED_ACTIONS bridge actions no operation drives yet: the P2 backlog of capability that
//                    already exists but has no tool. FIXES 80 清到 2：getActivePresentation /
//                    getActiveWorkbook —— 它们的返回其实已由 getOpenPresentations / getOpenWorkbooks
//                    的 active 标记覆盖，留作对称性缺口（Word 的 getActiveDocument 是工具化的）。
const spec = require(resolve('mcp/dist/spec/operations.js'));
// P4 的合并带走了三处改名参数（59 降到 59 以下），FIXES 77 把剩下的全部对齐到桥键，到此归零。
const ALIAS_DEBT = 0;
// P2 tooled 10 of the 21, and P3-1 tooled the four Word gaps, so the ratchet moves down with them:
// 21 -> 11 -> 7. What is left is deliberate (duplicate implementations) or belongs to P4.
const UNTOOLED_ACTIONS = 2;
// Actions the spec declares as "parameters cannot be read statically"; the generator refuses to
// skip anything that is not declared here, so this is a ledger rather than an allowance.
const NO_KEY_TABLE = Object.keys(require(resolve('mcp/dist/spec/aliases.js')).dynamicParamActions);

const keyMismatch = [];
const noTable = [];
const readButNotSent = [];
for (const [action, keys] of Object.entries(actionKeysGenerated)) {
  const accepted = hostKeys.get(action);
  if (!accepted) { noTable.push(action); continue; }
  for (const key of keys) if (!accepted.includes(key)) keyMismatch.push(action + '.' + key);
}
for (const [action, keys] of hostKeys) {
  const sent = actionKeysGenerated[action] || [];
  for (const key of keys) if (!sent.includes(key)) readButNotSent.push(action + '.' + key);
}

const unexpectedNoTable = noTable.filter((a) => !NO_KEY_TABLE.includes(a));
check('every generated action exists in the host key table', unexpectedNoTable.length === 0, unexpectedNoTable.length ? unexpectedNoTable.join(', ') : Object.keys(actionKeysGenerated).length + ' actions, ' + noTable.length + ' known exception(s)');

const specActions = new Set(Object.keys(actionKeysGenerated));
const untooled = [...hostKeys.keys()].filter((a) => !specActions.has(a));
let unresolved = 0;
let localParams = 0;
let containerParams = 0;
let aliased = 0;
for (const op of spec.operations) {
  for (const [name, p] of Object.entries(op.params)) {
    if (p.kind === 'unresolved') unresolved++;
    else if (p.kind === 'local') localParams++;
    else if (p.kind === 'container') containerParams++;
    if (p.kind !== 'local' && p.kind !== 'container' && op.aliases && op.aliases[name]) aliased++;
  }
}
check('every parameter has a declared destination', unresolved === 0, unresolved + ' unresolved; ' + localParams + " local to the handler, " + containerParams + ' flattened by a container');
check('every bridge parameter lands on a key the bridge reads', keyMismatch.length === 0, keyMismatch.length ? keyMismatch.slice(0, 6).join(', ') : aliased + ' of them differ only by name');
check('every parameter name equals the bridge key (P1-4 target reached)', aliased === 0, aliased + ' parameter(s) still differ only by name');
check('untooled-action backlog did not grow (P2 target: down to 0)', untooled.length <= UNTOOLED_ACTIONS, untooled.length + ' of ' + UNTOOLED_ACTIONS + ' recorded (getActivePresentation / getActiveWorkbook - both covered by the open* list tools and their active flag)');

// The bridge-side compatibility table is declared in the spec (mcp/src/spec/aliases.ts) and emitted as
// spec/param-aliases.json; the per-tool analysis view lives in operations.ts. They describe the same
// mapping from two angles, so they must agree.
// Two mechanisms perform a rename, and they must not be confused:
//   - the handler renames in code, so nothing is needed on the bridge side;
//   - the handler passes the public name through, and the bridge renames it via paramAliases.
// Only the second case requires a declared entry, so that is what is asserted. Declared entries whose
// public name no tool sends are legal too: they are the compatibility spellings kept on purpose.
const { paramAliases } = require(resolve('mcp/dist/spec/aliases.js'));
const missingDeclared = [];
let passedThrough = 0;
let legacyOnly = 0;
const sentToAction = new Set();
for (const op of spec.operations) {
  if (!op.action) continue;
  const sent = TOOL_MAP.get(op.tool);
  for (const key of (sent && sent.keys) || []) sentToAction.add(op.action + '.' + key);
}
for (const op of spec.operations) {
  if (!op.action || !op.aliases) continue;
  const declared = paramAliases[op.action] || {};
  const sent = new Set(((TOOL_MAP.get(op.tool) || {}).keys) || []);
  for (const [name, bridgeKey] of Object.entries(op.aliases)) {
    if (sent.has(name)) {
      passedThrough++;
      if (declared[name] !== bridgeKey) missingDeclared.push(op.action + '.' + name + ' (bridge declares ' + JSON.stringify(declared[name]) + ', needs ' + bridgeKey + ')');
    }
  }
}
for (const [action, map] of Object.entries(paramAliases)) {
  for (const name of Object.keys(map)) if (!sentToAction.has(action + '.' + name)) legacyOnly++;
}
check('every pass-through rename is declared on the bridge side', missingDeclared.length === 0, missingDeclared.length ? missingDeclared.slice(0, 5).join(', ') : passedThrough + ' pass-through renames checked, ' + legacyOnly + ' legacy spelling(s) declared');

// S4 coverage, now tiered (FIXES 74). The old check asked only "does the name appear somewhere in the
// test corpus" - comments counted, and it said nothing about how well a tool is exercised. The tiers
// come from scripts/lib/coverage-tiers.mjs so the CLI and CI cannot drift apart. Three ratchets, all of
// them deliberate: bespoke may not shrink; matrixOk and matrixAny may not grow. matrixOk used to be
// unguarded entirely - with a 72-tool margin on bespoke, a tool could slide from a real test into
// "the call succeeded" and nothing failed (fixed 2026-10-10).
const tiers = computeCoverageTiers();
check('every registered tool is driven by a test or the e2e run', tiers.notDriven.length === 0, tiers.notDriven.length ? tiers.notDriven.slice(0, 6).join(', ') : tiers.registered.length + ' driven');
check('bespoke coverage did not shrink (S4 ratchet)', tiers.bespoke.length >= TIER_LEDGER.bespoke, tiers.bespoke.length + ' tools have a non-matrix test (ratchet ' + TIER_LEDGER.bespoke + ')');
check('the "ok-only" middle tier only shrinks (S4 ratchet)', tiers.matrixOk.length <= TIER_LEDGER.matrixOk, tiers.matrixOk.length + ' tools are only exercised by a matrix "ok" row (ratchet ' + TIER_LEDGER.matrixOk + ')');
check('the weakest tier only shrinks (S4 ratchet)', tiers.matrixAny.length <= TIER_LEDGER.matrixAny, tiers.matrixAny.length + ' tools are only exercised by a matrix "any" row (ratchet ' + TIER_LEDGER.matrixAny + ')');

console.log('');
console.log('--- informational: keys the bridge reads that no tool sends: ' + readButNotSent.length);
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'SPEC REPRODUCTION OK (' + results.length + ')' : 'SPEC REPRODUCTION FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

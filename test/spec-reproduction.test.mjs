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
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

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
    let buf = '';
    const send = (o) => child.stdin.write(JSON.stringify(o) + '\n');
    child.stdout.on('data', (d) => {
      buf += d.toString();
      let i;
      while ((i = buf.indexOf('\n')) >= 0) {
        const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1);
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
//                    aligns the names (filePath -> path, marginTop -> topMargin, ...) and this
//                    number goes to 0.
//   UNTOOLED_ACTIONS bridge actions no operation drives yet: the P2 backlog of capability that
//                    already exists but has no tool.
const spec = require(resolve('mcp/dist/spec/operations.js'));
const ALIAS_DEBT = 62;
const UNTOOLED_ACTIONS = 21;
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
check('alias debt did not grow (P1-4 target: 0)', aliased <= ALIAS_DEBT, aliased + ' of ' + ALIAS_DEBT + ' recorded');
check('untooled-action backlog did not grow (P2 target: down to 0)', untooled.length <= UNTOOLED_ACTIONS, untooled.length + ' of ' + UNTOOLED_ACTIONS + ' recorded (autoFit*, named ranges, conditional formats, getComments, getBookmarks, ...)');

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

console.log('');
console.log('--- informational: keys the bridge reads that no tool sends: ' + readButNotSent.length);
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'SPEC REPRODUCTION OK (' + results.length + ')' : 'SPEC REPRODUCTION FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

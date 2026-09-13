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

// Two debts are recorded here instead of being asserted as zero, because zero is the P1-4/P2 goal
// and a red gate would just get disabled. Both are snapshots: they fail when the debt GROWS, and
// they must be updated deliberately when it shrinks.
//
// 1. RENAME_DEBT: tool parameters whose name differs from the bridge key they land on (the second
//    dialect, renamed inside the handlers today). P1-4 makes the public name equal the bridge key.
// 2. UNTOOLED_ACTIONS: bridge actions no operation drives yet — the P2 backlog of capability that
//    exists but has no tool.
// Ratchet: tightened from 57 to 27 once the bootstrap learned to read rename mappings out of the
// handler call sites (snake_case -> camelCase). What is left are synonyms (filePath vs path,
// style_name vs style): the public parameter name genuinely differs from the bridge key, which is
// exactly what P1-4 removes.
const RENAME_DEBT = 27;
const UNTOOLED_ACTIONS = 29;
const NO_KEY_TABLE = ['setCellFormat'];

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
check('rename debt did not grow (P1-4 target: 0)', keyMismatch.length <= RENAME_DEBT, keyMismatch.length + ' of ' + RENAME_DEBT + ' recorded, e.g. ' + keyMismatch.slice(0, 4).join(', '));
check('untooled-action backlog did not grow (P2 target: down to 0)', untooled.length <= UNTOOLED_ACTIONS, untooled.length + ' of ' + UNTOOLED_ACTIONS + ' recorded (autoFit*, named ranges, conditional formats, getComments, getBookmarks, ...)');

console.log('');
console.log('--- informational: keys the bridge reads that no tool sends: ' + readButNotSent.length);
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? 'SPEC REPRODUCTION OK (' + results.length + ')' : 'SPEC REPRODUCTION FAILED (' + failed + '/' + results.length + ')');
process.exit(failed === 0 ? 0 : 1);

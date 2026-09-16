// S4 tool-coverage matrix and read-only smoke.
//
// Static mode (default) needs no WPS: it classifies every registered tool, prints coverage per app and
// the list of tools no test/e2e names. That list is the S4 worklist; the ratchet itself is enforced by
// test/spec-reproduction.test.mjs.
//
// --check : exit non-zero if coverage dropped below the recorded ratchet.
// --live  : additionally start the MCP server and call every READ-ONLY tool once with placeholder
//           arguments, asserting each returns inside a timeout. Destructive actions are never called:
//           the plan's "call every tool once" cannot be taken literally without putting real data at
//           risk, so the live smoke is deliberately restricted to get/read/list/find tools.
//
// Usage: node scripts/smoke-tools.mjs [--check] [--live]
import { readFileSync, readdirSync } from 'node:fs';
import { spawn } from 'node:child_process';
import { join } from 'node:path';

const RATCHET = 161;
const READ_ONLY = /^wps_(excel|word|ppt|common)_(get|read|list|find|search|query)/;
const LIVE_TIMEOUT_MS = 20000;

const defs = JSON.parse(readFileSync('spec/tool-definitions.json', 'utf8'));
const names = defs.map((d) => d.name);
const corpus = readdirSync('test').filter((f) => f.endsWith('.test.mjs'))
  .map((f) => readFileSync(join('test', f), 'utf8')).join('\n') + '\n' + readFileSync('scripts/e2e.mjs', 'utf8');
const named = new Set([...corpus.matchAll(/\bwps_[a-z0-9_]+/g)].map((m) => m[0]));
const covered = names.filter((n) => named.has(n));
const uncovered = names.filter((n) => !named.has(n));

const appOf = (n) => (n.match(/^wps_([a-z]+)_/) || [])[1] || 'other';
const apps = new Map();
for (const n of names) {
  const a = appOf(n);
  if (!apps.has(a)) apps.set(a, { total: 0, covered: 0, uncovered: [] });
  const g = apps.get(a);
  g.total++;
  if (named.has(n)) g.covered++;
  else g.uncovered.push(n);
}

console.log('registered tools: ' + names.length + '   named by a test/e2e: ' + covered.length + '   (' + Math.round((covered.length / names.length) * 100) + '%)');
for (const a of [...apps.keys()].sort()) {
  const g = apps.get(a);
  console.log('  ' + a.padEnd(10) + String(g.covered).padStart(3) + ' / ' + String(g.total).padEnd(3) + (g.uncovered.length ? '  missing ' + g.uncovered.length : ''));
}
if (uncovered.length) {
  console.log('');
  console.log('uncovered (' + uncovered.length + ') — S4 worklist, PPT first:');
  for (const n of uncovered) console.log('  ' + n);
}

if (process.argv.includes('--check')) {
  if (covered.length < RATCHET) {
    console.error('\nS4 FAIL: coverage ' + covered.length + ' < ratchet ' + RATCHET);
    process.exit(1);
  }
  console.log('\nS4 COVERAGE OK (' + covered.length + ' >= ' + RATCHET + ')');
}

if (process.argv.includes('--live')) {
  const targets = defs.filter((d) => READ_ONLY.test(d.name));
  const child = spawn(process.execPath, ['mcp/dist/index.js'], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true, env: { ...process.env, WPS_OFFICE_TOOLSET: 'full' } });
  let buf = '';
  const pending = new Map();
  const send = (o) => child.stdin.write(JSON.stringify(o) + '\n');
  const req = (id, method, params) => new Promise((r) => { pending.set(id, r); send({ jsonrpc: '2.0', id, method, params }); });
  child.stdout.on('data', (d) => { buf += d.toString(); let i; while ((i = buf.indexOf('\n')) >= 0) { const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1); if (!line) continue; let m; try { m = JSON.parse(line); } catch { continue; } if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } } });
  child.stderr.on('data', () => {});
  await req(1, 'initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'smoke', version: '1' } });
  send({ jsonrpc: '2.0', method: 'notifications/initialized' });
  await new Promise((r) => setTimeout(r, 800));
  console.log('');
  console.log('live read-only smoke: ' + targets.length + ' tools');
  let id = 1000, returned = 0, timedOut = 0;
  for (const d of targets) {
    const res = await Promise.race([
      req(id++, 'tools/call', { name: d.name, arguments: placeholderArgs(d) }),
      new Promise((r) => setTimeout(() => r({ __timeout: true }), LIVE_TIMEOUT_MS)),
    ]);
    if (res && res.__timeout) { timedOut++; console.log('  TIMEOUT ' + d.name); }
    else returned++;
  }
  child.kill();
  console.log('live smoke: ' + returned + ' returned, ' + timedOut + ' timed out');
  if (timedOut > 0) process.exit(1);
}

function placeholderArgs(d) {
  const schema = d.inputSchema || {};
  const props = schema.properties || {};
  const args = {};
  for (const key of schema.required || []) {
    const p = props[key] || {};
    if (p.enum && p.enum.length) args[key] = p.enum[0];
    else if (p.type === 'number' || p.type === 'integer') args[key] = 1;
    else if (p.type === 'boolean') args[key] = false;
    else if (p.type === 'array') args[key] = [];
    else if (p.type === 'object') args[key] = {};
    else if (/range|cell|address|location/i.test(key)) args[key] = 'A1';
    else args[key] = 'x';
  }
  return args;
}

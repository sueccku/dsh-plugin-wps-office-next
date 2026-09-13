// Bootstrap the operation spec from today's truth (P1-1).
// Inputs: tools/list (full), the generated host key table, the generator's alias/container tables,
//         and the tool sources (to map tool -> bridge action).
// Output: mcp/src/spec/operations.ts (a declarative draft, curated afterwards).
import { spawn } from 'node:child_process';
import { readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const LIMIT = process.argv.includes('--limit') ? Number(process.argv[process.argv.indexOf('--limit') + 1]) : 0;

function listFull() {
  return new Promise((resolve) => {
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
        if (m.id === 2) { child.kill(); resolve(m.result.tools); }
      }
    });
    child.stderr.on('data', () => {});
    send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'bootstrap', version: '1' } } });
    setTimeout(() => send({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }), 1500);
    setTimeout(() => resolve([]), 40000);
  });
}

// --- host key table (the bridge truth about which keys each action reads) ---
const host = readFileSync('host/wps-actions.ps1', 'utf8');
const keysStart = host.indexOf('$script:ActionParamKeys');
const keysSeg = host.slice(keysStart, host.indexOf('$script:ActionParamAliases', keysStart));
const actionKeys = new Map();
for (const m of keysSeg.matchAll(/^\s*'([A-Za-z][A-Za-z0-9_]*)'\s*=\s*@\(([^)]*)\)/gm)) {
  actionKeys.set(m[1], [...m[2].matchAll(/'([^']+)'/g)].map((x) => x[1]));
}

// --- the two hand-written translation tables this spec is meant to replace ---
const gen = readFileSync('scripts/build-host-actions.ps1', 'utf8');
const aliasBlock = gen.slice(gen.indexOf('$paramAliases = @{'), gen.indexOf('$paramContainers = @{'));
const actionAliases = new Map();
for (const m of aliasBlock.matchAll(/'([A-Za-z][A-Za-z0-9_]*)'\s*=\s*@\{([^}]*)\}/g)) {
  const pairs = {};
  for (const p of m[2].matchAll(/'([^']+)'\s*=\s*'([^']+)'/g)) pairs[p[1]] = p[2];
  if (Object.keys(pairs).length) actionAliases.set(m[1], pairs);
}
const contBlock = gen.slice(gen.indexOf('$paramContainers = @{'), gen.indexOf('$helperKeys = @{'));
const actionContainers = new Map();
for (const m of contBlock.matchAll(/'([A-Za-z][A-Za-z0-9_]*)'\s*=\s*@\(([^)]*)\)/g)) {
  actionContainers.set(m[1], [...m[2].matchAll(/'([^']+)'/g)].map((x) => x[1]));
}

// --- tool -> bridge action, from the tool sources ---
const clientSrc = readFileSync('mcp/src/client/wps-client.ts', 'utf8');
const methodMap = new Map();
for (const m of clientSrc.matchAll(/async\s+([A-Za-z0-9_]+)\s*\([^)]*\)[\s\S]{0,900}?(?:invokeAction|executeMethod)(?:<[\s\S]{0,300}?>)?\(\s*'([A-Za-z0-9_]+)'/g)) if (!methodMap.has(m[1])) methodMap.set(m[1], m[2]);
function walk(dir, out = []) {
  for (const e of readdirSync(dir)) { const p = join(dir, e); if (statSync(p).isDirectory()) walk(p, out); else if (p.endsWith('.ts') && !p.endsWith('.d.ts')) out.push(p); }
  return out;
}
const toolAction = new Map();
for (const file of walk('mcp/src/tools')) {
  const src = readFileSync(file, 'utf8');
  const names = [...src.matchAll(/name:\s*'(wps_[a-z0-9_]+)'/g)];
  for (let i = 0; i < names.length; i++) {
    const from = names[i].index;
    const to = i + 1 < names.length ? names[i + 1].index : src.length;
    const block = src.slice(from, to);
    const found = new Set();
    for (const m of block.matchAll(/(?:executeMethod|invokeAction)(?:<[\s\S]{0,300}?>)?\(\s*'([A-Za-z0-9_]+)'/g)) found.add(m[1]);
    for (const m of block.matchAll(/wpsClient\.([A-Za-z0-9_]+)\s*\(/g)) { const a = methodMap.get(m[1]); if (a) found.add(a); }
    toolAction.set(names[i][1], [...found]);
  }
}

const toolset = await import('../mcp/dist/server/toolset.js');
const advertised = new Set([...toolset.STANDARD_TOOLS, ...toolset.FACADE_TOOLS]);

function appOf(name) {
  const m = /^wps_(excel|word|ppt|common|convert)_/.exec(name);
  if (m) return m[1] === 'convert' ? 'common' : m[1];
  return 'common';
}
function effectOf(name) {
  if (/_(delete|remove|clear|unmerge)/.test(name)) return 'delete';
  if (/_(close)/.test(name)) return 'lifecycle';
  if (/_(open|create|save|switch|convert)/.test(name)) return 'lifecycle';
  if (/_(export)/.test(name)) return 'export';
  if (/_(get|read|list|find|info|status|count|diagnose|evaluate)/.test(name)) return 'read';
  return 'write';
}
// The friendly fields cover the common case; anything richer (nested arrays, union types) is kept
// verbatim so the generator can reproduce the pre-P1 schema exactly. The count is a migration metric.
function expressible(p) {
  const keys = Object.keys(p).sort().join(',');
  if (keys === 'description,type' || keys === 'type' || keys === 'description,enum,type' || keys === 'enum,type' || keys === 'description,default,type' || keys === 'default,type') return true;
  if (keys === 'items,type' && p.type === 'array' && p.items && Object.keys(p.items).length === 1 && typeof p.items.type === 'string') return true;
  return false;
}
function toParamSpec(p) {
  const spec = {};
  if (p.type === 'array' && p.items && p.items.type === 'array') spec.type = 'array2d';
  else spec.type = p.type || 'string';
  if (p.description) spec.description = p.description;
  if (Array.isArray(p.enum) && p.enum.length) spec.enum = p.enum;
  if (p.items && spec.type === 'array') spec.items = { type: p.items.type || 'object' };
  if (Object.prototype.hasOwnProperty.call(p, 'default')) spec.default = p.default;
  if (!expressible(p)) spec.schema = p;
  return spec;
}

const tools = await listFull();
tools.sort((a, b) => (appOf(a.name) + a.name).localeCompare(appOf(b.name) + b.name));
const used = LIMIT ? tools.slice(0, LIMIT) : tools;
const stats = { tools: tools.length, emitted: used.length, mapped: 0, unmapped: 0, local: 0, withAliases: 0, withContainers: 0, missingRequired: [] };
const entries = [];
for (const tool of used) {
  const actions = toolAction.get(tool.name) || [];
  const action = actions.length === 1 ? actions[0] : null;
  if (action) stats.mapped++; else stats.unmapped++;
  if (!action && /generate_formula|proofread_basic/.test(tool.name)) stats.local++;
  const props = (tool.inputSchema && tool.inputSchema.properties) || {};
  const rawRequired = tool.inputSchema ? tool.inputSchema.required : undefined;
  const required = Array.isArray(rawRequired) ? rawRequired : [];
  const params = {};
  for (const [name, p] of Object.entries(props)) {
    const spec = toParamSpec(p);
    if (required.includes(name)) spec.required = true;
    params[name] = spec;
  }
  for (const r of required) if (!params[r]) stats.missingRequired.push(tool.name + ':' + r);
  const entry = { tool: tool.name, action, app: appOf(tool.name), summary: tool.description || '', params, effect: effectOf(tool.name), advertised: advertised.has(tool.name) };
  // verbatim, so an absent key stays absent and an empty array stays an empty array
  if (Array.isArray(rawRequired)) entry.required = rawRequired;
  entry.engine = action ? 'bridge' : (/generate_formula|proofread_basic/.test(tool.name) ? 'local' : 'opaque');
  const aliasSource = action ? actionAliases.get(action) : null;
  if (aliasSource) {
    const kept = Object.fromEntries(Object.entries(aliasSource).filter(([k]) => params[k]));
    if (Object.keys(kept).length) { entry.aliases = kept; stats.withAliases++; }
  }
  const containerSource = action ? actionContainers.get(action) : null;
  if (containerSource && containerSource.length) { entry.containers = containerSource; stats.withContainers++; }
  entries.push(entry);
}

const banner = [
  '// GENERATED DRAFT: bootstrap of the operation spec from the pre-P1 surface (scripts/extract-spec.mjs).',
  '// It is a starting point to curate, not the finished spec. Aliases and containers are carried over',
  '// from the old translation tables so the generator can reproduce the pre-P1 artifacts byte for byte',
  '// (that reproduction is the P1 acceptance test). Curation targets: drop aliases, add effect/enum/units,',
  '// split the union-parameter interfaces, and delete the opaque entries once nothing is unreadable.',
  '//',
  '// 一旦我被修改，请更新 docs/tool-roadmap.md 的 P1 状态。',
].join('\n');
const body = entries.map((e) => '  op(' + JSON.stringify(e, null, 2).split('\n').join('\n  ') + '),').join('\n');
const out = "import { op, OperationSpec } from './types';\n\n" + banner + '\n\nexport const operations: OperationSpec[] = [\n' + body + '\n];\n';
writeFileSync('mcp/src/spec/operations.ts', out, 'utf8');
console.log(JSON.stringify(stats, null, 1));
console.log('emitted ' + entries.length + ' operations, ' + out.split(/\n/).length + ' lines');
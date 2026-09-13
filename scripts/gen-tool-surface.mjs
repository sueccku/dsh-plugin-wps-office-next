// Generate the model-facing tool surface from the operation spec (P1-2).
//
// Input:  mcp/src/spec/*.ts, compiled to mcp/dist/spec.
// Output: spec/tool-definitions.json, spec/action-keys.json, spec/advertised.json, spec/signatures.json
//         These are committed artifacts: CI regenerates them and fails on any drift, and the COM host
//         generator reads action-keys.json instead of scraping the bridge with a hand-written table.
//
// Run: node scripts/gen-tool-surface.mjs
import { createRequire } from 'node:module';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

const require = createRequire(import.meta.url);
const { operations } = require(resolve('mcp/dist/spec/operations.js'));
// The bridge-side compatibility tables live in the spec now, not in the host generator.
const { paramAliases, paramContainers } = require(resolve('mcp/dist/spec/aliases.js'));

function paramSchema(p) {
  if (p.schema) return p.schema;
  const out = {};
  out.type = p.type === 'array2d' ? 'array' : p.type;
  if (p.description) out.description = p.description;
  if (p.enum) out.enum = p.enum;
  if (p.items) out.items = { type: p.items.type };
  if (Object.prototype.hasOwnProperty.call(p, 'default')) out.default = p.default;
  return out;
}

function requiredOf(op) {
  return Object.keys(op.params).filter((name) => op.params[name].required);
}

function inputSchemaOf(op) {
  const properties = {};
  for (const [name, p] of Object.entries(op.params)) properties[name] = paramSchema(p);
  // The hand-written definitions omit "required" when nothing is required, so a missing key is the
  // pre-P1 shape; emitting an empty array would change every such schema by 14 bytes.
  const schema = { type: 'object', properties };
  // Verbatim: an absent key must stay absent, an empty array must stay an empty array.
  if (op.required !== undefined) schema.required = op.required;
  else if (requiredOf(op).length) throw new Error('spec inconsistency: params marked required but no required list: ' + op.tool);
  return schema;
}

/** Compact signature such as "range*, sheet?" for listings, so a caller often needs no schema fetch. */
function signatureOf(op) {
  const parts = Object.entries(op.params).map(([name, p]) => name + (p.required ? '*' : '?'));
  return op.tool.replace(/^wps_/, '') + '(' + parts.join(', ') + ')';
}

const definitions = operations.map((op) => ({
  name: op.tool,
  description: op.summary,
  inputSchema: inputSchemaOf(op),
}));

// accepted bridge keys per action, with tool parameter names mapped onto the keys the bridge reads
const actionKeys = {};
for (const op of operations) {
  if (!op.action) continue;
  const keys = new Set(actionKeys[op.action] || []);
  for (const [name, p] of Object.entries(op.params)) {
    // A local parameter never reaches the bridge; a container parameter arrives as its container.
    if (p.kind === 'local') continue;
    if (p.kind === 'container') { keys.add(p.container || name); continue; }
    keys.add((op.aliases && op.aliases[name]) || name);
  }
  actionKeys[op.action] = [...keys].sort();
}

const advertised = operations.filter((op) => op.advertised).map((op) => op.tool).sort();
const signatures = operations.map((op) => ({ tool: op.tool, signature: signatureOf(op), effect: op.effect }));

mkdirSync('spec', { recursive: true });
const write = (name, value) => {
  writeFileSync('spec/' + name, JSON.stringify(value, null, 1) + '\n', 'utf8');
  console.log('  spec/' + name + '  (' + JSON.stringify(value).length + ' bytes)');
};
console.log('generated from ' + operations.length + ' operations:');
write('tool-definitions.json', definitions);
write('action-keys.json', actionKeys);
write('advertised.json', advertised);
write('signatures.json', signatures);
// consumed by scripts/build-host-actions.ps1 (Windows PowerShell 5.1 has ConvertFrom-Json but no YAML)
write('param-aliases.json', paramAliases);
write('param-containers.json', paramContainers);

const rawSchema = operations.reduce((n, op) => n + Object.values(op.params).filter((p) => p.schema).length, 0);
const aliased = operations.filter((op) => op.aliases).length;
const contained = operations.filter((op) => op.containers).length;
console.log('');
console.log('curation debt: raw schema fragments=' + rawSchema + ' ; tools with aliases=' + aliased + ' ; tools with containers=' + contained);
console.log('actions with a key table: ' + Object.keys(actionKeys).length);

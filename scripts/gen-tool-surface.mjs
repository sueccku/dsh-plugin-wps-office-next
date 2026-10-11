// Generate the model-facing tool surface from the operation spec (P1-2).
//
// Input:  mcp/src/spec/*.ts, compiled to mcp/dist/spec；工具自身的 description/schema 以
//         mcp/src/tools/**（编译到 mcp/dist/tools）里注册的定义为准。
// Output: spec/tool-definitions.json, spec/action-keys.json, spec/advertised.json, spec/signatures.json
//         These are committed artifacts: CI regenerates them and fails on any drift, and the COM host
//         generator reads action-keys.json instead of scraping the bridge with a hand-written table.
//
// 为什么以工具定义为准（而不是 op.summary）：`test/spec-reproduction.test.mjs` 拿 spec/tool-definitions.json
// 与**活的 tools/list** 逐字节对账，而活的 surface 就是注册表里的 ToolDefinition。以前这里用 op.summary，
// 两者一旦措辞不同（例如 `wps_ppt_set_background_image` 的描述只有工具定义里那份长的），
// 生成物就与真实面对不上——2026-10-11 加参数时就是这么暴露出来的。现在以定义为唯一来源，
// op.summary 只在注册表里找不到该工具时兜底。
//
// Run: node scripts/gen-tool-surface.mjs
import { createRequire } from 'node:module';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

const require = createRequire(import.meta.url);
const { operations } = require(resolve('mcp/dist/spec/operations.js'));
// The live registered definitions: the model-facing surface itself.
const { allTools } = require(resolve('mcp/dist/tools/index.js'));
// The bridge-side compatibility tables live in the spec now, not in the host generator.
const { paramAliases, paramContainers } = require(resolve('mcp/dist/spec/aliases.js'));
// Keys read by the shared resolvers: the host generator needs them to work out which keys an action
// really accepts (an action that hands its bare $p to one of those resolvers inherits their keys).
const { helperKeys } = require(resolve('mcp/dist/spec/bridge-helpers.js'));
// Actions whose parameter keys cannot be read statically. Declared, not silently skipped.
const { dynamicParamActions } = require(resolve('mcp/dist/spec/aliases.js'));

/** 注册表里每个工具的定义（name → ToolDefinition）。 */
const liveDefinitions = new Map();
for (const entry of allTools) {
  const def = entry && entry.definition ? entry.definition : entry;
  if (def && def.name) liveDefinitions.set(def.name, def);
}

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

// The live definition wins (that is what the model actually sees); the operation spec only supplies
// the fields the definition does not carry, and is the fallback for a tool nobody registered.
const definitions = operations.map((op) => {
  const live = liveDefinitions.get(op.tool);
  if (live) {
    return {
      name: live.name,
      description: live.description,
      inputSchema: live.inputSchema || inputSchemaOf(op),
    };
  }
  return {
    name: op.tool,
    description: op.summary,
    inputSchema: inputSchemaOf(op),
  };
});

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
write('param-helpers.json', helperKeys);
write('param-dynamic.json', dynamicParamActions);

const rawSchema = operations.reduce((n, op) => n + Object.values(op.params).filter((p) => p.schema).length, 0);
const aliased = operations.filter((op) => op.aliases).length;
const contained = operations.filter((op) => op.containers).length;
console.log('');
console.log('curation debt: raw schema fragments=' + rawSchema + ' ; tools with aliases=' + aliased + ' ; tools with containers=' + contained);
console.log('actions with a key table: ' + Object.keys(actionKeys).length);

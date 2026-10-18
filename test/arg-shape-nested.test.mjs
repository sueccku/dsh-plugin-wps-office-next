// 嵌套参数递归校验的单元测试（纯本地，不启动 MCP server、不需要 WPS）。
//
// 覆盖三件事：
//   ① 手写用例：形状 / 未知键 / 缺必填在嵌套层被逐条报出，路径带 `range.start`、`calls[0].tool` 这种写法；
//   ② 真实 schema 的**自洽扫描**：用每个工具的 inputSchema 合成一份"合法实参"，断言校验器不误报；
//      再对声明了嵌套 required / properties 的层各做一次负面改动，断言校验器确实报出对应路径；
//   ③ 标量互换与 null 的既有口径不变（42 之于 string、range:null 都放行）—— 这条最容易在递归时被破坏。
// Run: node test/arg-shape-nested.test.mjs
import { findArgumentProblem, findShapeProblem, closestParamName } from '../mcp/dist/server/arg-shape.js';
import { allTools } from '../mcp/dist/tools/index.js';

let pass = 0;
let fail = 0;
function check(name, ok, detail) {
  if (ok) { pass += 1; console.log('PASS ' + name); }
  else { fail += 1; console.log('FAIL ' + name + (detail ? '  ' + detail : '')); }
}

// ---- ① 手写用例 ---------------------------------------------------------
const dataSchema = { type: 'array', items: { type: 'array', items: { type: 'string' } } };
check('二维数组：JSON 字符串被拒', findShapeProblem(dataSchema, '[[1,2]]', 'data')?.kind === 'shape');
check('二维数组：内层标量被拒（路径 data[0]）', (() => {
  const p = findShapeProblem(dataSchema, ['a'], 'data');
  return p?.kind === 'shape' && p.path === 'data[0]' && p.expected === 'array' && p.actual === 'string';
})(), JSON.stringify(findShapeProblem(dataSchema, ['a'], 'data')));
check('二维数组：合法值不误报', findShapeProblem(dataSchema, [['a', 'b']], 'data') === null);

const callsSchema = {
  type: 'array',
  items: {
    type: 'object',
    properties: { tool: { type: 'string' }, args: { type: 'object' } },
    required: ['tool'],
  },
};
check('对象数组：calls 传对象被拒', findShapeProblem(callsSchema, {}, 'calls')?.kind === 'shape');
check('对象数组：缺嵌套必填（calls[0].tool）', (() => {
  const p = findShapeProblem(callsSchema, [{ args: {} }], 'calls');
  return p?.kind === 'missing-required' && p.path === 'calls[0].tool';
})(), JSON.stringify(findShapeProblem(callsSchema, [{ args: {} }], 'calls')));
check('对象数组：嵌套未知键（calls[0].bogus）', (() => {
  const p = findShapeProblem(callsSchema, [{ tool: 'x', bogus: 1 }], 'calls');
  return p?.kind === 'unknown' && p.path === 'calls[0].bogus' && p.accepted.join(',') === 'tool,args';
})(), JSON.stringify(findShapeProblem(callsSchema, [{ tool: 'x', bogus: 1 }], 'calls')));
check('对象数组：合法值不误报', findShapeProblem(callsSchema, [{ tool: 'x', args: {} }], 'calls') === null);

const rangeSchema = {
  type: ['object', 'string'],
  properties: { start: { type: 'number' }, end: { type: 'number' } },
  required: ['start', 'end'],
};
check('范围对象：缺 end（range.end）', (() => {
  const p = findShapeProblem(rangeSchema, { start: 0 }, 'range');
  return p?.kind === 'missing-required' && p.path === 'range.end';
})());
check('范围对象：未知键（range.unit）', findShapeProblem(rangeSchema, { start: 0, end: 5, unit: 'chars' }, 'range')?.kind === 'unknown');
check('范围对象："all" 是合法联合类型', findShapeProblem(rangeSchema, 'all', 'range') === null);

const nestedNumber = { type: 'object', properties: { left: { type: 'number' } } };
check('深层对象形状（position.left）', (() => {
  const p = findShapeProblem(nestedNumber, { left: { x: 1 } }, 'position');
  return p?.kind === 'shape' && p.path === 'position.left' && p.actual === 'object';
})());

// 既有口径：标量互换 / null 放行
check('标量互换：string 声明收到 42 不报', findShapeProblem({ type: 'string' }, 42, 'value') === null);
check('null 一律当作没传', findShapeProblem(rangeSchema, null, 'range') === null);
check('缺顶层必填（findArgumentProblem）', (() => {
  const p = findArgumentProblem({ type: 'object', properties: { a: { type: 'string' } }, required: ['a'] }, {});
  return p?.kind === 'missing-required' && p.path === 'a';
})());
check('顶层未知键仍报（findArgumentProblem）', (() => {
  const p = findArgumentProblem({ type: 'object', properties: { a: { type: 'string' } } }, { a: 'x', b: 1 });
  return p?.kind === 'unknown' && p.path === 'b';
})());
check('嵌套未知键的近似建议', closestParamName('toll', ['tool', 'args']) === 'tool');

// ---- ② 真实 schema 自洽扫描 ---------------------------------------------
function sample(schema) {
  if (!schema) return undefined;
  const types = Array.isArray(schema.type) ? schema.type : [schema.type];
  const pick = types.includes('object') ? 'object' : types.includes('array') ? 'array' : types[0];
  if (pick === 'object') {
    const obj = {};
    for (const [key, child] of Object.entries(schema.properties || {})) obj[key] = sample(child);
    return obj;
  }
  if (pick === 'array') return schema.items ? [sample(schema.items)] : [];
  if (pick === 'number') return 1;
  if (pick === 'boolean') return true;
  return 'x';
}

let toolsChecked = 0;
const falsePositives = [];
for (const tool of allTools) {
  const schema = tool.definition.inputSchema;
  const args = sample(schema);
  toolsChecked += 1;
  const problem = findArgumentProblem(schema, args);
  if (problem) falsePositives.push(tool.definition.name + ' -> ' + JSON.stringify(problem));
}
check('真实 schema：每个工具合成的合法实参都通过（' + toolsChecked + ' 个工具）', falsePositives.length === 0, falsePositives.slice(0, 5).join(' ; '));

/** 找到第一处"声明了 required 的对象层"，返回它的路径与必填键。 */
function firstNestedRequired(schema, path = '') {
  if (!schema) return null;
  if (schema.type && (Array.isArray(schema.type) ? schema.type.includes('object') : schema.type === 'object')) {
    if ((schema.required || []).length && path) return { path, keys: schema.required };
    for (const [key, child] of Object.entries(schema.properties || {})) {
      const found = firstNestedRequired(child, path ? path + '.' + key : key);
      if (found) return found;
    }
  }
  if (schema.type && (Array.isArray(schema.type) ? schema.type.includes('array') : schema.type === 'array')) {
    const found = firstNestedRequired(schema.items, path + '[0]');
    if (found) return found;
  }
  return null;
}

let negativeChecked = 0;
const negativeMisses = [];
for (const tool of allTools) {
  const schema = tool.definition.inputSchema;
  const target = firstNestedRequired(schema);
  if (!target) continue;
  const args = sample(schema);
  // 顺着路径改坏：删掉该层一个必填键
  const segments = target.path.replace(/\[0\]/g, '.0').split('.');
  let node = args;
  for (const segment of segments) {
    if (!node) { node = null; break; }
    node = Array.isArray(node) ? node[Number(segment)] : node[segment];
  }
  if (!node || typeof node !== 'object') continue;
  const dropKey = target.keys[0];
  delete node[dropKey];
  const problem = findArgumentProblem(schema, args);
  const wantPath = target.path + '.' + dropKey;
  negativeChecked += 1;
  if (!problem || problem.kind !== 'missing-required' || problem.path !== wantPath) {
    negativeMisses.push(tool.definition.name + ' 期望 ' + wantPath + ' 实际 ' + JSON.stringify(problem));
  }
}
check('真实 schema：删掉嵌套必填能被报出（' + negativeChecked + ' 个工具）', negativeChecked > 0 && negativeMisses.length === 0, negativeMisses.slice(0, 5).join(' ; '));

// 嵌套未知键：对第一个声明了 properties 的嵌套对象塞一个假键
let unknownChecked = 0;
const unknownMisses = [];
for (const tool of allTools) {
  const schema = tool.definition.inputSchema;
  const nested = Object.entries(schema.properties || {}).find(
    ([, child]) => child && child.properties && Object.keys(child.properties).length > 0
  );
  if (!nested) continue;
  const [key, child] = nested;
  const args = sample(schema);
  if (!args || typeof args[key] !== 'object' || Array.isArray(args[key])) continue;
  args[key].__bogus__ = 1;
  const problem = findArgumentProblem(schema, args);
  unknownChecked += 1;
  if (!problem || problem.kind !== 'unknown' || problem.path !== key + '.__bogus__') {
    unknownMisses.push(tool.definition.name + ' 期望 ' + key + '.__bogus__ 实际 ' + JSON.stringify(problem));
  }
}
check('真实 schema：嵌套未知键能被报出（' + unknownChecked + ' 个工具）', unknownChecked > 0 && unknownMisses.length === 0, unknownMisses.slice(0, 5).join(' ; '));

console.log(fail === 0 ? 'ARG SHAPE (NESTED) TESTS OK (' + pass + ')' : 'ARG SHAPE (NESTED) TESTS FAILED (' + fail + '/' + (pass + fail) + ')');
process.exit(fail === 0 ? 0 : 1);

/**
 * Input: Tool 的 inputSchema 与调用方实参
 * Output: 第一处结构性问题（形状不对 / 未知参数 / 缺必填），没有则 null
 * Pos: 参数结构校验（递归）。ToolRegistry 在 handler 之前调用它，挡掉"必然跑不通"的实参。
 *      只做**结构**校验，三条规则：
 *        1. 形状：容器/标量错位（data 传字符串、calls 传对象、range 传标量…）—— 标量之间一律放行
 *           （桥/COM 会把 42 和 "42" 都吞下去，按 schema 强校验只会误伤合法调用）；
 *        2. 未知键：该层声明了 properties 时，多传的键直接拒绝（handler 用显式解构拼参数，
 *           未知键到不了桥，桥的"未知键拒绝"永远看不见它 —— FIXES 99 的同一类问题，
 *           以前只挡了顶层，嵌套层仍是静默丢弃）；
 *        3. 必填：该层声明了 required 就逐项检查（以前只查顶层，`range: {start: 0}` 会退化成
 *           "作用在当前选区"，模型以为自己改了指定字符范围）。
 *      枚举值仍由桥裁决（错误措辞含 unknown xxx、未知 xxx，有测试覆盖），这里不重复校验。
 *      一旦我被修改，请更新我的头部注释。
 */
import { ToolInputSchema, ToolParameterSchema } from '../types/tools';

export type ShapeProblemKind = 'shape' | 'unknown' | 'missing-required';

export interface ShapeProblem {
  kind: ShapeProblemKind;
  /** 实参路径：顶层是 `range`，嵌套是 `range.start` / `calls[0].tool`。 */
  path: string;
  /** kind=shape 时：声明能接受哪些类型、实际收到什么。 */
  expected?: string;
  actual?: string;
  /** kind=unknown 时：这一层声明了哪些键。 */
  accepted?: string[];
}

/** 编辑距离：只用来给"未知参数"提一句"是不是想传 X"，不做任何语义判断。 */
function editDistance(a: string, b: string): number {
  const prev = new Array<number>(b.length + 1);
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    let diag = prev[0];
    prev[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const temp = prev[j];
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, diag + (a[i - 1] === b[j - 1] ? 0 : 1));
      diag = temp;
    }
  }
  return prev[b.length];
}

export function closestParamName(key: string, accepted: string[]): string | undefined {
  const norm = (s: string) => s.toLowerCase().replace(/[_-]/g, '');
  const target = norm(key);
  let best: string | undefined;
  let bestScore = Number.POSITIVE_INFINITY;
  for (const candidate of accepted) {
    const c = norm(candidate);
    if (c === target) return candidate;
    const score = editDistance(target, c);
    if (score < bestScore) {
      bestScore = score;
      best = candidate;
    }
  }
  const limit = Math.max(2, Math.floor(target.length / 3));
  return best !== undefined && bestScore <= limit ? best : undefined;
}

/** JSON 实参的运行时类型名（数组归为 array，null 由调用方先跳过）。 */
export function actualType(value: unknown): string {
  if (Array.isArray(value)) return 'array';
  if (value === null) return 'null';
  return typeof value === 'object' ? 'object' : typeof value;
}

function declaredTypes(schema: ToolParameterSchema): string[] {
  const type = (schema as { type?: unknown }).type;
  if (Array.isArray(type)) return type.map((t) => String(t));
  return type === undefined ? [] : [String(type)];
}

/** Node 里没有 JSON Schema 的标量互换语义：这里只把 string/number/boolean 视为可互转。 */
function isScalar(type: string): boolean {
  return type === 'string' || type === 'number' || type === 'boolean';
}

function childPath(base: string, key: string): string {
  return base ? base + '.' + key : key;
}

function indexPath(base: string, index: number): string {
  return base + '[' + index + ']';
}

/**
 * 递归找第一处结构问题。value 为 undefined/null 时一律当作"没传"（顶层与嵌套同一口径：
 * 顶层以前就是 `continue`，嵌套保持同样行为，避免把 `range: null` 从"忽略"改成"报错"）。
 */
export function findShapeProblem(
  schema: ToolParameterSchema,
  value: unknown,
  path: string
): ShapeProblem | null {
  if (value === undefined || value === null) return null;

  const expectedTypes = declaredTypes(schema);
  if (expectedTypes.length === 0) return null; // 没声明类型就不judge（未知键/必填仍由上层查）

  const actual = actualType(value);
  const shapeOk = expectedTypes.some((t) => t === actual || (isScalar(t) && isScalar(actual)));
  if (!shapeOk) {
    return { kind: 'shape', path, expected: expectedTypes.join(' | '), actual };
  }

  if (actual === 'object') {
    const record = value as Record<string, unknown>;
    const nestedRequired = (schema.required || []).filter((key) => record[key] === undefined || record[key] === null);
    if (nestedRequired.length > 0) {
      return { kind: 'missing-required', path: childPath(path, nestedRequired[0]) };
    }
    const nestedProperties = schema.properties;
    if (nestedProperties) {
      for (const [key, child] of Object.entries(record)) {
        if (!(key in nestedProperties)) {
          return { kind: 'unknown', path: childPath(path, key), accepted: Object.keys(nestedProperties) };
        }
        const problem = findShapeProblem(nestedProperties[key], child, childPath(path, key));
        if (problem) return problem;
      }
    }
  }

  if (actual === 'array' && schema.items) {
    const items = value as unknown[];
    for (let index = 0; index < items.length; index++) {
      const problem = findShapeProblem(schema.items, items[index], indexPath(path, index));
      if (problem) return problem;
    }
  }

  return null;
}

/** 顶层检查：必填 → 未知键 → 递归形状。返回第一处问题；null 表示结构上没问题。 */
export function findArgumentProblem(
  schema: ToolInputSchema | undefined,
  args: Record<string, unknown> | undefined
): ShapeProblem | null {
  if (!schema) return null;
  const input = args || {};
  const properties = schema.properties || {};

  for (const param of schema.required || []) {
    if (!(param in input) || input[param] === undefined || input[param] === null) {
      return { kind: 'missing-required', path: param };
    }
  }

  for (const [key, value] of Object.entries(input)) {
    if (value === undefined || value === null) continue;
    const propertySchema = properties[key];
    if (!propertySchema) {
      return { kind: 'unknown', path: key, accepted: Object.keys(properties) };
    }
    const problem = findShapeProblem(propertySchema, value, key);
    if (problem) return problem;
  }

  return null;
}

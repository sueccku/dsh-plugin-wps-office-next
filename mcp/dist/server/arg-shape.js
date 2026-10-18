"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.closestParamName = closestParamName;
exports.actualType = actualType;
exports.findShapeProblem = findShapeProblem;
exports.findArgumentProblem = findArgumentProblem;
/** 编辑距离：只用来给"未知参数"提一句"是不是想传 X"，不做任何语义判断。 */
function editDistance(a, b) {
    const prev = new Array(b.length + 1);
    for (let j = 0; j <= b.length; j++)
        prev[j] = j;
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
function closestParamName(key, accepted) {
    const norm = (s) => s.toLowerCase().replace(/[_-]/g, '');
    const target = norm(key);
    let best;
    let bestScore = Number.POSITIVE_INFINITY;
    for (const candidate of accepted) {
        const c = norm(candidate);
        if (c === target)
            return candidate;
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
function actualType(value) {
    if (Array.isArray(value))
        return 'array';
    if (value === null)
        return 'null';
    return typeof value === 'object' ? 'object' : typeof value;
}
function declaredTypes(schema) {
    const type = schema.type;
    if (Array.isArray(type))
        return type.map((t) => String(t));
    return type === undefined ? [] : [String(type)];
}
/** Node 里没有 JSON Schema 的标量互换语义：这里只把 string/number/boolean 视为可互转。 */
function isScalar(type) {
    return type === 'string' || type === 'number' || type === 'boolean';
}
function childPath(base, key) {
    return base ? base + '.' + key : key;
}
function indexPath(base, index) {
    return base + '[' + index + ']';
}
/**
 * 递归找第一处结构问题。value 为 undefined/null 时一律当作"没传"（顶层与嵌套同一口径：
 * 顶层以前就是 `continue`，嵌套保持同样行为，避免把 `range: null` 从"忽略"改成"报错"）。
 */
function findShapeProblem(schema, value, path) {
    if (value === undefined || value === null)
        return null;
    const expectedTypes = declaredTypes(schema);
    if (expectedTypes.length === 0)
        return null; // 没声明类型就不judge（未知键/必填仍由上层查）
    const actual = actualType(value);
    const shapeOk = expectedTypes.some((t) => t === actual || (isScalar(t) && isScalar(actual)));
    if (!shapeOk) {
        return { kind: 'shape', path, expected: expectedTypes.join(' | '), actual };
    }
    if (actual === 'object') {
        const record = value;
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
                if (problem)
                    return problem;
            }
        }
    }
    if (actual === 'array' && schema.items) {
        const items = value;
        for (let index = 0; index < items.length; index++) {
            const problem = findShapeProblem(schema.items, items[index], indexPath(path, index));
            if (problem)
                return problem;
        }
    }
    return null;
}
/** 顶层检查：必填 → 未知键 → 递归形状。返回第一处问题；null 表示结构上没问题。 */
function findArgumentProblem(schema, args) {
    if (!schema)
        return null;
    const input = args || {};
    const properties = schema.properties || {};
    for (const param of schema.required || []) {
        if (!(param in input) || input[param] === undefined || input[param] === null) {
            return { kind: 'missing-required', path: param };
        }
    }
    for (const [key, value] of Object.entries(input)) {
        if (value === undefined || value === null)
            continue;
        const propertySchema = properties[key];
        if (!propertySchema) {
            return { kind: 'unknown', path: key, accepted: Object.keys(properties) };
        }
        const problem = findShapeProblem(propertySchema, value, key);
        if (problem)
            return problem;
    }
    return null;
}
//# sourceMappingURL=arg-shape.js.map
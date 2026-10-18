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
export declare function closestParamName(key: string, accepted: string[]): string | undefined;
/** JSON 实参的运行时类型名（数组归为 array，null 由调用方先跳过）。 */
export declare function actualType(value: unknown): string;
/**
 * 递归找第一处结构问题。value 为 undefined/null 时一律当作"没传"（顶层与嵌套同一口径：
 * 顶层以前就是 `continue`，嵌套保持同样行为，避免把 `range: null` 从"忽略"改成"报错"）。
 */
export declare function findShapeProblem(schema: ToolParameterSchema, value: unknown, path: string): ShapeProblem | null;
/** 顶层检查：必填 → 未知键 → 递归形状。返回第一处问题；null 表示结构上没问题。 */
export declare function findArgumentProblem(schema: ToolInputSchema | undefined, args: Record<string, unknown> | undefined): ShapeProblem | null;
//# sourceMappingURL=arg-shape.d.ts.map
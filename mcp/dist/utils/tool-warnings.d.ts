/** 在本次工具调用的作用域里执行 handler，并把收集到的注意事项一并返回。 */
export declare function runWithWarningCollector<T>(fn: () => Promise<T>): Promise<{
    value: T;
    warnings: string[];
    shortfalls: string[];
}>;
/**
 * 把桥侧结果里的 shortfalls 收进当前调用（R11）。
 * 与 `collectToolWarnings` 是**两个不同性质的通道**：
 *   warnings   = 你没要求的坏消息（探测/读取/清理失败），请求的效果仍然达成了；
 *   shortfalls = 你要求的事没（完全）做到 → 注册表会把 success 改判为 false。
 */
export declare function collectToolShortfalls(value: unknown): void;
/** 把桥侧结果里的 warnings 收进当前调用；不在工具调用作用域里（例如启动自检）就静默忽略。 */
export declare function collectToolWarnings(value: unknown): void;
//# sourceMappingURL=tool-warnings.d.ts.map
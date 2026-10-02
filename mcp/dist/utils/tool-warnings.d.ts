/** 在本次工具调用的作用域里执行 handler，并把收集到的注意事项一并返回。 */
export declare function runWithWarningCollector<T>(fn: () => Promise<T>): Promise<{
    value: T;
    warnings: string[];
}>;
/** 把桥侧结果里的 warnings 收进当前调用；不在工具调用作用域里（例如启动自检）就静默忽略。 */
export declare function collectToolWarnings(value: unknown): void;
//# sourceMappingURL=tool-warnings.d.ts.map
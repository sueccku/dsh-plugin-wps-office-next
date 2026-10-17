export type FaultMode = 'timeout' | 'shortfall';
export interface FaultRule {
    /** 原始片段，仅用于日志与错误提示。 */
    raw: string;
    pattern: RegExp;
    mode: FaultMode;
    remaining: number;
}
export interface ParsedFaultSpec {
    rules: FaultRule[];
    /** 语法不对的片段：不静默忽略，交给调用方打日志（typo 会让注入不按预期发生）。 */
    invalid: string[];
}
/**
 * 解析规格。任何一段写坏都不抛异常：测试注入写错时宁可少注入（并留下 invalid），
 * 也不该让 MCP server 起不来 —— 起不来会把"实验条件错了"表现成"工具全坏了"。
 */
export declare function parseFaultSpec(raw: string | undefined | null): ParsedFaultSpec;
export declare class FaultInjector {
    private readonly rules;
    readonly invalid: string[];
    constructor(raw?: string | undefined | null);
    get active(): boolean;
    /** 命中则消耗一次并返回模式；没有命中返回 null。 */
    take(toolName: string, args?: Record<string, unknown>): FaultMode | null;
}
/** 注入超时时回给模型的文案：与 com-host 真实超时同一套说法（「状态未知」是判断依据）。 */
export declare function injectedTimeoutText(actionLabel: string): string;
/** 注入部分失败时回给模型的短欠项：写成"这一步没生效"，而不是含糊的"部分失败"。 */
export declare function injectedShortfallText(actionLabel: string): string;
//# sourceMappingURL=fault-inject.d.ts.map
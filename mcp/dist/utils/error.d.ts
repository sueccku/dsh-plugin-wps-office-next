/**
 * Input: 错误信息与错误码
 * Output: 标准化错误对象
 * Pos: MCP 错误处理模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 错误处理工具 - 老王的错误处理系统
 * 报错见了这个模块都得绕道走，艹
 */
/**
 * 错误码枚举 - 各种SB错误都有对应的码
 */
export declare enum ErrorCode {
    UNKNOWN = 1000,
    INVALID_PARAMS = 1001,
    TIMEOUT = 1002,
    INTERNAL_ERROR = 1003,
    WPS_CONNECTION_FAILED = 2001,
    WPS_NOT_RUNNING = 2002,
    WPS_API_ERROR = 2003,
    WPS_TIMEOUT = 2004,
    TOOL_NOT_FOUND = 3001,
    TOOL_EXECUTION_FAILED = 3002,
    TOOL_INVALID_ARGS = 3003,
    TOOL_ALREADY_REGISTERED = 3004,
    MCP_INVALID_REQUEST = 4001,
    MCP_METHOD_NOT_FOUND = 4002,
    MCP_PARSE_ERROR = 4003
}
/**
 * 自定义错误基类 - 所有错误的祖宗
 */
export declare class McpError extends Error {
    readonly code: ErrorCode;
    readonly details?: Record<string, unknown>;
    readonly timestamp: Date;
    constructor(message: string, code?: ErrorCode, details?: Record<string, unknown>);
    /**
     * 转换为JSON格式 - 方便传输
     */
    toJSON(): Record<string, unknown>;
}
/**
 * WPS连接错误 - 连不上WPS就用这个骂
 */
export declare class WpsConnectionError extends McpError {
    constructor(message: string, details?: Record<string, unknown>);
}
/**
 * WPS API错误 - WPS API返回的SB错误
 */
export declare class WpsApiError extends McpError {
    constructor(message: string, details?: Record<string, unknown>);
}
/**
 * Tool不存在错误 - 找不到Tool就骂这个
 */
export declare class ToolNotFoundError extends McpError {
    constructor(toolName: string);
}
/**
 * Tool执行错误 - Tool跑出问题了
 */
export declare class ToolExecutionError extends McpError {
    constructor(toolName: string, originalError: Error, details?: Record<string, unknown>);
}
/**
 * 参数验证错误 - 参数传错了就骂这个
 */
export declare class InvalidParamsError extends McpError {
    constructor(message: string, details?: Record<string, unknown>);
}
/**
 * 超时错误 - 等太久了，不等了
 */
export declare class TimeoutError extends McpError {
    constructor(operation: string, timeoutMs: number);
}
/**
 * 错误处理工具函数
 */
export declare const errorUtils: {
    /**
     * 包装错误 - 把各种SB错误统一包装成McpError
     */
    wrap(error: unknown, defaultMessage?: string): McpError;
    /**
     * 记录并重新抛出错误
     */
    logAndThrow(error: unknown, context?: string): never;
    /**
     * 安全执行函数 - 出错了返回默认值，不会炸
     */
    safeExecute<T>(fn: () => Promise<T>, defaultValue: T, context?: string): Promise<T>;
    /**
     * 判断是否是特定类型的错误
     */
    isErrorCode(error: unknown, code: ErrorCode): boolean;
};
/**
 * 格式化错误信息给用户看
 */
export declare const formatErrorForUser: (error: McpError) => string;
export default McpError;
//# sourceMappingURL=error.d.ts.map
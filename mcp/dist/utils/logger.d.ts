/**
 * Input: 日志消息与上下文
 * Output: 结构化日志输出
 * Pos: MCP 日志系统模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 日志工具 - 老王的日志系统
 * 写日志比写代码还规范，报错了至少知道是哪个SB搞出来的
 */
import winston from 'winston';
/**
 * 日志级别枚举
 */
export declare enum LogLevel {
    ERROR = "error",
    WARN = "warn",
    INFO = "info",
    DEBUG = "debug"
}
/**
 * 主Logger - MCP Server专用
 */
export declare const logger: winston.Logger;
/**
 * 创建子Logger - 各模块用自己的
 */
export declare const createChildLogger: (moduleName: string) => winston.Logger;
/**
 * 快捷日志函数 - 懒人专用
 */
export declare const log: {
    /**
     * 记录信息日志
     */
    info: (message: string, meta?: Record<string, unknown>) => void;
    /**
     * 记录警告日志
     */
    warn: (message: string, meta?: Record<string, unknown>) => void;
    /**
     * 记录错误日志 - 艹，又出错了
     */
    error: (message: string, error?: Error | unknown, meta?: Record<string, unknown>) => void;
    /**
     * 记录调试日志 - 排查问题用的
     */
    debug: (message: string, meta?: Record<string, unknown>) => void;
};
/**
 * 请求日志中间件风格的函数
 */
export declare const logRequest: (method: string, params?: Record<string, unknown>) => void;
/**
 * 响应日志
 */
export declare const logResponse: (method: string, success: boolean, duration: number, data?: unknown) => void;
export default logger;
//# sourceMappingURL=logger.d.ts.map
"use strict";
/**
 * Input: 日志消息与上下文
 * Output: 结构化日志输出
 * Pos: MCP 日志系统模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 日志工具 - 老王的日志系统
 * 写日志比写代码还规范，报错了至少知道是哪个SB搞出来的
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.logResponse = exports.logRequest = exports.log = exports.createChildLogger = exports.logger = exports.LogLevel = void 0;
const winston_1 = __importDefault(require("winston"));
const path_1 = __importDefault(require("path"));
/**
 * 日志级别枚举
 */
var LogLevel;
(function (LogLevel) {
    LogLevel["ERROR"] = "error";
    LogLevel["WARN"] = "warn";
    LogLevel["INFO"] = "info";
    LogLevel["DEBUG"] = "debug";
})(LogLevel || (exports.LogLevel = LogLevel = {}));
/**
 * 自定义日志格式 - 老王风格
 */
const wangFormat = winston_1.default.format.printf(({ level, message, timestamp, ...meta }) => {
    const metaStr = Object.keys(meta).length ? ` | ${JSON.stringify(meta)}` : '';
    return `[${timestamp}] [${level.toUpperCase()}] ${message}${metaStr}`;
});
/**
 * 解析布尔环境变量
 */
const isEnvTrue = (value) => {
    if (!value)
        return false;
    return ['1', 'true', 'yes', 'on'].includes(value.toLowerCase());
};
/**
 * 创建Logger实例
 */
const createLogger = (name) => {
    const homeDir = process.env.HOME || process.env.USERPROFILE || require('os').homedir();
    const logDir = path_1.default.join(homeDir, '.wps-office-mcp', 'logs');
    // MCP 默认走 stdio 协议，stdout 必须保持纯净；仅在显式开启时才输出 Console 日志
    const enableConsoleLog = isEnvTrue(process.env.MCP_CONSOLE_LOG);
    const transports = [
        // 文件输出 - 错误单独记录，出问题好找
        new winston_1.default.transports.File({
            filename: path_1.default.join(logDir, 'error.log'),
            level: 'error',
        }),
        // 所有日志
        new winston_1.default.transports.File({
            filename: path_1.default.join(logDir, 'combined.log'),
        }),
    ];
    if (enableConsoleLog) {
        transports.unshift(new winston_1.default.transports.Console({
            format: winston_1.default.format.combine(winston_1.default.format.colorize({ all: true }), winston_1.default.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }), wangFormat),
        }));
    }
    return winston_1.default.createLogger({
        level: process.env.LOG_LEVEL || 'info',
        format: winston_1.default.format.combine(winston_1.default.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }), winston_1.default.format.errors({ stack: true }), wangFormat),
        defaultMeta: { service: name },
        transports,
    });
};
/**
 * 主Logger - MCP Server专用
 */
exports.logger = createLogger('wps-mcp');
/**
 * 创建子Logger - 各模块用自己的
 */
const createChildLogger = (moduleName) => {
    return exports.logger.child({ module: moduleName });
};
exports.createChildLogger = createChildLogger;
/**
 * 快捷日志函数 - 懒人专用
 */
exports.log = {
    /**
     * 记录信息日志
     */
    info: (message, meta) => {
        exports.logger.info(message, meta);
    },
    /**
     * 记录警告日志
     */
    warn: (message, meta) => {
        exports.logger.warn(message, meta);
    },
    /**
     * 记录错误日志 - 艹，又出错了
     */
    error: (message, error, meta) => {
        if (error instanceof Error) {
            exports.logger.error(message, { ...meta, error: error.message, stack: error.stack });
        }
        else {
            exports.logger.error(message, { ...meta, error });
        }
    },
    /**
     * 记录调试日志 - 排查问题用的
     */
    debug: (message, meta) => {
        exports.logger.debug(message, meta);
    },
};
/**
 * 请求日志中间件风格的函数
 */
const logRequest = (method, params) => {
    exports.logger.info(`[REQUEST] ${method}`, { params });
};
exports.logRequest = logRequest;
/**
 * 响应日志
 */
const logResponse = (method, success, duration, data) => {
    const level = success ? 'info' : 'error';
    exports.logger.log(level, `[RESPONSE] ${method} - ${success ? 'SUCCESS' : 'FAILED'} (${duration}ms)`, {
        success,
        duration,
        ...(process.env.LOG_LEVEL === 'debug' ? { data } : {}),
    });
};
exports.logResponse = logResponse;
exports.default = exports.logger;
//# sourceMappingURL=logger.js.map
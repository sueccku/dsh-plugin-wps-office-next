"use strict";
/**
 * Input: 错误信息与错误码
 * Output: 标准化错误对象
 * Pos: MCP 错误处理模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 错误处理工具 - 老王的错误处理系统
 * 报错见了这个模块都得绕道走，艹
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.formatErrorForUser = exports.errorUtils = exports.TimeoutError = exports.InvalidParamsError = exports.ToolExecutionError = exports.ToolNotFoundError = exports.WpsApiError = exports.WpsConnectionError = exports.McpError = exports.ErrorCode = void 0;
const logger_1 = require("./logger");
/**
 * 错误码枚举 - 各种SB错误都有对应的码
 */
var ErrorCode;
(function (ErrorCode) {
    // 通用错误 1xxx
    ErrorCode[ErrorCode["UNKNOWN"] = 1000] = "UNKNOWN";
    ErrorCode[ErrorCode["INVALID_PARAMS"] = 1001] = "INVALID_PARAMS";
    ErrorCode[ErrorCode["TIMEOUT"] = 1002] = "TIMEOUT";
    ErrorCode[ErrorCode["INTERNAL_ERROR"] = 1003] = "INTERNAL_ERROR";
    // WPS连接错误 2xxx
    ErrorCode[ErrorCode["WPS_CONNECTION_FAILED"] = 2001] = "WPS_CONNECTION_FAILED";
    ErrorCode[ErrorCode["WPS_NOT_RUNNING"] = 2002] = "WPS_NOT_RUNNING";
    ErrorCode[ErrorCode["WPS_API_ERROR"] = 2003] = "WPS_API_ERROR";
    ErrorCode[ErrorCode["WPS_TIMEOUT"] = 2004] = "WPS_TIMEOUT";
    // Tool错误 3xxx
    ErrorCode[ErrorCode["TOOL_NOT_FOUND"] = 3001] = "TOOL_NOT_FOUND";
    ErrorCode[ErrorCode["TOOL_EXECUTION_FAILED"] = 3002] = "TOOL_EXECUTION_FAILED";
    ErrorCode[ErrorCode["TOOL_INVALID_ARGS"] = 3003] = "TOOL_INVALID_ARGS";
    ErrorCode[ErrorCode["TOOL_ALREADY_REGISTERED"] = 3004] = "TOOL_ALREADY_REGISTERED";
    // MCP协议错误 4xxx
    ErrorCode[ErrorCode["MCP_INVALID_REQUEST"] = 4001] = "MCP_INVALID_REQUEST";
    ErrorCode[ErrorCode["MCP_METHOD_NOT_FOUND"] = 4002] = "MCP_METHOD_NOT_FOUND";
    ErrorCode[ErrorCode["MCP_PARSE_ERROR"] = 4003] = "MCP_PARSE_ERROR";
})(ErrorCode || (exports.ErrorCode = ErrorCode = {}));
/**
 * 自定义错误基类 - 所有错误的祖宗
 */
class McpError extends Error {
    code;
    details;
    timestamp;
    constructor(message, code = ErrorCode.UNKNOWN, details) {
        super(message);
        this.name = 'McpError';
        this.code = code;
        this.details = details;
        this.timestamp = new Date();
        // 保持正确的原型链
        Object.setPrototypeOf(this, McpError.prototype);
    }
    /**
     * 转换为JSON格式 - 方便传输
     */
    toJSON() {
        return {
            name: this.name,
            message: this.message,
            code: this.code,
            details: this.details,
            timestamp: this.timestamp.toISOString(),
            stack: this.stack,
        };
    }
}
exports.McpError = McpError;
/**
 * WPS连接错误 - 连不上WPS就用这个骂
 */
class WpsConnectionError extends McpError {
    constructor(message, details) {
        super(message, ErrorCode.WPS_CONNECTION_FAILED, details);
        this.name = 'WpsConnectionError';
        Object.setPrototypeOf(this, WpsConnectionError.prototype);
    }
}
exports.WpsConnectionError = WpsConnectionError;
/**
 * WPS API错误 - WPS API返回的SB错误
 */
class WpsApiError extends McpError {
    constructor(message, details) {
        super(message, ErrorCode.WPS_API_ERROR, details);
        this.name = 'WpsApiError';
        Object.setPrototypeOf(this, WpsApiError.prototype);
    }
}
exports.WpsApiError = WpsApiError;
/**
 * Tool不存在错误 - 找不到Tool就骂这个
 */
class ToolNotFoundError extends McpError {
    constructor(toolName) {
        super(`Tool not found: ${toolName}`, ErrorCode.TOOL_NOT_FOUND, { toolName });
        this.name = 'ToolNotFoundError';
        Object.setPrototypeOf(this, ToolNotFoundError.prototype);
    }
}
exports.ToolNotFoundError = ToolNotFoundError;
/**
 * Tool执行错误 - Tool跑出问题了
 */
class ToolExecutionError extends McpError {
    constructor(toolName, originalError, details) {
        super(`Tool execution failed: ${toolName} - ${originalError.message}`, ErrorCode.TOOL_EXECUTION_FAILED, { toolName, originalError: originalError.message, ...details });
        this.name = 'ToolExecutionError';
        Object.setPrototypeOf(this, ToolExecutionError.prototype);
    }
}
exports.ToolExecutionError = ToolExecutionError;
/**
 * 参数验证错误 - 参数传错了就骂这个
 */
class InvalidParamsError extends McpError {
    constructor(message, details) {
        super(message, ErrorCode.INVALID_PARAMS, details);
        this.name = 'InvalidParamsError';
        Object.setPrototypeOf(this, InvalidParamsError.prototype);
    }
}
exports.InvalidParamsError = InvalidParamsError;
/**
 * 超时错误 - 等太久了，不等了
 */
class TimeoutError extends McpError {
    constructor(operation, timeoutMs) {
        super(`Operation timed out: ${operation} (${timeoutMs}ms)`, ErrorCode.TIMEOUT, { operation, timeoutMs });
        this.name = 'TimeoutError';
        Object.setPrototypeOf(this, TimeoutError.prototype);
    }
}
exports.TimeoutError = TimeoutError;
/**
 * 错误处理工具函数
 */
exports.errorUtils = {
    /**
     * 包装错误 - 把各种SB错误统一包装成McpError
     */
    wrap(error, defaultMessage = 'An error occurred') {
        if (error instanceof McpError) {
            return error;
        }
        if (error instanceof Error) {
            return new McpError(error.message, ErrorCode.INTERNAL_ERROR, {
                originalName: error.name,
                stack: error.stack,
            });
        }
        return new McpError(typeof error === 'string' ? error : defaultMessage, ErrorCode.UNKNOWN, { originalError: error });
    },
    /**
     * 记录并重新抛出错误
     */
    logAndThrow(error, context) {
        const mcpError = exports.errorUtils.wrap(error);
        logger_1.log.error(context ? `${context}: ${mcpError.message}` : mcpError.message, mcpError);
        throw mcpError;
    },
    /**
     * 安全执行函数 - 出错了返回默认值，不会炸
     */
    async safeExecute(fn, defaultValue, context) {
        try {
            return await fn();
        }
        catch (error) {
            const mcpError = exports.errorUtils.wrap(error);
            logger_1.log.error(context ? `${context}: ${mcpError.message}` : mcpError.message, mcpError);
            return defaultValue;
        }
    },
    /**
     * 判断是否是特定类型的错误
     */
    isErrorCode(error, code) {
        return error instanceof McpError && error.code === code;
    },
};
/**
 * 格式化错误信息给用户看
 */
const formatErrorForUser = (error) => {
    switch (error.code) {
        case ErrorCode.WPS_CONNECTION_FAILED:
        case ErrorCode.WPS_NOT_RUNNING:
            return '无法连接到WPS Office，请确保WPS已启动并且加载项已安装';
        case ErrorCode.WPS_TIMEOUT:
        case ErrorCode.TIMEOUT:
            return `操作超时，请稍后重试${error.details?.operation ? '（' + error.details.operation + '）' : ''}`;
        case ErrorCode.WPS_API_ERROR:
            return `WPS接口调用失败: ${error.message}`;
        case ErrorCode.TOOL_NOT_FOUND:
            return `找不到指定的工具: ${error.details?.toolName}`;
        case ErrorCode.TOOL_EXECUTION_FAILED:
            return `工具执行失败${error.details?.toolName ? '（' + error.details.toolName + '）' : ''}: ${error.message}`;
        case ErrorCode.TOOL_INVALID_ARGS:
        case ErrorCode.INVALID_PARAMS:
            return `参数错误: ${error.message}`;
        case ErrorCode.TOOL_ALREADY_REGISTERED:
            return `工具已注册，请勿重复注册: ${error.details?.toolName}`;
        case ErrorCode.MCP_INVALID_REQUEST:
            return '请求格式无效，请检查请求参数';
        case ErrorCode.MCP_METHOD_NOT_FOUND:
            return '请求的方法不存在，请检查方法名称';
        case ErrorCode.MCP_PARSE_ERROR:
            return '消息解析失败，请检查数据格式';
        case ErrorCode.INTERNAL_ERROR:
            return `内部错误: ${error.message}`;
        default:
            return `操作失败: ${error.message}`;
    }
};
exports.formatErrorForUser = formatErrorForUser;
exports.default = McpError;
//# sourceMappingURL=error.js.map
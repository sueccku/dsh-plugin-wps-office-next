"use strict";
/**
 * Input: Tool 定义与调用请求
 * Output: Tool 执行结果
 * Pos: MCP Tool 注册与调度中心。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tool注册管理 - 老王的Tool管理系统
 * 所有Tool都得在这儿注册，不注册的Tool就是野鸡Tool
 * 这个设计遵循OCP原则：扩展开放，修改关闭
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.registerTool = exports.toolRegistry = exports.ToolRegistry = void 0;
exports.RegisterTool = RegisterTool;
const uuid_1 = require("uuid");
const tools_1 = require("../types/tools");
const logger_1 = require("../utils/logger");
const error_1 = require("../utils/error");
const logger = (0, logger_1.createChildLogger)('ToolRegistry');
/**
 * Tool注册表 - 单例模式，全局唯一
 */
class ToolRegistry {
    static instance;
    tools;
    categories;
    constructor() {
        this.tools = new Map();
        this.categories = new Map();
        // 初始化分类
        Object.values(tools_1.ToolCategory).forEach((category) => {
            this.categories.set(category, new Set());
        });
        logger.info('ToolRegistry initialized');
    }
    /**
     * 获取单例实例
     */
    static getInstance() {
        if (!ToolRegistry.instance) {
            ToolRegistry.instance = new ToolRegistry();
        }
        return ToolRegistry.instance;
    }
    /**
     * 注册Tool - 把Tool加到注册表里
     * 遵循SRP：只负责注册，不负责其他
     */
    register(definition, handler) {
        const { name, category } = definition;
        // 检查是否已注册 - 跳过重复注册而非崩溃，防止过期编译产物导致启动失败
        if (this.tools.has(name)) {
            logger.warn(`Tool already registered, skipping: ${name}`);
            return;
        }
        // 注册Tool
        this.tools.set(name, { definition, handler });
        // 添加到分类
        const toolCategory = category || tools_1.ToolCategory.COMMON;
        this.categories.get(toolCategory)?.add(name);
        logger.info(`Tool registered: ${name}`, { category: toolCategory });
    }
    /**
     * 批量注册Tools - 一次注册一堆
     */
    registerAll(tools) {
        tools.forEach(({ definition, handler }) => {
            this.register(definition, handler);
        });
        logger.info(`Batch registered ${tools.length} tools`);
    }
    /**
     * 注销Tool - 把Tool从注册表里删掉
     */
    unregister(name) {
        const tool = this.tools.get(name);
        if (!tool) {
            return false;
        }
        this.tools.delete(name);
        // 从分类中移除
        const category = tool.definition.category || tools_1.ToolCategory.COMMON;
        this.categories.get(category)?.delete(name);
        logger.info(`Tool unregistered: ${name}`);
        return true;
    }
    /**
     * 获取Tool定义
     */
    getTool(name) {
        return this.tools.get(name);
    }
    /**
     * 检查Tool是否存在
     */
    hasTool(name) {
        return this.tools.has(name);
    }
    /**
     * 获取所有Tool定义 - MCP的tools/list用的
     */
    listTools() {
        const tools = Array.from(this.tools.values()).map((t) => t.definition);
        return { tools };
    }
    /**
     * 按分类获取Tools
     */
    getToolsByCategory(category) {
        const toolNames = this.categories.get(category);
        if (!toolNames) {
            return [];
        }
        return Array.from(toolNames)
            .map((name) => this.tools.get(name)?.definition)
            .filter((t) => t !== undefined);
    }
    /**
     * 调用Tool - 执行Tool的handler
     * 这是核心方法，处理Tool调用请求
     */
    async callTool(request) {
        const { id, name, arguments: args } = request;
        const startTime = Date.now();
        logger.debug(`Calling tool: ${name}`, { id, args });
        // 检查Tool是否存在
        const tool = this.tools.get(name);
        if (!tool) {
            throw new error_1.ToolNotFoundError(name);
        }
        try {
            // 验证参数
            this.validateArguments(tool.definition, args);
            // 执行handler
            const result = await tool.handler(args);
            const duration = Date.now() - startTime;
            logger.info(`Tool executed: ${name}`, { id, duration, success: result.success });
            return {
                ...result,
                id,
            };
        }
        catch (error) {
            const duration = Date.now() - startTime;
            if (error instanceof error_1.McpError) {
                logger.error(`Tool execution failed: ${name}`, error, { id, duration });
                return {
                    id,
                    success: false,
                    content: [{ type: 'text', text: error.message }],
                    error: error.message,
                };
            }
            const execError = new error_1.ToolExecutionError(name, error instanceof Error ? error : new Error(String(error)));
            logger.error(`Tool execution failed: ${name}`, execError, { id, duration });
            return {
                id,
                success: false,
                content: [{ type: 'text', text: execError.message }],
                error: execError.message,
            };
        }
    }
    /**
     * 验证参数 - 检查必填参数是否都有
     */
    validateArguments(definition, args) {
        const { inputSchema } = definition;
        const required = inputSchema.required || [];
        for (const param of required) {
            if (!(param in args) || args[param] === undefined || args[param] === null) {
                throw new error_1.InvalidParamsError(`Missing required parameter: ${param}`, {
                    toolName: definition.name,
                    missingParam: param,
                });
            }
        }
        // TODO: 可以加更详细的类型验证，但现在先这样，够用
    }
    /**
     * 获取Tool数量
     */
    get size() {
        return this.tools.size;
    }
    /**
     * 清空所有Tools - 测试用，生产环境别乱用
     */
    clear() {
        this.tools.clear();
        this.categories.forEach((set) => set.clear());
        logger.warn('All tools cleared');
    }
    /**
     * 创建Tool调用请求 - 辅助方法
     */
    static createRequest(name, args) {
        return {
            id: (0, uuid_1.v4)(),
            name,
            arguments: args,
        };
    }
}
exports.ToolRegistry = ToolRegistry;
// 导出单例
exports.toolRegistry = ToolRegistry.getInstance();
/**
 * 装饰器：注册Tool
 * 用法：@RegisterTool(definition)
 */
function RegisterTool(definition) {
    return function (_target, _propertyKey, descriptor) {
        const originalMethod = descriptor.value;
        exports.toolRegistry.register(definition, originalMethod);
        return descriptor;
    };
}
/**
 * 快捷注册函数
 */
const registerTool = (definition, handler) => {
    exports.toolRegistry.register(definition, handler);
};
exports.registerTool = registerTool;
exports.default = ToolRegistry;
//# sourceMappingURL=tool-registry.js.map
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
const tool_warnings_1 = require("../utils/tool-warnings");
const fault_inject_1 = require("../utils/fault-inject");
const arg_shape_1 = require("./arg-shape");
const error_1 = require("../utils/error");
const logger = (0, logger_1.createChildLogger)('ToolRegistry');
/**
 * Tool注册表 - 单例模式，全局唯一
 */
class ToolRegistry {
    static instance;
    tools;
    categories;
    /** 仅测试用的受控故障注入；WPS_OFFICE_FAULT_INJECT 未设置时 active=false，什么都不做。 */
    faults = new fault_inject_1.FaultInjector();
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
            // 仅测试用的受控故障注入（WPS_OFFICE_FAULT_INJECT 未设置时 take() 恒为 null）。
            // shortfall 在**执行前**判定：这一步真的不执行，"未生效"与文件状态一致；
            // timeout 在**执行后**改判：动作已落盘，模型看到的是"状态未知"，这正是要测的场景。
            const fault = this.faults.take(name, args);
            if (fault === 'shortfall') {
                const label = typeof args?.tool === 'string' ? args.tool : name;
                const partialResult = {
                    id,
                    success: true,
                    content: [{ type: 'text', text: '（这一步没有执行：注入的部分失败）' }],
                };
                const injected = this.applyShortfalls(partialResult, [(0, fault_inject_1.injectedShortfallText)(label)], name);
                logger.info(`Fault injected (shortfall): ${name}`, { id });
                return { ...injected, id };
            }
            // 执行handler，同时收集桥侧留下的 warnings / shortfalls
            const { value: result, warnings, shortfalls } = await (0, tool_warnings_1.runWithWarningCollector)(() => tool.handler(args));
            const withWarnings = this.attachWarnings(result, warnings);
            // R11 不变量：success === true ⟺ 调用方要求的事**全部生效**
            const finalResult = shortfalls.length ? this.applyShortfalls(withWarnings, shortfalls, name) : withWarnings;
            if (fault === 'timeout') {
                // 动作已经执行完（可能已落盘），但调用方只看到"状态未知"—— 模型必须自己核对再决定是否重试。
                const label = typeof args?.tool === 'string' ? args.tool : name;
                const message = (0, fault_inject_1.injectedTimeoutText)(label);
                logger.info(`Fault injected (timeout after execution): ${name}`, { id });
                return { id, success: false, content: [{ type: 'text', text: message }], error: message };
            }
            const duration = Date.now() - startTime;
            logger.info(`Tool executed: ${name}`, { id, duration, success: finalResult.success, warnings: warnings.length, shortfalls: shortfalls.length });
            return {
                ...finalResult,
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
     * R11：桥侧用 `Add-WpsShortfall` 标记「**要求的事没做到**」。这里统一改判为失败 ——
     * **不变量：`success === true` ⟺ 调用方要求的事全部生效**。
     *
     * 内容一个字不丢：已生效的部分照旧返回，短欠清单追加在正文末尾，并置 `partial` / `shortfalls` 供程序化消费。
     * 这条**替换**了旧政策（D16）的「部分成功也回 `success: true`，只用 warnings 提示」——
     * warnings 是咨询性的，只检查 `success` 的调用方会被误导，而「假成功」正是本项目最贵的缺陷类别。
     */
    applyShortfalls(result, shortfalls, toolName) {
        const headline = '部分未生效（动作：' +
            toolName +
            '）' +
            shortfalls.join('；') +
            '下一步：按上面的条目修正参数或前置条件后重试；已经生效的部分见正文。';
        const note = '\n\n未生效（' + shortfalls.length + ' 条，其余已生效）：\n' + shortfalls.map((s) => '- ' + s).join('\n');
        const body = (result.content ?? [])
            .filter((b) => b.type === 'text' && typeof b.text === 'string')
            .map((b) => b.text)
            .join('\n')
            .trim();
        // 直通工具（wps_call / wps_execute_method / wps_batch）的正文就是桥的**原始 JSON**，
        // 契约是「原样透传」——那种情况下正文里已经带着 shortfalls 字段，**一个字都不能改**
        // （改了会让调用方的 JSON.parse 失败）。只改判结论，不动内容。
        const alreadyVerbatim = shortfalls.every((s) => body.includes(s));
        return {
            ...result,
            success: false,
            error: headline,
            partial: true,
            shortfalls,
            content: alreadyVerbatim ? result.content : [{ type: 'text', text: (body || headline) + note }],
        };
    }
    /**
     * 把「尽力而为的失败」附在结果文本后面。第一方 handler 会丢掉桥侧的 warnings，这里统一补上；
     * 已经出现在文本里的（wps_call / wps_execute_method / wps_batch 的原样透传）不重复追加（FIXES 73）。
     */
    attachWarnings(result, warnings) {
        if (!warnings.length)
            return result;
        const missing = warnings.filter((w) => !result.content.some((block) => block.type === 'text' && typeof block.text === 'string' && block.text.includes(w)));
        if (!missing.length)
            return result;
        const note = '\n\n注意（' + missing.length + ' 条，主操作已完成，但其中某些步骤没有成功）：\n' +
            missing.map((w) => '- ' + w).join('\n');
        if (!result.content.length) {
            return { ...result, content: [{ type: 'text', text: note.trim() }] };
        }
        return {
            ...result,
            content: result.content.map((block, index) => index === 0 && block.type === 'text' ? { ...block, text: (block.text ?? '') + note } : block),
        };
    }
    /**
     * 验证参数 - 必填 / 未知键 / 形状（**递归**，见 ./arg-shape）。
     *
     * 结构性问题必须在 handler 之前挡掉，理由见 arg-shape.ts 顶部：handler 用显式解构拼参数，
     * 未知键与形状错位到不了桥，桥的拒绝永远看不见它们。
     * 只挡结构：标量之间放行（桥/COM 会互转），枚举仍由桥裁决。
     */
    validateArguments(definition, args) {
        const problem = (0, arg_shape_1.findArgumentProblem)(definition.inputSchema, args);
        if (!problem)
            return;
        if (problem.kind === 'missing-required') {
            throw new error_1.InvalidParamsError(`Missing required parameter: ${problem.path}`, {
                toolName: definition.name,
                missingParam: problem.path,
            });
        }
        if (problem.kind === 'unknown') {
            const accepted = problem.accepted || [];
            const leaf = problem.path.split('.').pop() || problem.path;
            const hint = (0, arg_shape_1.closestParamName)(leaf, accepted);
            throw new error_1.InvalidParamsError(`未知参数: ${problem.path}（动作：${definition.name}）` +
                `可选参数: ${accepted.length ? accepted.join(' / ') : '(无)'}` +
                (hint ? `。是不是想传 ${hint}？` : '。') +
                `下一步：改成上面列出的参数名重试；若这个参数确实用不上，把它从调用里删掉。`, { toolName: definition.name, unknownParam: problem.path });
        }
        throw new error_1.InvalidParamsError(`参数 ${problem.path} 的形状不对：需要 ${problem.expected}，实际收到 ${problem.actual}`, { toolName: definition.name, param: problem.path, expected: problem.expected, actual: problem.actual });
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
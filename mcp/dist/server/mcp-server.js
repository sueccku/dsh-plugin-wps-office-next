"use strict";
/**
 * Input: MCP 协议请求
 * Output: Tool 调用响应
 * Pos: MCP Server 核心实现。一旦我被修改，请更新我的头部注释，并在 docs/FIXES.md 记账。
 * MCP Server实现 - 老王的MCP协议实现
 * 这是整个系统的核心，处理MCP协议通信
 * 基于stdio传输，Claude Desktop就是这么连的
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.createMcpServer = exports.WpsMcpServer = void 0;
const index_js_1 = require("@modelcontextprotocol/sdk/server/index.js");
const stdio_js_1 = require("@modelcontextprotocol/sdk/server/stdio.js");
const types_js_1 = require("@modelcontextprotocol/sdk/types.js");
const tool_registry_1 = require("./tool-registry");
const wps_client_1 = require("../client/wps-client");
const com_host_1 = require("../client/com-host");
const tools_1 = require("../types/tools");
const toolset_1 = require("./toolset");
const tools_2 = require("../tools");
const logger_1 = require("../utils/logger");
const wps_version_1 = require("../utils/wps-version");
const error_1 = require("../utils/error");
const logger = (0, logger_1.createChildLogger)('McpServer');
/**
 * 默认配置
 */
const DEFAULT_CONFIG = {
    name: 'wps-office-mcp',
    version: '1.0.0',
    debug: false,
};
/**
 * WPS MCP Server - 核心服务器类
 */
class WpsMcpServer {
    config;
    server;
    registry;
    isRunning = false;
    constructor(config) {
        this.config = { ...DEFAULT_CONFIG, ...config };
        this.registry = tool_registry_1.toolRegistry;
        // 创建MCP Server实例
        this.server = new index_js_1.Server({
            name: this.config.name,
            version: this.config.version,
        }, {
            capabilities: {
                tools: {},
            },
        });
        // 注册请求处理器
        this.setupRequestHandlers();
        // 错误处理
        this.server.onerror = (error) => {
            logger.error('MCP Server error', error);
        };
        logger.info('MCP Server created', {
            name: this.config.name,
            version: this.config.version,
        });
    }
    /**
     * 设置请求处理器
     */
    setupRequestHandlers() {
        // 处理 tools/list 请求
        this.server.setRequestHandler(types_js_1.ListToolsRequestSchema, async () => {
            logger.debug('Handling tools/list request');
            // 工具面收敛：只广告当前档位选择的工具，其余仍可通过 wps_call 调用
            const mode = (0, toolset_1.resolveMode)(process.env.WPS_OFFICE_TOOLSET);
            const { tools } = this.registry.listTools();
            const advertised = (0, toolset_1.selectTools)(mode, tools);
            logger.info('Returning ' + advertised.length + '/' + tools.length + ' tools (mode=' + mode + ')');
            return {
                tools: advertised.map((tool) => ({
                    name: tool.name,
                    description: mode === 'full' ? tool.description : (0, toolset_1.compactDescription)(tool.description),
                    inputSchema: tool.inputSchema,
                })),
            };
        });
        // 处理 tools/call 请求
        this.server.setRequestHandler(types_js_1.CallToolRequestSchema, async (request) => {
            const { name, arguments: args } = request.params;
            logger.debug('Handling tools/call request', { name, args });
            const targetName = name;
            const targetArgs = args || {};
            // 检查Tool是否存在
            if (!this.registry.hasTool(targetName)) {
                throw new types_js_1.McpError(types_js_1.ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
            }
            try {
                // 创建调用请求并执行
                const callRequest = tool_registry_1.ToolRegistry.createRequest(targetName, targetArgs);
                const result = await this.registry.callTool(callRequest);
                return {
                    content: result.content,
                    isError: !result.success,
                };
            }
            catch (error) {
                logger.error('Tool call failed', error);
                if (error instanceof error_1.McpError) {
                    return {
                        content: [{ type: 'text', text: error.message }],
                        isError: true,
                    };
                }
                throw new types_js_1.McpError(types_js_1.ErrorCode.InternalError, error instanceof Error ? error.message : 'Unknown error');
            }
        });
    }
    /**
     * 注册内置Tools - 一些基础的WPS操作Tool
     */
    /**
     * 注册逃生舱工具。原有的 11 个 builtin 已删除（重复/缓存/连接检查），
     * 只保留 wps_execute_method：覆盖尚未工具化时唯一的自逃生路径，文档里明确不推荐。
     */
    registerBuiltinTools() {
        logger.info('Registering the escape-hatch built-in tool');
        this.registry.register({
            name: 'wps_execute_method',
            // P5-1 决策：保留为**隐藏**逃生舱（不进广告面）。理由：未工具化 action 台账只剩 7 个刻意的
            // 重复实现，日常不需要它；但水印与文档属性是实测的 WPS 缺口，没有它就没有任何出路。
            // 描述里把契约说死，免得模型把它当常规手段。
            description: '最后手段：直接调用原始 WPS COM 方法。优先用现成工具（先用 wps_help 或技能参考表找）；只有在确认没有对应工具时才用它，参数与返回值都不会被校验。',
            inputSchema: {
                type: 'object',
                properties: {
                    method: {
                        type: 'string',
                        description: 'API方法名',
                    },
                    params: {
                        type: 'object',
                        description: '方法参数',
                    },
                    appType: {
                        type: 'string',
                        description: '应用类型：wps（文字）、et（表格）、wpp（演示）',
                        enum: ['wps', 'et', 'wpp'],
                    },
                },
                required: ['method'],
            },
            category: tools_1.ToolCategory.COMMON,
        }, async (args) => {
            const method = args.method;
            const params = args.params;
            const appType = args.appType;
            const response = await wps_client_1.wpsClient.executeMethod(method, params, appType);
            return {
                id: '',
                success: response.success,
                content: [
                    {
                        type: 'text',
                        text: JSON.stringify(response),
                    },
                ],
            };
        });
    }
    /**
     * 注册门面工具 - 常驻广告的四个入口
     * wps_call 让全部已注册工具保持可用，而 tools/list 只广告一小部分
     */
    registerFacadeTools() {
        const currentMode = () => (0, toolset_1.resolveMode)(process.env.WPS_OFFICE_TOOLSET);
        const text = (value) => ({
            id: '',
            success: true,
            content: [{ type: 'text', text: JSON.stringify(value) }],
        });
        const failure = (message) => ({
            id: '',
            success: false,
            content: [{ type: 'text', text: message }],
            error: message,
        });
        const callable = (name) => {
            if (!name || toolset_1.FACADE_TOOLS.includes(name))
                return false;
            return this.registry.hasTool(name);
        };
        // 状态总览
        this.registry.register({
            name: 'wps_status',
            description: '查看 WPS 连接状态、当前活动应用与所选工具面；编辑前先调用它',
            inputSchema: { type: 'object', properties: {} },
            category: tools_1.ToolCategory.COMMON,
        }, async () => {
            const started = Date.now();
            let connected = false;
            let appInfo = null;
            let note = '';
            try {
                const ping = await wps_client_1.wpsClient.executeMethod('ping');
                connected = ping.success === true;
            }
            catch (error) {
                note = error instanceof Error ? error.message : String(error);
            }
            if (connected) {
                try {
                    const info = await wps_client_1.wpsClient.executeMethod('getAppInfo');
                    appInfo = info.success ? (info.data || null) : null;
                }
                catch (error) {
                    note = error instanceof Error ? error.message : String(error);
                }
            }
            // 版本前置检查：判定逻辑是 utils/wps-version.ts 里的纯函数，CI 有单测；这里只负责读与报。
            // 必须用 fileVersion（exe 文件版本）：Application.Version 是 Office 兼容值，WPS 12.1 会报 12.0。
            const info = (appInfo || {});
            const version = (0, wps_version_1.parseWpsVersion)(info.fileVersion);
            const tooOld = (0, wps_version_1.wpsVersionTooOld)(version);
            const compatibilityWarning = (0, wps_version_1.wpsCompatibilityWarning)(info.appName, info.fileVersion);
            const all = this.registry.listTools().tools;
            const advertised = (0, toolset_1.selectTools)(currentMode(), all);
            return text({
                connected,
                appInfo,
                wpsVersion: info.fileVersion || undefined,
                officeCompatVersion: info.version || undefined,
                wpsBuild: info.build || undefined,
                wpsVersionSupported: version.length === 0 ? undefined : !tooOld,
                compatibilityWarning,
                toolset: currentMode(),
                advertisedTools: advertised.length,
                registeredTools: all.length,
                hiddenTools: all.length - advertised.length,
                note: note || undefined,
                latencyMs: Date.now() - started,
            });
        });
        // 目录与 schema 查询
        this.registry.register({
            name: 'wps_help',
            description: '查询工具目录与参数 schema：无参看分组概览，传 app 查某应用目录（默认分页 200 条，隐藏工具优先），传 query 搜索，传 tool 取完整参数 schema。',
            inputSchema: {
                type: 'object',
                properties: {
                    app: { type: 'string', description: '应用：excel / word / ppt / common' },
                    query: { type: 'string', description: '按名称或描述搜索关键字（支持中文场景词，如 题注/文档比较/母版）' },
                    tool: { type: 'string', description: '工具名，返回完整 inputSchema' },
                    offset: { type: 'number', description: '分页起点（从 0 开始），默认 0' },
                    limit: { type: 'number', description: '本页最多返回几条，默认 200，上限 500' },
                },
            },
            category: tools_1.ToolCategory.COMMON,
        }, async (args) => {
            const all = this.registry.listTools().tools;
            const discoverable = all;
            const wanted = typeof args.tool === 'string' ? args.tool.trim() : '';
            if (wanted) {
                const exact = all.find((tool) => tool.name === wanted);
                const found = exact || all.find((tool) => tool.name.endsWith(wanted));
                if (!found)
                    return failure('未找到工具 ' + wanted + '，请先用 wps_help 查询目录');
                return text({ name: found.name, description: found.description, inputSchema: found.inputSchema });
            }
            const app = typeof args.app === 'string' ? args.app.trim().toLowerCase() : '';
            const query = typeof args.query === 'string' ? args.query.trim().toLowerCase() : '';
            const requestedOffset = typeof args.offset === 'number' ? Math.floor(args.offset) : 0;
            const requestedLimit = typeof args.limit === 'number' ? Math.floor(args.limit) : 200;
            const offset = Number.isFinite(requestedOffset) ? Math.max(0, requestedOffset) : 0;
            const limit = Number.isFinite(requestedLimit) ? Math.min(500, Math.max(1, requestedLimit)) : 200;
            const advertisedNames = new Set((0, toolset_1.selectTools)(currentMode(), all).map((tool) => tool.name));
            const appOf = (name) => {
                const match = /^wps_(excel|word|ppt|common|convert)_/.exec(name);
                return match ? match[1] : 'builtin';
            };
            if (!app && !query) {
                const groups = {};
                for (const tool of discoverable) {
                    groups[appOf(tool.name)] = (groups[appOf(tool.name)] || 0) + 1;
                }
                return text({
                    toolset: currentMode(),
                    groups,
                    total: discoverable.length,
                    advertised: advertisedNames.size,
                    hidden: discoverable.length - advertisedNames.size,
                    usage: 'wps_help 传 app 查某应用目录，传 query 搜索，传 tool 取参数 schema，然后用 wps_call 执行',
                });
            }
            let pool = discoverable;
            if (app) {
                pool = pool.filter((tool) => tool.name.startsWith('wps_' + app + '_'));
            }
            let scored = pool.map((tool) => ({ tool, score: 0 }));
            if (query) {
                // Free-text search with synonym expansion: "题注/更新域/母版" do not appear in any
                // tool name, so score the expanded terms too; the original bigram path still runs.
                const tokens = (0, toolset_1.tokenizeHelpQuery)(query);
                scored = scored
                    .map((entry) => ({
                    ...entry,
                    score: (0, toolset_1.scoreToolHelpMatchExpanded)(entry.tool.name, entry.tool.description || '', query, tokens),
                }))
                    .filter((entry) => entry.score > 0);
            }
            // Hidden tools first: wps_help exists to surface the 175 tools the model cannot see.
            scored.sort((a, b) => {
                if (b.score !== a.score)
                    return b.score - a.score;
                const aAdv = advertisedNames.has(a.tool.name) ? 1 : 0;
                const bAdv = advertisedNames.has(b.tool.name) ? 1 : 0;
                if (aAdv !== bAdv)
                    return aAdv - bAdv;
                return a.tool.name.localeCompare(b.tool.name);
            });
            const matched = scored.length;
            if (query && matched === 0) {
                return text({
                    matched: 0,
                    shown: 0,
                    offset,
                    limit,
                    truncated: false,
                    nextOffset: null,
                    tools: [],
                    scenarios: (0, toolset_1.matchHelpScenarios)(query),
                    hint: '没有匹配的工具。改试 wps_help {app:"excel"|"word"|"ppt"} 列该应用目录，或 wps_help {tool:"完整工具名"} 取参数 schema。',
                });
            }
            const page = scored.slice(offset, offset + limit);
            const shown = page.map((entry) => ({
                name: entry.tool.name,
                app: appOf(entry.tool.name),
                advertised: advertisedNames.has(entry.tool.name),
                description: (0, toolset_1.compactDescription)(entry.tool.description, 80),
            }));
            const apps = {};
            for (const entry of page)
                apps[appOf(entry.tool.name)] = (apps[appOf(entry.tool.name)] || 0) + 1;
            const truncated = offset + shown.length < matched;
            const result = {
                matched,
                offset,
                limit,
                shown: shown.length,
                truncated,
                nextOffset: truncated ? offset + shown.length : null,
                apps,
                tools: shown,
            };
            if (query) {
                const scenarios = (0, toolset_1.matchHelpScenarios)(query);
                if (scenarios.length)
                    result.scenarios = scenarios;
                if (Object.keys(apps).length > 1) {
                    result.disambiguation = '结果跨多个应用：app 字段标出归属；同名能力（如"保护/公式/排序"）请按当前活动应用选择，勿把 Excel 工具当 Word 工具用。';
                }
            }
            return text(result);
        });
        // 通用派发
        this.registry.register({
            name: 'wps_call',
            description: '执行任意已注册但未直接广告的 WPS 工具；先用 wps_help 取到工具名与参数',
            inputSchema: {
                type: 'object',
                properties: {
                    tool: { type: 'string', description: '完整工具名，如 wps_ppt_set_animation' },
                    args: { type: 'object', description: '该工具的参数对象' },
                },
                required: ['tool'],
            },
            category: tools_1.ToolCategory.COMMON,
        }, async (args) => {
            const name = typeof args.tool === 'string' ? args.tool.trim() : '';
            if (!name)
                return failure('缺少 tool 参数');
            if (toolset_1.FACADE_TOOLS.includes(name))
                return failure('门面工具 ' + name + ' 不能通过 wps_call 调用');
            if (!this.registry.hasTool(name)) {
                return failure('未知工具 ' + name + '，请先用 wps_help 查询');
            }
            const inner = args.args && typeof args.args === 'object' ? args.args : {};
            return this.registry.callTool(tool_registry_1.ToolRegistry.createRequest(name, inner));
        });
        // 批量执行
        this.registry.register({
            name: 'wps_batch',
            description: '按顺序批量执行多个 WPS 工具调用，用于跨应用的连续操作；单次最多 50 项。' +
                '默认遇到第一个失败就停止（stopOnError=true），整体返回 status: completed/partial/failed 与逐项结果；' +
                '整体 success 只表示「批处理本身跑完了」，每一项是否生效看 results[].success。',
            inputSchema: {
                type: 'object',
                properties: {
                    calls: {
                        type: 'array',
                        description: '调用列表，每项为 {tool, args}',
                        items: {
                            type: 'object',
                            properties: {
                                tool: { type: 'string' },
                                args: { type: 'object' },
                            },
                            required: ['tool'],
                        },
                    },
                    stopOnError: {
                        type: 'boolean',
                        description: 'true 时遇到第一个失败就停止后续调用（默认 true）。确有独立步骤要逐项继续时显式传 false。',
                    },
                    maxResultChars: {
                        type: 'number',
                        description: '每项结果最多保留多少字符（默认 2000，范围 200–20000）；超长会标 truncated。',
                    },
                },
                required: ['calls'],
            },
            category: tools_1.ToolCategory.COMMON,
        }, async (args) => {
            const raw = Array.isArray(args.calls) ? args.calls : [];
            if (raw.length === 0)
                return failure('calls 不能为空');
            if (raw.length > 50)
                return failure('单次批量最多 50 项');
            const stopOnError = args.stopOnError !== false;
            const requestedMax = typeof args.maxResultChars === 'number' ? Math.floor(args.maxResultChars) : 2000;
            const maxResultChars = Number.isFinite(requestedMax) ? Math.min(20000, Math.max(200, requestedMax)) : 2000;
            const results = [];
            let succeeded = 0;
            let failedCount = 0;
            let stoppedAt = null;
            for (let index = 0; index < raw.length; index++) {
                const item = raw[index];
                const entry = (item || {});
                const tool = typeof entry.tool === 'string' ? entry.tool : '';
                if (!callable(tool)) {
                    failedCount++;
                    results.push({ index, tool, success: false, error: '无效或不可调用的工具名' });
                    if (stopOnError) {
                        stoppedAt = index;
                        break;
                    }
                    continue;
                }
                const inner = entry.args && typeof entry.args === 'object' ? entry.args : {};
                const outcome = await this.registry.callTool(tool_registry_1.ToolRegistry.createRequest(tool, inner));
                const blocks = Array.isArray(outcome.content) ? outcome.content : [];
                const joined = blocks
                    .map((block) => (block && typeof block === 'object' && 'text' in block ? String(block.text || '') : ''))
                    .join('\n');
                const record = { index, tool, success: outcome.success, result: joined.slice(0, maxResultChars) };
                if (joined.length > maxResultChars) {
                    record.truncated = true;
                    record.originalChars = joined.length;
                }
                if (outcome.success) {
                    succeeded++;
                }
                else {
                    failedCount++;
                    // 结构化错误**不截断**：截断过的错误没法诊断，而正文才是有长度压力的那部分。
                    if (typeof outcome.error === 'string' && outcome.error)
                        record.error = outcome.error;
                    if (outcome.partial === true)
                        record.partial = true;
                    if (Array.isArray(outcome.shortfalls) && outcome.shortfalls.length)
                        record.shortfalls = outcome.shortfalls;
                    if (stopOnError) {
                        stoppedAt = index;
                        results.push(record);
                        break;
                    }
                }
                results.push(record);
            }
            // 整体状态是显式字段，不靠调用方去推理："批处理跑完了吗、全成了吗"必须一眼能答。
            const status = failedCount === 0 ? 'completed' : succeeded > 0 ? 'partial' : 'failed';
            return text({
                count: results.length,
                requested: raw.length,
                succeeded,
                failed: failedCount,
                status,
                stopOnError,
                stoppedAt,
                results,
            });
        });
        logger.info('Registered facade tools', { tools: toolset_1.FACADE_TOOLS });
    }
    /**
     * 启动服务器
     */
    async start() {
        if (this.isRunning) {
            logger.warn('Server is already running');
            return;
        }
        logger.info('Starting MCP Server...');
        // 注册内置Tools
        this.registerBuiltinTools();
        // 注册门面工具（status / help / call / batch）
        this.registerFacadeTools();
        // 注册Excel、Word、PPT专业Tools - 这才是老王的核心功能
        this.registry.registerAll(tools_2.allTools);
        logger.info(`Registered ${tools_2.allTools.length} professional tools (Excel/Word/PPT)`);
        // 创建stdio传输层
        const transport = new stdio_js_1.StdioServerTransport();
        // 连接传输层
        await this.server.connect(transport);
        this.isRunning = true;
        logger.info('MCP Server started successfully');
    }
    /**
     * 停止服务器
     */
    async stop() {
        if (!this.isRunning) {
            logger.warn('Server is not running');
            return;
        }
        logger.info('Stopping MCP Server...');
        await this.server.close();
        // Release the COM host; it quits the WPS instances it started (never the user's) before exiting.
        try {
            await com_host_1.comHost.stop();
        }
        catch { /* shutting down anyway */ }
        this.isRunning = false;
        logger.info('MCP Server stopped');
    }
    /**
     * 获取服务器状态
     */
    getStatus() {
        return {
            running: this.isRunning,
            toolCount: this.registry.size,
        };
    }
}
exports.WpsMcpServer = WpsMcpServer;
// 导出单例创建函数
const createMcpServer = (config) => {
    return new WpsMcpServer(config);
};
exports.createMcpServer = createMcpServer;
exports.default = WpsMcpServer;
//# sourceMappingURL=mcp-server.js.map
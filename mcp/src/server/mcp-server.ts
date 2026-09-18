/**
 * Input: MCP 协议请求
 * Output: Tool 调用响应
 * Pos: MCP Server 核心实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * MCP Server实现 - 老王的MCP协议实现
 * 这是整个系统的核心，处理MCP协议通信
 * 基于stdio传输，Claude Desktop就是这么连的
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ErrorCode as McpErrorCode,
  McpError as SdkMcpError,
} from '@modelcontextprotocol/sdk/types.js';

import { toolRegistry, ToolRegistry } from './tool-registry';
import { wpsClient } from '../client/wps-client';
import { comHost } from '../client/com-host';
import { ToolCallResult, ToolCategory } from '../types/tools';
import {
  ToolsetMode,
  resolveMode,
  selectTools,
  compactDescription,
  scoreToolHelpMatch,
  tokenizeHelpQuery,
  FACADE_TOOLS,
} from './toolset';
import { DEPRECATED_TOOLS, DEPRECATED_NAMES, renameArgs } from '../tools/deprecated';
import { allTools } from '../tools';
import { createChildLogger } from '../utils/logger';
import { McpError } from '../utils/error';

const logger = createChildLogger('McpServer');

/**
 * MCP Server配置
 */
export interface McpServerConfig {
  /** 服务器名称 */
  name: string;
  /** 服务器版本 */
  version: string;
  /** 是否启用调试模式 */
  debug?: boolean;
}

/**
 * 默认配置
 */
const DEFAULT_CONFIG: McpServerConfig = {
  name: 'wps-office-mcp',
  version: '1.0.0',
  debug: false,
};

/**
 * WPS MCP Server - 核心服务器类
 */
export class WpsMcpServer {
  private readonly config: McpServerConfig;
  private readonly server: Server;
  private readonly registry: ToolRegistry;
  private isRunning: boolean = false;

  // 已合并掉的重复工具数量，供 wps_status 汇报
  private deprecatedToolCount = 0;

  constructor(config?: Partial<McpServerConfig>) {
    this.config = { ...DEFAULT_CONFIG, ...config };
    this.registry = toolRegistry;

    // 创建MCP Server实例
    this.server = new Server(
      {
        name: this.config.name,
        version: this.config.version,
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

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
  private setupRequestHandlers(): void {
    // 处理 tools/list 请求
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      logger.debug('Handling tools/list request');

      // 工具面收敛：只广告当前档位选择的工具，其余仍可通过 wps_call 调用
      const mode = resolveMode(process.env.WPS_OFFICE_TOOLSET);
      const { tools } = this.registry.listTools();
      const advertised = selectTools(mode, tools);

      logger.info('Returning ' + advertised.length + '/' + tools.length + ' tools (mode=' + mode + ')');

      return {
        tools: advertised.map((tool) => ({
          name: tool.name,
          description: mode === 'full' ? tool.description : compactDescription(tool.description),
          inputSchema: tool.inputSchema,
        })),
      };
    });

    // 处理 tools/call 请求
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      logger.debug('Handling tools/call request', { name, args });

      // 已废弃的旧名字在派发期解析成规范工具：不再各占一个注册位与一份重复 schema
      const aliased = this.resolveDeprecated(name, args || {});
      const targetName = aliased ? aliased.name : name;
      const targetArgs = aliased ? aliased.args : (args || {});

      // 检查Tool是否存在
      if (!this.registry.hasTool(targetName)) {
        throw new SdkMcpError(
          McpErrorCode.MethodNotFound,
          `Unknown tool: ${name}`
        );
      }

      try {
        // 创建调用请求并执行
        const callRequest = ToolRegistry.createRequest(targetName, targetArgs);
        const result: ToolCallResult = await this.registry.callTool(callRequest);

        return {
          content: result.content,
          isError: !result.success,
        };
      } catch (error) {
        logger.error('Tool call failed', error);

        if (error instanceof McpError) {
          return {
            content: [{ type: 'text', text: error.message }],
            isError: true,
          };
        }

        throw new SdkMcpError(
          McpErrorCode.InternalError,
          error instanceof Error ? error.message : 'Unknown error'
        );
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
  registerBuiltinTools(): void {
    logger.info('Registering the escape-hatch built-in tool');

    this.registry.register(
      {
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
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const method = args.method as string;
        const params = args.params as Record<string, unknown> | undefined;
        const appType = args.appType as string | undefined;

        const response = await wpsClient.executeMethod(
          method,
          params,
          appType as any
        );

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
      }
    );
  }

  /**
   * 注册门面工具 - 常驻广告的四个入口
   * wps_call 让全部已注册工具保持可用，而 tools/list 只广告一小部分
   */
  registerFacadeTools(): void {
    const currentMode = (): ToolsetMode => resolveMode(process.env.WPS_OFFICE_TOOLSET);

    const text = (value: unknown): ToolCallResult => ({
      id: '',
      success: true,
      content: [{ type: 'text', text: JSON.stringify(value) }],
    });

    const failure = (message: string): ToolCallResult => ({
      id: '',
      success: false,
      content: [{ type: 'text', text: message }],
      error: message,
    });

    // 废弃名不再注册，所以要同时认规范工具与可解析的别名
    const callable = (name: string): boolean => {
      if (!name || FACADE_TOOLS.includes(name)) return false;
      if (this.registry.hasTool(name)) return true;
      const spec = DEPRECATED_TOOLS[name];
      return !!spec && this.registry.hasTool(spec.canonical);
    };

    // 状态总览
    this.registry.register(
      {
        name: 'wps_status',
        description: '查看 WPS 连接状态、当前活动应用与所选工具面；编辑前先调用它',
        inputSchema: { type: 'object', properties: {} },
        category: ToolCategory.COMMON,
      },
      async () => {
        const started = Date.now();
        let connected = false;
        let appInfo: unknown = null;
        let note = '';
        try {
          const ping = await wpsClient.executeMethod('ping');
          connected = ping.success === true;
        } catch (error) {
          note = error instanceof Error ? error.message : String(error);
        }
        if (connected) {
          try {
            const info = await wpsClient.executeMethod('getAppInfo');
            appInfo = info.success ? (info.data || null) : null;
          } catch (error) {
            note = error instanceof Error ? error.message : String(error);
          }
        }
        const all = this.registry.listTools().tools;
        const advertised = selectTools(currentMode(), all);
        return text({
          connected,
          appInfo,
          toolset: currentMode(),
          advertisedTools: advertised.length,
          registeredTools: all.length,
          hiddenTools: all.length - advertised.length,
          deprecatedTools: this.deprecatedToolCount,
          note: note || undefined,
          latencyMs: Date.now() - started,
        });
      }
    );

    // 目录与 schema 查询
    this.registry.register(
      {
        name: 'wps_help',
        description: '查询未直接广告的工具：无参看分组概览，传 app 或 query 查目录，传 tool 取完整参数 schema',
        inputSchema: {
          type: 'object',
          properties: {
            app: { type: 'string', description: '应用：excel / word / ppt / common' },
            query: { type: 'string', description: '按名称或描述搜索关键字' },
            tool: { type: 'string', description: '工具名，返回完整 inputSchema' },
          },
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const all = this.registry.listTools().tools;
        const discoverable = all.filter((tool) => !DEPRECATED_NAMES.has(tool.name));
        const wanted = typeof args.tool === 'string' ? args.tool.trim() : '';
        if (wanted) {
          // 废弃名已不在注册表里，必须先查别名表，否则 wps_help {tool:"旧名"} 会变成"未找到"
          const deprecatedSpec = DEPRECATED_TOOLS[wanted];
          if (deprecatedSpec) {
            const canonical = all.find((tool) => tool.name === deprecatedSpec.canonical);
            return text({
              name: wanted,
              deprecated: true,
              canonical: deprecatedSpec.canonical,
              reason: deprecatedSpec.reason,
              inputSchema: canonical ? canonical.inputSchema : undefined,
            });
          }
          const exact = all.find((tool) => tool.name === wanted);
          const found = exact || all.find((tool) => tool.name.endsWith(wanted));
          if (!found) return failure('未找到工具 ' + wanted + '，请先用 wps_help 查询目录');
          return text({ name: found.name, description: found.description, inputSchema: found.inputSchema });
        }

        const app = typeof args.app === 'string' ? args.app.trim().toLowerCase() : '';
        const query = typeof args.query === 'string' ? args.query.trim().toLowerCase() : '';

        if (!app && !query) {
          const groups: Record<string, number> = {};
          for (const tool of discoverable) {
            const match = /^wps_(excel|word|ppt)_/.exec(tool.name);
            const key = match ? match[1] : (/^wps_(common|convert)_/.test(tool.name) ? 'common' : 'builtin');
            groups[key] = (groups[key] || 0) + 1;
          }
          return text({
            toolset: currentMode(),
            groups,
            total: discoverable.length,
            usage: 'wps_help 传 app 查某应用目录，传 query 搜索，传 tool 取参数 schema，然后用 wps_call 执行',
          });
        }

        let pool = discoverable;
        if (app) {
          pool = pool.filter((tool) => tool.name.startsWith('wps_' + app + '_'));
        }
        if (query) {
          // Free-text search: score per token (plus CJK bigrams) instead of testing the whole
          // phrase as one substring, which answered "matched: 0" for natural multi-word asks.
          const tokens = tokenizeHelpQuery(query);
          pool = pool
            .map((tool) => ({ tool, score: scoreToolHelpMatch(tool.name, tool.description || '', query, tokens) }))
            .filter((entry) => entry.score > 0)
            .sort((a, b) => b.score - a.score || a.tool.name.localeCompare(b.tool.name))
            .map((entry) => entry.tool);
        }
        if (query && pool.length === 0) {
          return text({
            matched: 0,
            shown: 0,
            truncated: false,
            tools: [],
            hint: '没有匹配的工具。改试 wps_help {app:"excel"|"word"|"ppt"} 列该应用目录，或 wps_help {tool:"完整工具名"} 取参数 schema。',
          });
        }
        const shown = pool.slice(0, 60).map((tool) => ({
          name: tool.name,
          description: compactDescription(tool.description, 80),
        }));
        return text({ matched: pool.length, shown: shown.length, truncated: pool.length > shown.length, tools: shown });
      }
    );

    // 通用派发
    this.registry.register(
      {
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
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const name = typeof args.tool === 'string' ? args.tool.trim() : '';
        if (!name) return failure('缺少 tool 参数');
        if (FACADE_TOOLS.includes(name)) return failure('门面工具 ' + name + ' 不能通过 wps_call 调用');
        if (!this.registry.hasTool(name)) {
          const spec = DEPRECATED_TOOLS[name];
          if (!spec || !this.registry.hasTool(spec.canonical)) {
            return failure('未知工具 ' + name + '，请先用 wps_help 查询');
          }
        }
        const inner = args.args && typeof args.args === 'object' ? (args.args as Record<string, unknown>) : {};
        const aliased = this.resolveDeprecated(name, inner);
        return this.registry.callTool(
          ToolRegistry.createRequest(aliased ? aliased.name : name, aliased ? aliased.args : inner)
        );
      }
    );

    // 批量执行
    this.registry.register(
      {
        name: 'wps_batch',
        description: '按顺序批量执行多个 WPS 工具调用，用于跨应用的连续操作；单次最多 50 项',
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
          },
          required: ['calls'],
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const raw = Array.isArray(args.calls) ? args.calls : [];
        if (raw.length === 0) return failure('calls 不能为空');
        if (raw.length > 50) return failure('单次批量最多 50 项');
        const results: unknown[] = [];
        for (const item of raw) {
          const entry = (item || {}) as Record<string, unknown>;
          const tool = typeof entry.tool === 'string' ? entry.tool : '';
          if (!callable(tool)) {
            results.push({ tool, success: false, error: '无效或不可调用的工具名' });
            continue;
          }
          const inner = entry.args && typeof entry.args === 'object' ? (entry.args as Record<string, unknown>) : {};
          const aliased = this.resolveDeprecated(tool, inner);
          const outcome = await this.registry.callTool(
            ToolRegistry.createRequest(aliased ? aliased.name : tool, aliased ? aliased.args : inner)
          );
          const blocks = Array.isArray(outcome.content) ? outcome.content : [];
          const joined = blocks
            .map((block) => (block && typeof block === 'object' && 'text' in block ? String((block as { text?: string }).text || '') : ''))
            .join('\n');
          results.push({ tool, success: outcome.success, result: joined.slice(0, 2000) });
        }
        return text({ count: results.length, results });
      }
    );

    logger.info('Registered facade tools', { tools: FACADE_TOOLS });
  }

  /**
   * 校验废弃别名指向的规范工具都还在，并把可解析的别名数量记下来供 wps_status 汇报。
   * 别名不再注册成工具：旧名字在派发期解析（resolveDeprecated），既不占注册位也不重复 schema。
   */
  resolveDeprecatedTools(): void {
    let applied = 0;
    for (const [name, spec] of Object.entries(DEPRECATED_TOOLS)) {
      if (!this.registry.hasTool(spec.canonical)) {
        logger.warn('Canonical tool missing for deprecated alias', { name, canonical: spec.canonical });
        continue;
      }
      applied++;
    }
    this.deprecatedToolCount = applied;
    logger.info('Resolved deprecated tool aliases', { applied });
  }

  /** 废弃名 → {规范工具名, 改名后的参数}；不是废弃名、或规范工具缺失时返回 null。 */
  private resolveDeprecated(
    name: string,
    args: Record<string, unknown>
  ): { name: string; args: Record<string, unknown> } | null {
    const spec = DEPRECATED_TOOLS[name];
    if (!spec || !this.registry.hasTool(spec.canonical)) return null;
    logger.warn('Deprecated tool called', { name, canonical: spec.canonical });
    return { name: spec.canonical, args: renameArgs(args, spec.paramMap) };
  }

  /**
   * 启动服务器
   */
  async start(): Promise<void> {
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
    this.registry.registerAll(allTools);
    logger.info(`Registered ${allTools.length} professional tools (Excel/Word/PPT)`);

    // 把完全等价的重复工具收敛到规范名
    this.resolveDeprecatedTools();

    // 创建stdio传输层
    const transport = new StdioServerTransport();

    // 连接传输层
    await this.server.connect(transport);

    this.isRunning = true;
    logger.info('MCP Server started successfully');
  }

  /**
   * 停止服务器
   */
  async stop(): Promise<void> {
    if (!this.isRunning) {
      logger.warn('Server is not running');
      return;
    }

    logger.info('Stopping MCP Server...');

    await this.server.close();

    // Release the COM host; it quits the WPS instances it started (never the user's) before exiting.
    try { await comHost.stop(); } catch { /* shutting down anyway */ }

    this.isRunning = false;
    logger.info('MCP Server stopped');
  }

  /**
   * 获取服务器状态
   */
  getStatus(): { running: boolean; toolCount: number } {
    return {
      running: this.isRunning,
      toolCount: this.registry.size,
    };
  }
}

// 导出单例创建函数
export const createMcpServer = (
  config?: Partial<McpServerConfig>
): WpsMcpServer => {
  return new WpsMcpServer(config);
};

export default WpsMcpServer;

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
import { ToolCallResult, ToolCategory } from '../types/tools';
import { ToolsetMode, resolveMode, selectTools, compactDescription, FACADE_TOOLS } from './toolset';
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

  // 跨应用数据缓存 - 解决macOS WPS无法跨应用操作的P0问题
  private static dataCache: Map<string, { data: unknown; timestamp: number; appType: string }> = new Map();

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

      // 检查Tool是否存在
      if (!this.registry.hasTool(name)) {
        throw new SdkMcpError(
          McpErrorCode.MethodNotFound,
          `Unknown tool: ${name}`
        );
      }

      try {
        // 创建调用请求并执行
        const callRequest = ToolRegistry.createRequest(name, args || {});
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
  registerBuiltinTools(): void {
    logger.info('Registering built-in tools');

    // 连接状态检查Tool
    this.registry.register(
      {
        name: 'wps_check_connection',
        description: '检查WPS Office连接状态',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.COMMON,
      },
      async () => {
        const connected = await wpsClient.checkConnection();
        const status = wpsClient.getStatus();

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                connected,
                status,
              }),
            },
          ],
        };
      }
    );

    // 获取当前文档信息
    this.registry.register(
      {
        name: 'wps_get_active_document',
        description: '获取当前打开的WPS文字文档信息',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.DOCUMENT,
      },
      async () => {
        const doc = await wpsClient.getActiveDocument();

        return {
          id: '',
          success: doc !== null,
          content: [
            {
              type: 'text',
              text: doc
                ? JSON.stringify(doc)
                : '没有打开的文档',
            },
          ],
        };
      }
    );

    // 在文档中插入文本
    this.registry.register(
      {
        name: 'wps_insert_text',
        description: '在当前文档中插入文本',
        inputSchema: {
          type: 'object',
          properties: {
            text: {
              type: 'string',
              description: '要插入的文本内容',
            },
            position: {
              type: 'number',
              description: '插入位置（可选，不指定则在光标处插入）',
            },
          },
          required: ['text'],
        },
        category: ToolCategory.DOCUMENT,
      },
      async (args) => {
        const text = args.text as string;
        const position = args.position as number | undefined;

        const success = await wpsClient.insertText(text, position);

        return {
          id: '',
          success,
          content: [
            {
              type: 'text',
              text: success ? '文本插入成功' : '文本插入失败',
            },
          ],
        };
      }
    );

    // 获取当前工作簿信息
    this.registry.register(
      {
        name: 'wps_get_active_workbook',
        description: '获取当前打开的WPS表格工作簿信息',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.SPREADSHEET,
      },
      async () => {
        const workbook = await wpsClient.getActiveWorkbook();

        return {
          id: '',
          success: workbook !== null,
          content: [
            {
              type: 'text',
              text: workbook
                ? JSON.stringify(workbook)
                : '没有打开的工作簿',
            },
          ],
        };
      }
    );

    // 读取单元格值
    this.registry.register(
      {
        name: 'wps_get_cell_value',
        description: '读取指定单元格的值',
        inputSchema: {
          type: 'object',
          properties: {
            sheet: {
              type: 'string',
              description: '工作表名称或索引',
            },
            row: {
              type: 'number',
              description: '行号（从1开始）',
            },
            col: {
              type: 'number',
              description: '列号（从1开始）',
            },
          },
          required: ['sheet', 'row', 'col'],
        },
        category: ToolCategory.SPREADSHEET,
      },
      async (args) => {
        const sheet = args.sheet as string | number;
        const row = args.row as number;
        const col = args.col as number;

        const value = await wpsClient.getCellValue(sheet, row, col);

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({ value }),
            },
          ],
        };
      }
    );

    // 设置单元格值
    this.registry.register(
      {
        name: 'wps_set_cell_value',
        description: '设置指定单元格的值',
        inputSchema: {
          type: 'object',
          properties: {
            sheet: {
              type: 'string',
              description: '工作表名称或索引',
            },
            row: {
              type: 'number',
              description: '行号（从1开始）',
            },
            col: {
              type: 'number',
              description: '列号（从1开始）',
            },
            value: {
              type: 'string',
              description: '要设置的值',
            },
          },
          required: ['sheet', 'row', 'col', 'value'],
        },
        category: ToolCategory.SPREADSHEET,
      },
      async (args) => {
        const sheet = args.sheet as string | number;
        const row = args.row as number;
        const col = args.col as number;
        const value = args.value;

        const success = await wpsClient.setCellValue(sheet, row, col, value);

        return {
          id: '',
          success,
          content: [
            {
              type: 'text',
              text: success ? '单元格值设置成功' : '单元格值设置失败',
            },
          ],
        };
      }
    );

    // 获取当前演示文稿信息
    this.registry.register(
      {
        name: 'wps_get_active_presentation',
        description: '获取当前打开的WPS演示文稿信息',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.PRESENTATION,
      },
      async () => {
        const presentation = await wpsClient.getActivePresentation();

        return {
          id: '',
          success: presentation !== null,
          content: [
            {
              type: 'text',
              text: presentation
                ? JSON.stringify(presentation)
                : '没有打开的演示文稿',
            },
          ],
        };
      }
    );

    // 执行自定义WPS方法
    this.registry.register(
      {
        name: 'wps_execute_method',
        description: '执行自定义WPS API方法',
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

    // ==================== 跨应用数据缓存工具 ====================
    // 解决macOS WPS加载项无法跨应用操作的P0问题
    // Excel读取数据 → 缓存到MCP Server → PPT获取缓存 → 创建演示文稿

    // 缓存数据
    this.registry.register(
      {
        name: 'wps_cache_data',
        description: '缓存数据到MCP Server，用于跨应用数据传递。例如：从Excel读取数据后缓存，然后在PPT中使用。',
        inputSchema: {
          type: 'object',
          properties: {
            key: {
              type: 'string',
              description: '缓存键名，用于后续获取数据',
            },
            data: {
              type: 'object',
              description: '要缓存的数据（任意JSON对象）',
            },
            appType: {
              type: 'string',
              description: '数据来源应用类型：et（表格）、wps（文字）、wpp（演示）',
              enum: ['et', 'wps', 'wpp'],
            },
          },
          required: ['key', 'data'],
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const key = args.key as string;
        const data = args.data;
        const appType = (args.appType as string) || 'unknown';

        WpsMcpServer.dataCache.set(key, {
          data,
          timestamp: Date.now(),
          appType,
        });

        logger.info(`Data cached: key=${key}, appType=${appType}`);

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                key,
                message: `数据已缓存，可在其他应用中通过 wps_get_cached_data 获取`,
                cacheSize: WpsMcpServer.dataCache.size,
              }),
            },
          ],
        };
      }
    );

    // 获取缓存数据
    this.registry.register(
      {
        name: 'wps_get_cached_data',
        description: '从MCP Server获取缓存的数据，用于跨应用数据传递。',
        inputSchema: {
          type: 'object',
          properties: {
            key: {
              type: 'string',
              description: '缓存键名',
            },
          },
          required: ['key'],
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const key = args.key as string;
        const cached = WpsMcpServer.dataCache.get(key);

        if (!cached) {
          return {
            id: '',
            success: false,
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: false,
                  error: `缓存键 "${key}" 不存在`,
                  availableKeys: Array.from(WpsMcpServer.dataCache.keys()),
                }),
              },
            ],
          };
        }

        logger.info(`Cache hit: key=${key}, age=${Date.now() - cached.timestamp}ms`);

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                key,
                data: cached.data,
                appType: cached.appType,
                cachedAt: new Date(cached.timestamp).toISOString(),
              }),
            },
          ],
        };
      }
    );

    // 列出所有缓存
    this.registry.register(
      {
        name: 'wps_list_cache',
        description: '列出MCP Server中所有缓存的数据键名',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.COMMON,
      },
      async () => {
        const cacheList = Array.from(WpsMcpServer.dataCache.entries()).map(([key, value]) => ({
          key,
          appType: value.appType,
          cachedAt: new Date(value.timestamp).toISOString(),
          ageMs: Date.now() - value.timestamp,
        }));

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                count: cacheList.length,
                caches: cacheList,
              }),
            },
          ],
        };
      }
    );

    // 清除缓存
    this.registry.register(
      {
        name: 'wps_clear_cache',
        description: '清除MCP Server中的缓存数据',
        inputSchema: {
          type: 'object',
          properties: {
            key: {
              type: 'string',
              description: '要清除的缓存键名，不指定则清除所有缓存',
            },
          },
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const key = args.key as string | undefined;

        if (key) {
          const deleted = WpsMcpServer.dataCache.delete(key);
          logger.info(`Cache cleared: key=${key}, deleted=${deleted}`);

          return {
            id: '',
            success: true,
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: deleted ? `缓存 "${key}" 已清除` : `缓存 "${key}" 不存在`,
                }),
              },
            ],
          };
        } else {
          const count = WpsMcpServer.dataCache.size;
          WpsMcpServer.dataCache.clear();
          logger.info(`All cache cleared: count=${count}`);

          return {
            id: '',
            success: true,
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: `已清除所有缓存，共 ${count} 条`,
                }),
              },
            ],
          };
        }
      }
    );

    logger.info(`Registered ${this.registry.size} built-in tools`);
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

    const callable = (name: string): boolean =>
      !!name && !FACADE_TOOLS.includes(name) && this.registry.hasTool(name);

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
          const exact = all.find((tool) => tool.name === wanted);
          const found = exact || all.find((tool) => tool.name.endsWith(wanted));
          if (!found) return failure('未找到工具 ' + wanted + '，请先用 wps_help 查询目录');
          const spec = DEPRECATED_TOOLS[found.name];
          if (spec) {
            const canonical = all.find((tool) => tool.name === spec.canonical);
            return text({
              name: found.name,
              deprecated: true,
              canonical: spec.canonical,
              reason: spec.reason,
              inputSchema: canonical ? canonical.inputSchema : found.inputSchema,
            });
          }
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
          pool = pool.filter((tool) => tool.name.toLowerCase().includes(query) || (tool.description || '').toLowerCase().includes(query));
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
        if (!this.registry.hasTool(name)) return failure('未知工具 ' + name + '，请先用 wps_help 查询');
        const inner = args.args && typeof args.args === 'object' ? (args.args as Record<string, unknown>) : {};
        return this.registry.callTool(ToolRegistry.createRequest(name, inner));
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
          const outcome = await this.registry.callTool(ToolRegistry.createRequest(tool, inner));
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
   * 把与规范工具完全等价的重复工具改成转发别名
   * 旧名字仍然可用，但不再出现在 wps_help 的目录里
   */
  applyDeprecatedTools(): void {
    let applied = 0;
    for (const [name, spec] of Object.entries(DEPRECATED_TOOLS)) {
      const existing = this.registry.getTool(name);
      if (!existing) continue;
      if (!this.registry.hasTool(spec.canonical)) {
        logger.warn('Canonical tool missing for deprecated alias', { name, canonical: spec.canonical });
        continue;
      }
      const definition = {
        ...existing.definition,
        description: '[已废弃] ' + existing.definition.description + ' 请改用 ' + spec.canonical + '。',
      };
      this.registry.unregister(name);
      this.registry.register(definition, async (args) => {
        logger.warn('Deprecated tool called', { name, canonical: spec.canonical });
        const mapped = renameArgs(args, spec.paramMap);
        return this.registry.callTool(ToolRegistry.createRequest(spec.canonical, mapped));
      });
      applied++;
    }
    this.deprecatedToolCount = applied;
    logger.info('Applied deprecated tool aliases', { applied });
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
    this.applyDeprecatedTools();

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

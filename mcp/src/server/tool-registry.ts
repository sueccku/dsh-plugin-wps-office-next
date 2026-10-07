/**
 * Input: Tool 定义与调用请求
 * Output: Tool 执行结果
 * Pos: MCP Tool 注册与调度中心。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tool注册管理 - 老王的Tool管理系统
 * 所有Tool都得在这儿注册，不注册的Tool就是野鸡Tool
 * 这个设计遵循OCP原则：扩展开放，修改关闭
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  RegisteredTool,
  ToolCallRequest,
  ToolCallResult,
  ToolCategory,
  ListToolsResponse,
} from '../types/tools';
import { createChildLogger } from '../utils/logger';
import { runWithWarningCollector } from '../utils/tool-warnings';
import {
  ToolNotFoundError,
  ToolExecutionError,
  InvalidParamsError,
  McpError,
} from '../utils/error';

const logger = createChildLogger('ToolRegistry');

/**
 * Tool注册表 - 单例模式，全局唯一
 */
/** 编辑距离：只用来给“未知参数”提一句“是不是想传 X”，不做任何语义判断。 */
function editDistance(a: string, b: string): number {
  const prev = new Array<number>(b.length + 1);
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    let diag = prev[0];
    prev[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const tmp = prev[j];
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, diag + (a[i - 1] === b[j - 1] ? 0 : 1));
      diag = tmp;
    }
  }
  return prev[b.length];
}

/** 在声明的参数里找最接近的一个；只有足够接近才提示，避免把人带偏。 */
function closestParamName(key: string, accepted: string[]): string | undefined {
  const norm = (s: string) => s.toLowerCase().replace(/[_-]/g, '');
  const target = norm(key);
  let best: string | undefined;
  let bestScore = Number.POSITIVE_INFINITY;
  for (const candidate of accepted) {
    const c = norm(candidate);
    if (c === target) return candidate;
    const score = editDistance(target, c);
    if (score < bestScore) { bestScore = score; best = candidate; }
  }
  const limit = Math.max(2, Math.floor(target.length / 3));
  return best !== undefined && bestScore <= limit ? best : undefined;
}

export class ToolRegistry {
  private static instance: ToolRegistry;
  private readonly tools: Map<string, RegisteredTool>;
  private readonly categories: Map<ToolCategory, Set<string>>;

  private constructor() {
    this.tools = new Map();
    this.categories = new Map();

    // 初始化分类
    Object.values(ToolCategory).forEach((category) => {
      this.categories.set(category, new Set());
    });

    logger.info('ToolRegistry initialized');
  }

  /**
   * 获取单例实例
   */
  static getInstance(): ToolRegistry {
    if (!ToolRegistry.instance) {
      ToolRegistry.instance = new ToolRegistry();
    }
    return ToolRegistry.instance;
  }

  /**
   * 注册Tool - 把Tool加到注册表里
   * 遵循SRP：只负责注册，不负责其他
   */
  register(definition: ToolDefinition, handler: ToolHandler): void {
    const { name, category } = definition;

    // 检查是否已注册 - 跳过重复注册而非崩溃，防止过期编译产物导致启动失败
    if (this.tools.has(name)) {
      logger.warn(`Tool already registered, skipping: ${name}`);
      return;
    }

    // 注册Tool
    this.tools.set(name, { definition, handler });

    // 添加到分类
    const toolCategory = category || ToolCategory.COMMON;
    this.categories.get(toolCategory)?.add(name);

    logger.info(`Tool registered: ${name}`, { category: toolCategory });
  }

  /**
   * 批量注册Tools - 一次注册一堆
   */
  registerAll(tools: Array<{ definition: ToolDefinition; handler: ToolHandler }>): void {
    tools.forEach(({ definition, handler }) => {
      this.register(definition, handler);
    });
    logger.info(`Batch registered ${tools.length} tools`);
  }

  /**
   * 注销Tool - 把Tool从注册表里删掉
   */
  unregister(name: string): boolean {
    const tool = this.tools.get(name);
    if (!tool) {
      return false;
    }

    this.tools.delete(name);

    // 从分类中移除
    const category = tool.definition.category || ToolCategory.COMMON;
    this.categories.get(category)?.delete(name);

    logger.info(`Tool unregistered: ${name}`);
    return true;
  }

  /**
   * 获取Tool定义
   */
  getTool(name: string): RegisteredTool | undefined {
    return this.tools.get(name);
  }

  /**
   * 检查Tool是否存在
   */
  hasTool(name: string): boolean {
    return this.tools.has(name);
  }

  /**
   * 获取所有Tool定义 - MCP的tools/list用的
   */
  listTools(): ListToolsResponse {
    const tools = Array.from(this.tools.values()).map((t) => t.definition);
    return { tools };
  }

  /**
   * 按分类获取Tools
   */
  getToolsByCategory(category: ToolCategory): ToolDefinition[] {
    const toolNames = this.categories.get(category);
    if (!toolNames) {
      return [];
    }

    return Array.from(toolNames)
      .map((name) => this.tools.get(name)?.definition)
      .filter((t): t is ToolDefinition => t !== undefined);
  }

  /**
   * 调用Tool - 执行Tool的handler
   * 这是核心方法，处理Tool调用请求
   */
  async callTool(request: ToolCallRequest): Promise<ToolCallResult> {
    const { id, name, arguments: args } = request;
    const startTime = Date.now();

    logger.debug(`Calling tool: ${name}`, { id, args });

    // 检查Tool是否存在
    const tool = this.tools.get(name);
    if (!tool) {
      throw new ToolNotFoundError(name);
    }

    try {
      // 验证参数
      this.validateArguments(tool.definition, args);

      // 执行handler，同时收集桥侧留下的 warnings
      const { value: result, warnings } = await runWithWarningCollector(() => tool.handler(args));
      const withWarnings = this.attachWarnings(result, warnings);

      const duration = Date.now() - startTime;
      logger.info(`Tool executed: ${name}`, { id, duration, success: result.success, warnings: warnings.length });

      return {
        ...withWarnings,
        id,
      };
    } catch (error) {
      const duration = Date.now() - startTime;

      if (error instanceof McpError) {
        logger.error(`Tool execution failed: ${name}`, error, { id, duration });
        return {
          id,
          success: false,
          content: [{ type: 'text', text: error.message }],
          error: error.message,
        };
      }

      const execError = new ToolExecutionError(
        name,
        error instanceof Error ? error : new Error(String(error))
      );

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
   * 把「尽力而为的失败」附在结果文本后面。第一方 handler 会丢掉桥侧的 warnings，这里统一补上；
   * 已经出现在文本里的（wps_call / wps_execute_method / wps_batch 的原样透传）不重复追加（FIXES 73）。
   */
  private attachWarnings(result: ToolCallResult, warnings: string[]): ToolCallResult {
    if (!warnings.length) return result;
    const missing = warnings.filter(
      (w) => !result.content.some((block) => block.type === 'text' && typeof block.text === 'string' && block.text.includes(w))
    );
    if (!missing.length) return result;
    const note =
      '\n\n注意（' + missing.length + ' 条，主操作已完成，但其中某些步骤没有成功）：\n' +
      missing.map((w) => '- ' + w).join('\n');
    if (!result.content.length) {
      return { ...result, content: [{ type: 'text', text: note.trim() }] };
    }
    return {
      ...result,
      content: result.content.map((block, index) =>
        index === 0 && block.type === 'text' ? { ...block, text: (block.text ?? '') + note } : block
      ),
    };
  }

  /**
   * 验证参数 - 检查必填参数是否都有
   */
  private validateArguments(
    definition: ToolDefinition,
    args: Record<string, unknown>
  ): void {
    const { inputSchema } = definition;
    const required = inputSchema.required || [];

    for (const param of required) {
      if (!(param in args) || args[param] === undefined || args[param] === null) {
        throw new InvalidParamsError(`Missing required parameter: ${param}`, {
          toolName: definition.name,
          missingParam: param,
        });
      }
    }

    // 结构性类型校验：只挡“形状”明显错位的入参，标量之间一律放行。
    // 只挡形状的原因：schema 里 value: string 这类声明比桥接层的真实契约更窄（桥接层/COM 会把
    // 42 和 "42" 都吞下去），按 schema 强校验会误伤合法调用；而 data: "[[1,2]]"、calls: {} 这类
    // 容器/标量错位，桥接层只能给出很晦涩的报错，提前挡掉更清楚。
    // 枚举值仍然由桥接层裁决（错误措辞含 unknown xxx、未知 xxx，有测试覆盖），这里不重复校验。
    const properties = inputSchema.properties || {};
    const isScalar = (t: string) => t === 'string' || t === 'number' || t === 'boolean';

    for (const [key, value] of Object.entries(args)) {
      if (value === undefined || value === null) {
        continue;
      }

      const schema = properties[key];
      if (!schema) {
        // FIXES 99（W6-4）：以前这里是 `continue`，注释写着“未声明的参数交给桥接层报未知参数”。
        // 那个假设是**错的**：TS handler 用显式解构拼参数（`{ range, delimiter, sheet }`），
        // 未知键在到达桥之前就被丢掉了，桥的“未知键拒绝”永远看不见它 —— 于是调用方传错参数名时
        // **不报错、按默认值执行、回报还写着自己传的那个值**。
        // 实证：`wps_excel_text_to_columns { range:"A1:A3", sep:"," }` 不报错，回报「分隔符: ","」。
        // 这条路径上必须响亮拒绝 —— 它是唯一还看得见原始实参的地方。
        const accepted = Object.keys(properties);
        const hint = closestParamName(key, accepted);
        throw new InvalidParamsError(
          `未知参数: ${key}（动作：${definition.name}）` +
            `可选参数: ${accepted.length ? accepted.join(' / ') : '(无)'}` +
            (hint ? `。是不是想传 ${hint}？` : '。') +
            `下一步：改成上面列出的参数名重试；若这个参数确实用不上，把它从调用里删掉。`,
          { toolName: definition.name, unknownParam: key }
        );
      }

      const actual = Array.isArray(value) ? 'array' : typeof value === 'object' ? 'object' : typeof value;
      // type 允许是数组（JSON Schema 合法写法，例如 ['object','string']）。以前只读 schema.type，
      // 数组会落进"需要 X，实际 Y"的假报错里（FIXES 90：apply_style 的 range 就踩了这个）。
      const expectedTypes = Array.isArray(schema.type) ? (schema.type as string[]) : [schema.type as string];
      const shapeOk = expectedTypes.some((t) => t === actual || (isScalar(t) && isScalar(actual)));
      const expected = expectedTypes.join(' | ');

      if (!shapeOk) {
        throw new InvalidParamsError(
          `参数 ${key} 的形状不对：需要 ${expected}，实际收到 ${actual}`,
          { toolName: definition.name, param: key, expected, actual }
        );
      }
    }
  }

  /**
   * 获取Tool数量
   */
  get size(): number {
    return this.tools.size;
  }

  /**
   * 清空所有Tools - 测试用，生产环境别乱用
   */
  clear(): void {
    this.tools.clear();
    this.categories.forEach((set) => set.clear());
    logger.warn('All tools cleared');
  }

  /**
   * 创建Tool调用请求 - 辅助方法
   */
  static createRequest(
    name: string,
    args: Record<string, unknown>
  ): ToolCallRequest {
    return {
      id: uuidv4(),
      name,
      arguments: args,
    };
  }
}

// 导出单例
export const toolRegistry = ToolRegistry.getInstance();

/**
 * 装饰器：注册Tool
 * 用法：@RegisterTool(definition)
 */
export function RegisterTool(definition: ToolDefinition) {
  return function (
    _target: unknown,
    _propertyKey: string,
    descriptor: PropertyDescriptor
  ): PropertyDescriptor {
    const originalMethod = descriptor.value as ToolHandler;
    toolRegistry.register(definition, originalMethod);
    return descriptor;
  };
}

/**
 * 快捷注册函数
 */
export const registerTool = (
  definition: ToolDefinition,
  handler: ToolHandler
): void => {
  toolRegistry.register(definition, handler);
};

export default ToolRegistry;

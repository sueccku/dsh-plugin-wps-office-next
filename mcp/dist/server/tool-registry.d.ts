/**
 * Input: Tool 定义与调用请求
 * Output: Tool 执行结果
 * Pos: MCP Tool 注册与调度中心。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tool注册管理 - 老王的Tool管理系统
 * 所有Tool都得在这儿注册，不注册的Tool就是野鸡Tool
 * 这个设计遵循OCP原则：扩展开放，修改关闭
 */
import { ToolDefinition, ToolHandler, RegisteredTool, ToolCallRequest, ToolCallResult, ToolCategory, ListToolsResponse } from '../types/tools';
/**
 * Tool注册表 - 单例模式，全局唯一
 */
export declare class ToolRegistry {
    private static instance;
    private readonly tools;
    private readonly categories;
    private constructor();
    /**
     * 获取单例实例
     */
    static getInstance(): ToolRegistry;
    /**
     * 注册Tool - 把Tool加到注册表里
     * 遵循SRP：只负责注册，不负责其他
     */
    register(definition: ToolDefinition, handler: ToolHandler): void;
    /**
     * 批量注册Tools - 一次注册一堆
     */
    registerAll(tools: Array<{
        definition: ToolDefinition;
        handler: ToolHandler;
    }>): void;
    /**
     * 注销Tool - 把Tool从注册表里删掉
     */
    unregister(name: string): boolean;
    /**
     * 获取Tool定义
     */
    getTool(name: string): RegisteredTool | undefined;
    /**
     * 检查Tool是否存在
     */
    hasTool(name: string): boolean;
    /**
     * 获取所有Tool定义 - MCP的tools/list用的
     */
    listTools(): ListToolsResponse;
    /**
     * 按分类获取Tools
     */
    getToolsByCategory(category: ToolCategory): ToolDefinition[];
    /**
     * 调用Tool - 执行Tool的handler
     * 这是核心方法，处理Tool调用请求
     */
    callTool(request: ToolCallRequest): Promise<ToolCallResult>;
    /**
     * 验证参数 - 检查必填参数是否都有
     */
    private validateArguments;
    /**
     * 获取Tool数量
     */
    get size(): number;
    /**
     * 清空所有Tools - 测试用，生产环境别乱用
     */
    clear(): void;
    /**
     * 创建Tool调用请求 - 辅助方法
     */
    static createRequest(name: string, args: Record<string, unknown>): ToolCallRequest;
}
export declare const toolRegistry: ToolRegistry;
/**
 * 装饰器：注册Tool
 * 用法：@RegisterTool(definition)
 */
export declare function RegisterTool(definition: ToolDefinition): (_target: unknown, _propertyKey: string, descriptor: PropertyDescriptor) => PropertyDescriptor;
/**
 * 快捷注册函数
 */
export declare const registerTool: (definition: ToolDefinition, handler: ToolHandler) => void;
export default ToolRegistry;
//# sourceMappingURL=tool-registry.d.ts.map
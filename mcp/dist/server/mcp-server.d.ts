/**
 * Input: MCP 协议请求
 * Output: Tool 调用响应
 * Pos: MCP Server 核心实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * MCP Server实现 - 老王的MCP协议实现
 * 这是整个系统的核心，处理MCP协议通信
 * 基于stdio传输，Claude Desktop就是这么连的
 */
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
 * WPS MCP Server - 核心服务器类
 */
export declare class WpsMcpServer {
    private readonly config;
    private readonly server;
    private readonly registry;
    private isRunning;
    private static dataCache;
    private deprecatedToolCount;
    constructor(config?: Partial<McpServerConfig>);
    /**
     * 设置请求处理器
     */
    private setupRequestHandlers;
    /**
     * 注册内置Tools - 一些基础的WPS操作Tool
     */
    registerBuiltinTools(): void;
    /**
     * 注册门面工具 - 常驻广告的四个入口
     * wps_call 让全部已注册工具保持可用，而 tools/list 只广告一小部分
     */
    registerFacadeTools(): void;
    /**
     * 把与规范工具完全等价的重复工具改成转发别名
     * 旧名字仍然可用，但不再出现在 wps_help 的目录里
     */
    applyDeprecatedTools(): void;
    /**
     * 启动服务器
     */
    start(): Promise<void>;
    /**
     * 停止服务器
     */
    stop(): Promise<void>;
    /**
     * 获取服务器状态
     */
    getStatus(): {
        running: boolean;
        toolCount: number;
    };
}
export declare const createMcpServer: (config?: Partial<McpServerConfig>) => WpsMcpServer;
export default WpsMcpServer;
//# sourceMappingURL=mcp-server.d.ts.map
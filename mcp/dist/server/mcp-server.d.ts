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
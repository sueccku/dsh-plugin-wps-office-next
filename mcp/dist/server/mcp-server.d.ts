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
    private deprecatedToolCount;
    constructor(config?: Partial<McpServerConfig>);
    /**
     * 设置请求处理器
     */
    private setupRequestHandlers;
    /**
     * 注册内置Tools - 一些基础的WPS操作Tool
     */
    /**
     * 注册逃生舱工具。原有的 11 个 builtin 已删除（重复/缓存/连接检查），
     * 只保留 wps_execute_method：覆盖尚未工具化时唯一的自逃生路径，文档里明确不推荐。
     */
    registerBuiltinTools(): void;
    /**
     * 注册门面工具 - 常驻广告的四个入口
     * wps_call 让全部已注册工具保持可用，而 tools/list 只广告一小部分
     */
    registerFacadeTools(): void;
    /**
     * 校验废弃别名指向的规范工具都还在，并把可解析的别名数量记下来供 wps_status 汇报。
     * 别名不再注册成工具：旧名字在派发期解析（resolveDeprecated），既不占注册位也不重复 schema。
     */
    resolveDeprecatedTools(): void;
    /** 废弃名 → {规范工具名, 改名后的参数}；不是废弃名、或规范工具缺失时返回 null。 */
    private resolveDeprecated;
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
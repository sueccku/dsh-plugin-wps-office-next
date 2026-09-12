/**
 * Input: 运行环境与启动参数
 * Output: MCP Server 导出与启动日志
 * Pos: MCP 服务入口模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS Office MCP Server - 入口文件
 * 老王出品，让AI能操控WPS Office
 *
 * 启动方式：
 * - 开发模式：npm run dev
 * - 生产模式：npm run build && npm start
 */
export { WpsMcpServer, createMcpServer } from './server/mcp-server';
export { ToolRegistry, toolRegistry, registerTool } from './server/tool-registry';
export { WpsClient, wpsClient } from './client/wps-client';
export { log, logger, createChildLogger } from './utils/logger';
export * from './utils/error';
export * from './types/tools';
export * from './types/wps';
/**
 * 主函数 - 程序入口
 */
declare function main(): Promise<void>;
export default main;
//# sourceMappingURL=index.d.ts.map
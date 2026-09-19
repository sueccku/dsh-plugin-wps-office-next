"use strict";
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
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __exportStar = (this && this.__exportStar) || function(m, exports) {
    for (var p in m) if (p !== "default" && !Object.prototype.hasOwnProperty.call(exports, p)) __createBinding(exports, m, p);
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.createChildLogger = exports.logger = exports.log = exports.wpsClient = exports.WpsClient = exports.registerTool = exports.toolRegistry = exports.ToolRegistry = exports.createMcpServer = exports.WpsMcpServer = void 0;
const node_fs_1 = require("node:fs");
const node_path_1 = require("node:path");
const mcp_server_1 = require("./server/mcp-server");
const logger_1 = require("./utils/logger");
// 导出所有模块，方便外部使用
var mcp_server_2 = require("./server/mcp-server");
Object.defineProperty(exports, "WpsMcpServer", { enumerable: true, get: function () { return mcp_server_2.WpsMcpServer; } });
Object.defineProperty(exports, "createMcpServer", { enumerable: true, get: function () { return mcp_server_2.createMcpServer; } });
var tool_registry_1 = require("./server/tool-registry");
Object.defineProperty(exports, "ToolRegistry", { enumerable: true, get: function () { return tool_registry_1.ToolRegistry; } });
Object.defineProperty(exports, "toolRegistry", { enumerable: true, get: function () { return tool_registry_1.toolRegistry; } });
Object.defineProperty(exports, "registerTool", { enumerable: true, get: function () { return tool_registry_1.registerTool; } });
var wps_client_1 = require("./client/wps-client");
Object.defineProperty(exports, "WpsClient", { enumerable: true, get: function () { return wps_client_1.WpsClient; } });
Object.defineProperty(exports, "wpsClient", { enumerable: true, get: function () { return wps_client_1.wpsClient; } });
var logger_2 = require("./utils/logger");
Object.defineProperty(exports, "log", { enumerable: true, get: function () { return logger_2.log; } });
Object.defineProperty(exports, "logger", { enumerable: true, get: function () { return logger_2.logger; } });
Object.defineProperty(exports, "createChildLogger", { enumerable: true, get: function () { return logger_2.createChildLogger; } });
__exportStar(require("./utils/error"), exports);
__exportStar(require("./types/tools"), exports);
__exportStar(require("./types/wps"), exports);
/**
 * 主函数 - 程序入口
 */
async function main() {
    const mainLogger = (0, logger_1.createChildLogger)('Main');
    mainLogger.info('='.repeat(50));
    mainLogger.info('WPS Office MCP Server');
    mainLogger.info('老王出品，必属精品');
    mainLogger.info('='.repeat(50));
    // The reported server version follows the published bundle version, so there is exactly one
    // place to bump on release (the root package.json). Falls back if the file cannot be read.
    let bundleVersion = '0.0.0';
    try {
        bundleVersion = JSON.parse((0, node_fs_1.readFileSync)((0, node_path_1.join)(__dirname, '..', '..', 'package.json'), 'utf8')).version || bundleVersion;
    }
    catch { /* the version is informational; never fail startup over it */ }
    // 创建服务器实例
    const server = (0, mcp_server_1.createMcpServer)({
        name: 'wps-office-mcp',
        version: bundleVersion,
        debug: process.env.DEBUG === 'true',
    });
    // 优雅关闭处理
    const gracefulShutdown = async (signal) => {
        mainLogger.info(`Received ${signal}, shutting down gracefully...`);
        try {
            await server.stop();
            mainLogger.info('Server stopped successfully');
            process.exit(0);
        }
        catch (error) {
            mainLogger.error('Error during shutdown', error);
            process.exit(1);
        }
    };
    // 监听进程信号
    process.on('SIGINT', () => gracefulShutdown('SIGINT'));
    process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));
    // 未捕获异常处理
    process.on('uncaughtException', (error) => {
        mainLogger.error('Uncaught exception', error);
        process.exit(1);
    });
    process.on('unhandledRejection', (reason) => {
        mainLogger.error('Unhandled rejection', reason);
        process.exit(1);
    });
    try {
        // 启动服务器
        await server.start();
        const status = server.getStatus();
        mainLogger.info('Server is running', status);
        mainLogger.info('Waiting for MCP client connection...');
        // 服务器会一直运行，等待MCP客户端连接
        // stdio传输层会保持进程活跃
    }
    catch (error) {
        mainLogger.error('Failed to start server', error);
        process.exit(1);
    }
}
// 如果是直接运行而不是被导入，则启动服务器
if (require.main === module) {
    main().catch((error) => {
        logger_1.log.error('Fatal error', error);
        process.exit(1);
    });
}
exports.default = main;
//# sourceMappingURL=index.js.map
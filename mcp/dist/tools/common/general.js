"use strict";
/**
 * Input: 通用操作工具参数
 * Output: 操作结果
 * Pos: 通用操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 通用操作Tools - 保存/连接检测/文本选取等基础模块
 *
 * 包含：
 * - wps_common_save: 保存当前文档
 * - wps_common_save_as: 另存为
 * - wps_common_ping: 检测WPS连接
 * - wps_common_wire_check: 检查通信线路
 * - wps_common_get_app_info: 获取WPS应用信息
 * - wps_common_get_selected_text: 获取选中文本
 * - wps_common_set_selected_text: 替换选中文本
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.generalTools = exports.setSelectedTextHandler = exports.setSelectedTextDefinition = exports.getSelectedTextHandler = exports.getSelectedTextDefinition = exports.getAppInfoHandler = exports.getAppInfoDefinition = exports.wireCheckHandler = exports.wireCheckDefinition = exports.pingHandler = exports.pingDefinition = exports.saveAsHandler = exports.saveAsDefinition = exports.saveHandler = exports.saveDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
// ============================================================
// 1. wps_common_save - 保存当前文档
// ============================================================
exports.saveDefinition = {
    name: 'wps_common_save',
    description: `保存当前文档。

使用场景：
- "保存一下"
- "Ctrl+S"
- "把修改存起来"

特点：
- 自动保存当前活动文档
- 如果文档从未保存过，可能会提示选择保存路径`,
    category: tools_1.ToolCategory.COMMON,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const saveHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('save', {});
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    { type: 'text', text: `文档保存成功！${response.data?.message || ''}` },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `保存失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `保存出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.saveHandler = saveHandler;
// ============================================================
// 2. wps_common_save_as - 另存为
// ============================================================
exports.saveAsDefinition = {
    name: 'wps_common_save_as',
    description: `将当前文档另存为指定路径和格式。

使用场景：
- "另存为到桌面"
- "换个名字保存"
- "保存一份副本到指定位置"

特点：
- 支持指定完整文件路径
- 可选指定保存格式`,
    category: tools_1.ToolCategory.COMMON,
    inputSchema: {
        type: 'object',
        properties: {
            filePath: {
                type: 'string',
                description: '目标文件完整路径，包含文件名和扩展名',
            },
            format: {
                type: 'string',
                description: '保存格式（可选），如 docx, xlsx, pptx 等',
            },
        },
        required: ['filePath'],
    },
};
const saveAsHandler = async (args) => {
    const { filePath, format } = args;
    if (!filePath || filePath.trim() === '') {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '文件路径不能为空，请指定保存路径' }],
            error: '文件路径为空',
        };
    }
    try {
        const params = {
            filePath,
            path: filePath,
            outputPath: filePath,
        };
        if (format) {
            params.format = format.toLowerCase().replace(/^\./, '');
        }
        const response = await wps_client_1.wpsClient.executeMethod('saveAs', params);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `另存为成功！\n输出路径: ${response.data.outputPath || filePath}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `另存为失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `另存为出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.saveAsHandler = saveAsHandler;
// ============================================================
// 3. wps_common_ping - 检测WPS连接
// ============================================================
exports.pingDefinition = {
    name: 'wps_common_ping',
    description: `检测WPS应用连接状态。

使用场景：
- "WPS连上了吗"
- "检查一下WPS是否在线"
- "测试WPS连接"

特点：
- 快速检测WPS加载项是否可达
- 返回连接状态信息`,
    category: tools_1.ToolCategory.COMMON,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const pingHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('ping', {});
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    { type: 'text', text: `WPS连接正常！${response.data?.message || 'pong'}` },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `WPS连接失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `WPS连接检测出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.pingHandler = pingHandler;
// ============================================================
// 4. wps_common_wire_check - 检查通信线路
// ============================================================
exports.wireCheckDefinition = {
    name: 'wps_common_wire_check',
    description: `检查与WPS加载项之间的通信线路状态。

使用场景：
- "检查通信状态"
- "诊断连接问题"
- "通信线路是否正常"

特点：
- 比ping更详细的通信诊断
- 返回线路延迟、协议状态等信息`,
    category: tools_1.ToolCategory.COMMON,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const wireCheckHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('wireCheck', {});
        if (response.success && response.data) {
            const data = response.data;
            let info = '通信线路正常！';
            if (data.latency !== undefined) {
                info += `\n延迟: ${data.latency}ms`;
            }
            if (data.protocol) {
                info += `\n协议: ${data.protocol}`;
            }
            if (data.message) {
                info += `\n${data.message}`;
            }
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: info }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `通信线路检查失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `通信线路检查出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.wireCheckHandler = wireCheckHandler;
// ============================================================
// 5. wps_common_get_app_info - 获取WPS应用信息
// ============================================================
exports.getAppInfoDefinition = {
    name: 'wps_common_get_app_info',
    description: `获取WPS应用的基本信息。

使用场景：
- "WPS是什么版本"
- "查看WPS信息"
- "获取应用状态"

特点：
- 返回WPS版本号、构建信息
- 返回当前打开的文档信息
- 返回运行平台信息`,
    category: tools_1.ToolCategory.COMMON,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const getAppInfoHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getAppInfo', {});
        if (response.success && response.data) {
            const data = response.data;
            const lines = ['WPS应用信息：'];
            if (data.version)
                lines.push(`版本: ${data.version}`);
            if (data.build)
                lines.push(`构建号: ${data.build}`);
            if (data.platform)
                lines.push(`平台: ${data.platform}`);
            if (data.activeDocument)
                lines.push(`当前文档: ${data.activeDocument}`);
            if (data.message)
                lines.push(data.message);
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: lines.join('\n') }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取应用信息失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取应用信息出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getAppInfoHandler = getAppInfoHandler;
// ============================================================
// 6. wps_common_get_selected_text - 获取选中文本
// ============================================================
exports.getSelectedTextDefinition = {
    name: 'wps_common_get_selected_text',
    description: `获取当前文档中选中的文本内容。

使用场景：
- "读取我选中的内容"
- "获取当前选区的文字"
- "看看我选了什么"

特点：
- 返回当前选区的纯文本内容
- 适用于Word、Excel、PPT`,
    category: tools_1.ToolCategory.COMMON,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const getSelectedTextHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSelectedText', {});
        if (response.success && response.data) {
            const text = response.data.text;
            if (!text || text.length === 0) {
                return {
                    id: (0, uuid_1.v4)(),
                    success: true,
                    content: [{ type: 'text', text: '当前没有选中任何文本。' }],
                };
            }
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    { type: 'text', text: `选中的文本内容：\n${text}` },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取选中文本失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取选中文本出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getSelectedTextHandler = getSelectedTextHandler;
// ============================================================
// 7. wps_common_set_selected_text - 替换选中文本
// ============================================================
exports.setSelectedTextDefinition = {
    name: 'wps_common_set_selected_text',
    description: `替换当前文档中选中的文本内容。

使用场景：
- "把选中的内容替换成xxx"
- "修改选区的文字"
- "用新内容替换选中部分"

特点：
- 将当前选区的文本替换为指定内容
- 适用于Word、Excel、PPT`,
    category: tools_1.ToolCategory.COMMON,
    inputSchema: {
        type: 'object',
        properties: {
            text: {
                type: 'string',
                description: '要替换为的新文本内容',
            },
        },
        required: ['text'],
    },
};
const setSelectedTextHandler = async (args) => {
    const { text } = args;
    if (text === undefined || text === null) {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '替换文本不能为空，请指定要替换的内容' }],
            error: '替换文本为空',
        };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSelectedText', { text });
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    { type: 'text', text: `选中文本已替换成功！${response.data?.message || ''}` },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `替换选中文本失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `替换选中文本出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setSelectedTextHandler = setSelectedTextHandler;
// ============================================================
// 导出所有通用操作Tools
// ============================================================
exports.generalTools = [
    { definition: exports.saveDefinition, handler: exports.saveHandler },
    { definition: exports.saveAsDefinition, handler: exports.saveAsHandler },
    { definition: exports.pingDefinition, handler: exports.pingHandler },
    { definition: exports.wireCheckDefinition, handler: exports.wireCheckHandler },
    { definition: exports.getAppInfoDefinition, handler: exports.getAppInfoHandler },
    { definition: exports.getSelectedTextDefinition, handler: exports.getSelectedTextHandler },
    { definition: exports.setSelectedTextDefinition, handler: exports.setSelectedTextHandler },
];
exports.default = exports.generalTools;
//# sourceMappingURL=general.js.map
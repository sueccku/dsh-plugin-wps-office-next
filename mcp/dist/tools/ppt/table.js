"use strict";
/**
 * Input: PPT 表格操作参数
 * Output: 表格操作结果
 * Pos: PPT 表格工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT表格Tools - 表格管理模块
 * 处理表格的插入、单元格读写、样式设置等操作
 *
 * 包含：
 * - wps_ppt_insert_table: 在幻灯片中插入表格
 * - wps_ppt_set_table_cell: 设置表格单元格文本
 * - wps_ppt_get_table_cell: 获取表格单元格文本
 * - wps_ppt_set_table_style: 设置表格整体样式
 * - wps_ppt_set_table_cell_style: 设置表格单元格样式
 * - wps_ppt_set_table_row_style: 设置表格行样式
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.tableTools = exports.setPptTableRowStyleHandler = exports.setPptTableRowStyleDefinition = exports.setPptTableCellStyleHandler = exports.setPptTableCellStyleDefinition = exports.setPptTableStyleHandler = exports.setPptTableStyleDefinition = exports.getPptTableCellHandler = exports.getPptTableCellDefinition = exports.setPptTableCellHandler = exports.setPptTableCellDefinition = exports.insertPptTableHandler = exports.insertPptTableDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 在幻灯片中插入表格
 * 支持指定行列数和位置
 */
exports.insertPptTableDefinition = {
    name: 'wps_ppt_insert_table',
    description: `在幻灯片中插入表格。

支持指定：
- 表格行列数
- 表格位置（左上角坐标）

使用场景：
- "在第2页插入一个3行4列的表格"
- "添加一个表格到当前页"
- "插入数据表格"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            rows: {
                type: 'number',
                description: '表格行数',
            },
            cols: {
                type: 'number',
                description: '表格列数',
            },
            left: {
                type: 'number',
                description: '表格左上角X坐标（磅），可选',
            },
            top: {
                type: 'number',
                description: '表格左上角Y坐标（磅），可选',
            },
        },
        required: ['slideIndex', 'rows', 'cols'],
    },
};
const insertPptTableHandler = async (args) => {
    const { slideIndex, rows, cols, left, top } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertPptTable', { slideIndex, rows, cols, left, top }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `表格插入成功！\n幻灯片: 第 ${slideIndex} 页\n规格: ${rows} 行 × ${cols} 列\n表格索引: ${response.data.tableIndex}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `插入表格失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `插入表格出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertPptTableHandler = insertPptTableHandler;
/**
 * 设置表格单元格文本
 * 修改指定单元格的内容
 */
exports.setPptTableCellDefinition = {
    name: 'wps_ppt_set_table_cell',
    description: `设置PPT表格中指定单元格的文本内容。

使用场景：
- "把表格第1行第2列的内容改成'销售额'"
- "设置单元格文本"
- "填写表格内容"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            tableIndex: {
                type: 'number',
                description: '表格索引（从1开始）',
            },
            row: {
                type: 'number',
                description: '行号（从1开始）',
            },
            col: {
                type: 'number',
                description: '列号（从1开始）',
            },
            text: {
                type: 'string',
                description: '要设置的文本内容',
            },
        },
        required: ['slideIndex', 'tableIndex', 'row', 'col', 'text'],
    },
};
const setPptTableCellHandler = async (args) => {
    const { slideIndex, tableIndex, row, col, text } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setPptTableCell', { slideIndex, tableIndex, row, col, text }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `单元格文本设置成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n位置: 第 ${row} 行第 ${col} 列\n内容: ${text}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置单元格文本失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置单元格文本出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setPptTableCellHandler = setPptTableCellHandler;
/**
 * 获取表格单元格文本
 * 读取指定单元格的内容
 */
exports.getPptTableCellDefinition = {
    name: 'wps_ppt_get_table_cell',
    description: `获取PPT表格中指定单元格的文本内容。

使用场景：
- "读取表格第1行第1列的内容"
- "获取单元格文本"
- "查看表格数据"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            tableIndex: {
                type: 'number',
                description: '表格索引（从1开始）',
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
        required: ['slideIndex', 'tableIndex', 'row', 'col'],
    },
};
const getPptTableCellHandler = async (args) => {
    const { slideIndex, tableIndex, row, col } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getPptTableCell', { slideIndex, tableIndex, row, col }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `单元格内容获取成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n位置: 第 ${row} 行第 ${col} 列\n内容: ${response.data.text}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取单元格内容失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取单元格内容出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getPptTableCellHandler = getPptTableCellHandler;
/**
 * 设置表格整体样式
 * 支持边框、背景色、字体等整体样式配置
 */
exports.setPptTableStyleDefinition = {
    name: 'wps_ppt_set_table_style',
    description: `设置PPT表格的整体样式。

支持的样式属性（通过style对象传入）：
- borderColor: 边框颜色（如 "#000000"）
- borderWidth: 边框宽度（磅）
- backgroundColor: 背景色（如 "#FFFFFF"）
- fontName: 字体名称（如 "微软雅黑"）
- fontSize: 字体大小（磅）
- fontColor: 字体颜色（如 "#333333"）
- headerBackground: 表头行背景色
- alternateRowColor: 隔行变色颜色

使用场景：
- "设置表格为蓝色主题"
- "修改表格边框和背景"
- "美化表格样式"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            tableIndex: {
                type: 'number',
                description: '表格索引（从1开始）',
            },
            style: {
                type: 'object',
                description: '样式配置对象，包含borderColor、borderWidth、backgroundColor、fontName、fontSize、fontColor、headerBackground、alternateRowColor等属性',
            },
        },
        required: ['slideIndex', 'tableIndex', 'style'],
    },
};
const setPptTableStyleHandler = async (args) => {
    const { slideIndex, tableIndex, style } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setPptTableStyle', { slideIndex, tableIndex, style }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const styleKeys = Object.keys(style).join('、');
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `表格样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n已设置属性: ${styleKeys}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置表格样式失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置表格样式出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setPptTableStyleHandler = setPptTableStyleHandler;
/**
 * 设置表格单元格样式
 * 修改指定单元格的样式（字体、颜色、背景等）
 */
exports.setPptTableCellStyleDefinition = {
    name: 'wps_ppt_set_table_cell_style',
    description: `设置PPT表格中指定单元格的样式。

支持的样式属性（通过style对象传入）：
- backgroundColor: 单元格背景色
- fontName: 字体名称
- fontSize: 字体大小（磅）
- fontColor: 字体颜色
- bold: 是否加粗（boolean）
- italic: 是否斜体（boolean）
- alignment: 文本对齐方式（left/center/right）
- verticalAlignment: 垂直对齐方式（top/middle/bottom）

使用场景：
- "把表格第1行第1列的背景设为蓝色"
- "加粗标题单元格"
- "设置单元格居中对齐"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            tableIndex: {
                type: 'number',
                description: '表格索引（从1开始）',
            },
            row: {
                type: 'number',
                description: '行号（从1开始）',
            },
            col: {
                type: 'number',
                description: '列号（从1开始）',
            },
            style: {
                type: 'object',
                description: '样式配置对象，包含backgroundColor、fontName、fontSize、fontColor、bold、italic、alignment、verticalAlignment等属性',
            },
        },
        required: ['slideIndex', 'tableIndex', 'row', 'col', 'style'],
    },
};
const setPptTableCellStyleHandler = async (args) => {
    const { slideIndex, tableIndex, row, col, style } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setPptTableCellStyle', { slideIndex, tableIndex, row, col, style }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const styleKeys = Object.keys(style).join('、');
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `单元格样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n位置: 第 ${row} 行第 ${col} 列\n已设置属性: ${styleKeys}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置单元格样式失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置单元格样式出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setPptTableCellStyleHandler = setPptTableCellStyleHandler;
/**
 * 设置表格行样式
 * 批量修改指定行的样式
 */
exports.setPptTableRowStyleDefinition = {
    name: 'wps_ppt_set_table_row_style',
    description: `设置PPT表格中指定行的样式。

支持的样式属性（通过style对象传入）：
- backgroundColor: 行背景色
- fontName: 字体名称
- fontSize: 字体大小（磅）
- fontColor: 字体颜色
- bold: 是否加粗（boolean）
- italic: 是否斜体（boolean）
- alignment: 文本对齐方式（left/center/right）
- height: 行高（磅）

使用场景：
- "把表头行设为蓝色背景白色字"
- "加粗第一行"
- "设置表格行高"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            tableIndex: {
                type: 'number',
                description: '表格索引（从1开始）',
            },
            row: {
                type: 'number',
                description: '行号（从1开始）',
            },
            style: {
                type: 'object',
                description: '样式配置对象，包含backgroundColor、fontName、fontSize、fontColor、bold、italic、alignment、height等属性',
            },
        },
        required: ['slideIndex', 'tableIndex', 'row', 'style'],
    },
};
const setPptTableRowStyleHandler = async (args) => {
    const { slideIndex, tableIndex, row, style } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setPptTableRowStyle', { slideIndex, tableIndex, row, style }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const styleKeys = Object.keys(style).join('、');
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `表格行样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n行: 第 ${row} 行\n已设置属性: ${styleKeys}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置行样式失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置行样式出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setPptTableRowStyleHandler = setPptTableRowStyleHandler;
/**
 * 导出所有表格相关的Tools
 */
exports.tableTools = [
    { definition: exports.insertPptTableDefinition, handler: exports.insertPptTableHandler },
    { definition: exports.setPptTableCellDefinition, handler: exports.setPptTableCellHandler },
    { definition: exports.getPptTableCellDefinition, handler: exports.getPptTableCellHandler },
    { definition: exports.setPptTableStyleDefinition, handler: exports.setPptTableStyleHandler },
    { definition: exports.setPptTableCellStyleDefinition, handler: exports.setPptTableCellStyleHandler },
    { definition: exports.setPptTableRowStyleDefinition, handler: exports.setPptTableRowStyleHandler },
];
exports.default = exports.tableTools;
//# sourceMappingURL=table.js.map
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
exports.tableTools = exports.setPptTableFormatHandler = exports.setPptTableFormatDefinition = exports.getPptTableCellHandler = exports.getPptTableCellDefinition = exports.setPptTableCellHandler = exports.setPptTableCellDefinition = exports.insertPptTableHandler = exports.insertPptTableDefinition = void 0;
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
 * 导出所有表格相关的Tools
 */
/** 表格外观：位置尺寸 + 单元格/整行/整表样式（P4 由三个工具合并而来） */
exports.setPptTableFormatDefinition = {
    name: 'wps_ppt_set_table_format',
    description: '设置表格的外观与样式。位置尺寸：left/top/width/height。样式作用域：给 row 和 col 就只改那一格，只给 row 改一整行，都不给但有样式键就改整张表；样式键为 backgroundColor/fontColor/fontSize/bold。使用场景：给表头行加底色、把表格挪到中间。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            left: { type: 'number', description: '表格左边缘位置' },
            top: { type: 'number', description: '表格上边缘位置' },
            width: { type: 'number', description: '表格宽度' },
            height: { type: 'number', description: '表格高度' },
            row: { type: 'number', description: '第几行（从 1 开始）；与 col 一起用就是单元格' },
            col: { type: 'number', description: '第几列（从 1 开始）' },
            backgroundColor: { type: 'string', description: '底纹颜色，十六进制如 #D9E2F3' },
            fontColor: { type: 'string', description: '文字颜色，十六进制如 #1A365D' },
            fontSize: { type: 'number', description: '字号' },
            bold: { type: 'boolean', description: '是否加粗' },
            slideIndex: { type: 'number', description: '第几页（从 1 开始），默认 1' },
            tableIndex: { type: 'number', description: '该页第几张表（从 1 开始），默认 1' },
            tableName: { type: 'string', description: '表格形状名；给了它就用名称定位' },
            presentationName: { type: 'string', description: '演示文稿名；不填用当前文稿' },
        },
    },
};
const setPptTableFormatHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setPptTableFormat', {
            presentationName: args.presentationName, slideIndex: args.slideIndex, tableIndex: args.tableIndex, tableName: args.tableName,
            row: args.row, col: args.col, backgroundColor: args.backgroundColor, fontColor: args.fontColor,
            fontSize: args.fontSize, bold: args.bold, left: args.left, top: args.top, width: args.width, height: args.height,
        }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '设置表格外观失败: ' + response.error }], error: response.error };
        }
        const d = response.data || {};
        const applied = d.applied || [];
        const scopeName = d.scope === 'cell' ? '单元格' : d.scope === 'row' ? '整行' : d.scope === 'table' ? '整张表' : '仅位置尺寸';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '表格 ' + String(d.name || '') + ' 已更新（作用域 ' + scopeName + '，已应用 ' + applied.join(', ') + '）' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '设置表格外观出错: ' + errMsg }], error: errMsg };
    }
};
exports.setPptTableFormatHandler = setPptTableFormatHandler;
exports.tableTools = [
    { definition: exports.setPptTableFormatDefinition, handler: exports.setPptTableFormatHandler },
    { definition: exports.insertPptTableDefinition, handler: exports.insertPptTableHandler },
    { definition: exports.setPptTableCellDefinition, handler: exports.setPptTableCellHandler },
    { definition: exports.getPptTableCellDefinition, handler: exports.getPptTableCellHandler },
];
exports.default = exports.tableTools;
//# sourceMappingURL=table.js.map
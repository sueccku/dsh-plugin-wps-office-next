"use strict";
/**
 * Input: 行列操作工具参数
 * Output: 行列插入/删除/隐藏/分组结果
 * Pos: Excel 行列操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel行列操作Tools - 行列管理模块
 * 处理行列的插入、删除、隐藏、显示、分组等操作
 *
 * 包含：
 * - wps_excel_insert_rows: 插入多行
 * - wps_excel_insert_columns: 插入多列
 * - wps_excel_delete_rows: 删除多行
 * - wps_excel_delete_columns: 删除多列
 * - wps_excel_hide_rows: 隐藏行
 * - wps_excel_show_rows: 显示行
 * - wps_excel_show_columns: 显示列
 * - wps_excel_group_rows: 分组行
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.rowColumnTools = exports.groupRowsHandler = exports.groupRowsDefinition = exports.showColumnsHandler = exports.showColumnsDefinition = exports.showRowsHandler = exports.showRowsDefinition = exports.hideRowsHandler = exports.hideRowsDefinition = exports.deleteColumnsHandler = exports.deleteColumnsDefinition = exports.deleteRowsHandler = exports.deleteRowsDefinition = exports.insertColumnsHandler = exports.insertColumnsDefinition = exports.insertRowsHandler = exports.insertRowsDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 插入多行
 */
exports.insertRowsDefinition = {
    name: 'wps_excel_insert_rows',
    description: '在Excel中指定位置插入一行或多行。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            row: { type: 'number', description: '在第几行前插入（从1开始）' },
            count: { type: 'number', description: '插入行数，默认1' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['row'],
    },
};
const insertRowsHandler = async (args) => {
    const { row, count, sheet } = args;
    const insertCount = count || 1;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertRows', { row, count: insertCount, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入行失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `插入行完成！在第${row}行前插入了${insertCount}行` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入行出错: ${errMsg}` }], error: errMsg };
    }
};
exports.insertRowsHandler = insertRowsHandler;
/**
 * 插入多列
 */
exports.insertColumnsDefinition = {
    name: 'wps_excel_insert_columns',
    description: '在Excel中指定位置插入一列或多列。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            column: { type: 'string', description: '在哪一列前插入，如 A、B、C' },
            count: { type: 'number', description: '插入列数，默认1' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['column'],
    },
};
const insertColumnsHandler = async (args) => {
    const { column, count, sheet } = args;
    const insertCount = count || 1;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertColumns', { column, count: insertCount, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入列失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `插入列完成！在${column}列前插入了${insertCount}列` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入列出错: ${errMsg}` }], error: errMsg };
    }
};
exports.insertColumnsHandler = insertColumnsHandler;
/**
 * 删除多行
 */
exports.deleteRowsDefinition = {
    name: 'wps_excel_delete_rows',
    description: '删除Excel中指定位置的一行或多行。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            startRow: { type: 'number', description: '起始行号（从1开始）' },
            row: { type: 'number', description: '要删除的行号（与 startRow 等价，便于合并旧工具）' },
            count: { type: 'number', description: '删除行数，默认1' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: [],
    },
};
const deleteRowsHandler = async (args) => {
    const { startRow, row, count, sheet } = args;
    const deleteCount = count || 1;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteRows', { startRow, row, count: deleteCount, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `删除行失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `删除行完成！从第${startRow}行开始删除了${deleteCount}行` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `删除行出错: ${errMsg}` }], error: errMsg };
    }
};
exports.deleteRowsHandler = deleteRowsHandler;
/**
 * 删除多列
 */
exports.deleteColumnsDefinition = {
    name: 'wps_excel_delete_columns',
    description: '删除Excel中指定位置的一列或多列。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            column: { type: 'string', description: '起始列字母，如 A、B、C' },
            count: { type: 'number', description: '删除列数，默认1' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['column'],
    },
};
const deleteColumnsHandler = async (args) => {
    const { column, count, sheet } = args;
    const deleteCount = count || 1;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteColumns', { column, count: deleteCount, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `删除列失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `删除列完成！从${column}列开始删除了${deleteCount}列` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `删除列出错: ${errMsg}` }], error: errMsg };
    }
};
exports.deleteColumnsHandler = deleteColumnsHandler;
/**
 * 隐藏行
 */
exports.hideRowsDefinition = {
    name: 'wps_excel_hide_rows',
    description: '隐藏Excel中指定范围的行。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            startRow: { type: 'number', description: '起始行号（从1开始）' },
            endRow: { type: 'number', description: '结束行号（从1开始）' },
            row: { type: 'number', description: '单行行号（与 startRow 等价）' },
            rows: { type: 'array', description: '行号数组' },
            count: { type: 'number', description: '从 row 起的连续行数' },
            hide: { type: 'boolean', description: 'true 隐藏（默认），false 显示' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: [],
    },
};
const hideRowsHandler = async (args) => {
    const { startRow, endRow, row, rows, count, hide, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('hideRows', { startRow, endRow, row, rows, count, hide, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `隐藏行失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `隐藏行完成！已隐藏第${startRow}行到第${endRow}行` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `隐藏行出错: ${errMsg}` }], error: errMsg };
    }
};
exports.hideRowsHandler = hideRowsHandler;
/**
 * 显示行
 */
exports.showRowsDefinition = {
    name: 'wps_excel_show_rows',
    description: '显示Excel中已隐藏的行。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            startRow: { type: 'number', description: '起始行号（从1开始）' },
            endRow: { type: 'number', description: '结束行号（从1开始）' },
            row: { type: 'number', description: '单行行号（与 startRow 等价）' },
            rows: { type: 'array', description: '行号数组' },
            count: { type: 'number', description: '从 row 起的连续行数' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: [],
    },
};
const showRowsHandler = async (args) => {
    const { startRow, endRow, row, rows, count, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('showRows', { startRow, endRow, row, rows, count, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `显示行失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `显示行完成！已显示第${startRow}行到第${endRow}行` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `显示行出错: ${errMsg}` }], error: errMsg };
    }
};
exports.showRowsHandler = showRowsHandler;
/**
 * 显示列
 */
exports.showColumnsDefinition = {
    name: 'wps_excel_show_columns',
    description: '显示Excel中已隐藏的列。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            startColumn: { type: 'string', description: '起始列字母，如 A' },
            endColumn: { type: 'string', description: '结束列字母，如 D' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['startColumn', 'endColumn'],
    },
};
const showColumnsHandler = async (args) => {
    const { startColumn, endColumn, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('showColumns', { startColumn, endColumn, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `显示列失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `显示列完成！已显示${startColumn}列到${endColumn}列` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `显示列出错: ${errMsg}` }], error: errMsg };
    }
};
exports.showColumnsHandler = showColumnsHandler;
/**
 * 分组行
 */
exports.groupRowsDefinition = {
    name: 'wps_excel_group_rows',
    description: '对Excel中指定范围的行进行分组，便于折叠/展开管理。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            startRow: { type: 'number', description: '起始行号（从1开始）' },
            endRow: { type: 'number', description: '结束行号（从1开始）' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['startRow', 'endRow'],
    },
};
const groupRowsHandler = async (args) => {
    const { startRow, endRow, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('groupRows', { startRow, endRow, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `分组行失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `分组行完成！已将第${startRow}行到第${endRow}行分组` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `分组行出错: ${errMsg}` }], error: errMsg };
    }
};
exports.groupRowsHandler = groupRowsHandler;
/**
 * 导出所有行列操作相关的Tools
 */
exports.rowColumnTools = [
    { definition: exports.insertRowsDefinition, handler: exports.insertRowsHandler },
    { definition: exports.insertColumnsDefinition, handler: exports.insertColumnsHandler },
    { definition: exports.deleteRowsDefinition, handler: exports.deleteRowsHandler },
    { definition: exports.deleteColumnsDefinition, handler: exports.deleteColumnsHandler },
    { definition: exports.hideRowsDefinition, handler: exports.hideRowsHandler },
    { definition: exports.showRowsDefinition, handler: exports.showRowsHandler },
    { definition: exports.showColumnsDefinition, handler: exports.showColumnsHandler },
    { definition: exports.groupRowsDefinition, handler: exports.groupRowsHandler },
];
exports.default = exports.rowColumnTools;
//# sourceMappingURL=row-column.js.map
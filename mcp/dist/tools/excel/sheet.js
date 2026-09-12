"use strict";
/**
 * Input: 工作表管理工具参数
 * Output: 工作表操作结果
 * Pos: Excel 工作表管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel工作表管理Tools
 * 工作表的创建、删除、重命名、复制、移动、切换等操作
 *
 * 包含：
 * - wps_excel_create_sheet: 创建新工作表
 * - wps_excel_delete_sheet: 删除指定工作表
 * - wps_excel_rename_sheet: 重命名工作表
 * - wps_excel_copy_sheet: 复制工作表
 * - wps_excel_get_sheet_list: 获取工作表列表
 * - wps_excel_switch_sheet: 切换工作表
 * - wps_excel_move_sheet: 移动工作表
 * - wps_excel_get_selection: 获取当前选中区域
 * - wps_excel_delete_row: 删除指定行
 * - wps_excel_insert_column: 插入列
 * - wps_excel_delete_column: 删除指定列
 * - wps_excel_freeze_panes: 冻结/取消冻结窗格
 * - wps_excel_auto_fill: 自动填充单元格区域
 * - wps_excel_set_named_range: 设置命名范围
 * - wps_excel_hide_column: 隐藏/显示列
 * - wps_excel_auto_sum: 对指定列/行自动求和
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.sheetTools = exports.autoSumHandler = exports.autoSumDefinition = exports.hideColumnHandler = exports.hideColumnDefinition = exports.setNamedRangeHandler = exports.setNamedRangeDefinition = exports.autoFillHandler = exports.autoFillDefinition = exports.freezePanesHandler = exports.freezePanesDefinition = exports.deleteColumnHandler = exports.deleteColumnDefinition = exports.insertColumnHandler = exports.insertColumnDefinition = exports.deleteRowHandler = exports.deleteRowDefinition = exports.getSelectionHandler = exports.getSelectionDefinition = exports.moveSheetHandler = exports.moveSheetDefinition = exports.switchSheetHandler = exports.switchSheetDefinition = exports.getSheetListHandler = exports.getSheetListDefinition = exports.copySheetHandler = exports.copySheetDefinition = exports.renameSheetHandler = exports.renameSheetDefinition = exports.deleteSheetHandler = exports.deleteSheetDefinition = exports.createSheetHandler = exports.createSheetDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 创建新工作表
 */
exports.createSheetDefinition = {
    name: 'wps_excel_create_sheet',
    description: '在当前工作簿中创建新的工作表。可指定名称和插入位置。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            name: {
                type: 'string',
                description: '新工作表的名称',
            },
            position: {
                type: 'number',
                description: '插入位置索引（从0开始），不填则添加到末尾',
            },
        },
        required: ['name'],
    },
};
const createSheetHandler = async (args) => {
    const { name, position } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('createSheet', { name, position }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success || !response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `创建工作表失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `工作表创建成功！\n名称: ${response.data.name}\n位置: 第${response.data.index + 1}个`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `创建工作表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.createSheetHandler = createSheetHandler;
/**
 * 删除指定工作表
 */
exports.deleteSheetDefinition = {
    name: 'wps_excel_delete_sheet',
    description: '删除当前工作簿中的指定工作表。注意：此操作不可撤销。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            name: {
                type: 'string',
                description: '要删除的工作表名称',
            },
        },
        required: ['name'],
    },
};
const deleteSheetHandler = async (args) => {
    const { name } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteSheet', { name }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `删除工作表失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `工作表 "${name}" 已成功删除`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `删除工作表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.deleteSheetHandler = deleteSheetHandler;
/**
 * 重命名工作表
 */
exports.renameSheetDefinition = {
    name: 'wps_excel_rename_sheet',
    description: '重命名当前工作簿中的指定工作表。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            oldName: {
                type: 'string',
                description: '当前工作表名称',
            },
            newName: {
                type: 'string',
                description: '新的工作表名称',
            },
        },
        required: ['oldName', 'newName'],
    },
};
const renameSheetHandler = async (args) => {
    const { oldName, newName } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('renameSheet', { oldName, newName }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `重命名工作表失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `工作表重命名成功！\n"${oldName}" → "${newName}"`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `重命名工作表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.renameSheetHandler = renameSheetHandler;
/**
 * 复制工作表
 */
exports.copySheetDefinition = {
    name: 'wps_excel_copy_sheet',
    description: '复制当前工作簿中的指定工作表。可指定新名称和插入位置。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            name: {
                type: 'string',
                description: '要复制的工作表名称',
            },
            newName: {
                type: 'string',
                description: '复制后的工作表名称，不填则自动生成',
            },
            position: {
                type: 'number',
                description: '插入位置索引（从0开始），不填则添加到末尾',
            },
        },
        required: ['name'],
    },
};
const copySheetHandler = async (args) => {
    const { name, newName, position } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('copySheet', { name, newName, position }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success || !response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `复制工作表失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `工作表复制成功！\n源工作表: ${response.data.sourceName}\n新工作表: ${response.data.newName}\n位置: 第${response.data.index + 1}个`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `复制工作表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.copySheetHandler = copySheetHandler;
/**
 * 获取工作表列表
 */
exports.getSheetListDefinition = {
    name: 'wps_excel_get_sheet_list',
    description: '获取当前工作簿的所有工作表列表，包含名称、索引和是否为活动工作表。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {},
    },
};
const getSheetListHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSheetList', {}, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success || !response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取工作表列表失败: ${response.error}` }],
                error: response.error,
            };
        }
        const { sheets, count } = response.data;
        let output = `当前工作簿共有 ${count} 个工作表：\n\n`;
        sheets.forEach((sheet) => {
            const activeFlag = sheet.active ? ' [活动]' : '';
            output += `${sheet.index + 1}. ${sheet.name}${activeFlag}\n`;
        });
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [{ type: 'text', text: output }],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取工作表列表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getSheetListHandler = getSheetListHandler;
/**
 * 切换到指定工作表
 */
exports.switchSheetDefinition = {
    name: 'wps_excel_switch_sheet',
    description: '切换到指定的工作表，使其成为活动工作表。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            name: {
                type: 'string',
                description: '要切换到的工作表名称',
            },
        },
        required: ['name'],
    },
};
const switchSheetHandler = async (args) => {
    const { name } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('switchSheet', { name }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `切换工作表失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `已切换到工作表 "${name}"`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `切换工作表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.switchSheetHandler = switchSheetHandler;
/**
 * 移动工作表到指定位置
 */
exports.moveSheetDefinition = {
    name: 'wps_excel_move_sheet',
    description: '移动指定工作表到新的位置。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            name: {
                type: 'string',
                description: '要移动的工作表名称',
            },
            position: {
                type: 'number',
                description: '目标位置索引（从0开始）',
            },
        },
        required: ['name', 'position'],
    },
};
const moveSheetHandler = async (args) => {
    const { name, position } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('moveSheet', { name, position }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `移动工作表失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `工作表 "${name}" 已移动到第${position + 1}个位置`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `移动工作表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.moveSheetHandler = moveSheetHandler;
/**
 * 获取当前选中区域信息
 */
exports.getSelectionDefinition = {
    name: 'wps_excel_get_selection',
    description: '获取当前Excel中选中区域的信息，包括范围地址、行列数等。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {},
    },
};
const getSelectionHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSelection', {}, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success || !response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取选中区域失败: ${response.error}` }],
                error: response.error,
            };
        }
        const { address, rowCount, columnCount, sheet } = response.data;
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `当前选中区域信息：\n工作表: ${sheet}\n范围: ${address}\n行数: ${rowCount}\n列数: ${columnCount}`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取选中区域出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getSelectionHandler = getSelectionHandler;
/**
 * 删除指定行
 */
exports.deleteRowDefinition = {
    name: 'wps_excel_delete_row',
    description: '删除指定行。可指定起始行号和删除行数。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            row: {
                type: 'number',
                description: '要删除的起始行号（从1开始）',
            },
            count: {
                type: 'number',
                description: '要删除的行数，默认1',
            },
        },
        required: ['row'],
    },
};
const deleteRowHandler = async (args) => {
    const { row, count = 1 } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteRows', { row, count }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `删除行失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `已成功删除第${row}行起共${count}行`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `删除行出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.deleteRowHandler = deleteRowHandler;
/**
 * 插入列
 */
exports.insertColumnDefinition = {
    name: 'wps_excel_insert_column',
    description: '在指定位置插入列。可指定起始列号和插入列数。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            column: {
                type: 'number',
                description: '要插入的列号（从1开始）',
            },
            count: {
                type: 'number',
                description: '要插入的列数，默认1',
            },
        },
        required: ['column'],
    },
};
const insertColumnHandler = async (args) => {
    const { column, count = 1 } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertColumns', { column, count }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `插入列失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `已在第${column}列处成功插入${count}列`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `插入列出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertColumnHandler = insertColumnHandler;
/**
 * 删除指定列
 */
exports.deleteColumnDefinition = {
    name: 'wps_excel_delete_column',
    description: '删除指定列。可指定起始列号和删除列数。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            column: {
                type: 'number',
                description: '要删除的起始列号（从1开始）',
            },
            count: {
                type: 'number',
                description: '要删除的列数，默认1',
            },
        },
        required: ['column'],
    },
};
const deleteColumnHandler = async (args) => {
    const { column, count = 1 } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteColumns', { column, count }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `删除列失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `已成功删除第${column}列起共${count}列`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `删除列出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.deleteColumnHandler = deleteColumnHandler;
/**
 * 冻结/取消冻结窗格
 */
exports.freezePanesDefinition = {
    name: 'wps_excel_freeze_panes',
    description: '冻结/取消冻结窗格。可指定冻结的行和列位置。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            row: {
                type: 'number',
                description: '冻结到第几行（从1开始），不填则不冻结行',
            },
            column: {
                type: 'number',
                description: '冻结到第几列（从1开始），不填则不冻结列',
            },
            freeze: {
                type: 'boolean',
                description: '是否冻结，默认true。设为false则取消冻结',
            },
        },
    },
};
const freezePanesHandler = async (args) => {
    const { row, column, freeze = true } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('freezePanes', { row, column, freeze }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `冻结窗格操作失败: ${response.error}` }],
                error: response.error,
            };
        }
        const action = freeze ? '冻结' : '取消冻结';
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `窗格${action}成功！${row ? `\n冻结行: ${row}` : ''}${column ? `\n冻结列: ${column}` : ''}`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `冻结窗格操作出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.freezePanesHandler = freezePanesHandler;
/**
 * 自动填充单元格区域
 */
exports.autoFillDefinition = {
    name: 'wps_excel_auto_fill',
    description: '自动填充单元格区域。根据源区域的数据模式自动填充到目标区域。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            sourceRange: {
                type: 'string',
                description: '源数据区域，如 "A1:A5"',
            },
            targetRange: {
                type: 'string',
                description: '目标填充区域，如 "A1:A20"',
            },
        },
        required: ['sourceRange', 'targetRange'],
    },
};
const autoFillHandler = async (args) => {
    const { sourceRange, targetRange } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('fillSeries', { sourceRange, targetRange }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `自动填充失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `自动填充成功！\n源区域: ${sourceRange}\n目标区域: ${targetRange}`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `自动填充出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.autoFillHandler = autoFillHandler;
/**
 * 设置命名范围
 */
exports.setNamedRangeDefinition = {
    name: 'wps_excel_set_named_range',
    description: '设置命名范围。为指定单元格区域创建或更新命名范围。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            name: {
                type: 'string',
                description: '命名范围的名称',
            },
            range: {
                type: 'string',
                description: '单元格区域，如 "A1:D10"',
            },
        },
        required: ['name', 'range'],
    },
};
const setNamedRangeHandler = async (args) => {
    const { name, range } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('createNamedRange', { name, range }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置命名范围失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `命名范围设置成功！\n名称: ${name}\n范围: ${range}`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置命名范围出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setNamedRangeHandler = setNamedRangeHandler;
/**
 * 隐藏/显示列
 */
exports.hideColumnDefinition = {
    name: 'wps_excel_hide_column',
    description: '隐藏或显示指定列。可指定起始列号、列数和隐藏/显示状态。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            column: {
                type: 'number',
                description: '起始列号（从1开始）',
            },
            count: {
                type: 'number',
                description: '列数，默认1',
            },
            hide: {
                type: 'boolean',
                description: '是否隐藏，true为隐藏，false为显示',
            },
        },
        required: ['column', 'count', 'hide'],
    },
};
const hideColumnHandler = async (args) => {
    const { column, count = 1, hide } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('hideColumns', { column, count, hide }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `${hide ? '隐藏' : '显示'}列失败: ${response.error}` }],
                error: response.error,
            };
        }
        const action = hide ? '隐藏' : '显示';
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `已成功${action}第${column}列起共${count}列`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `隐藏/显示列出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.hideColumnHandler = hideColumnHandler;
/**
 * 对指定列/行自动求和
 */
exports.autoSumDefinition = {
    name: 'wps_excel_auto_sum',
    description: '对指定范围的列或行自动求和，并将结果写入目标单元格。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: {
                type: 'string',
                description: '要求和的数据范围，如 "A1:A10" 或 "B2:F2"',
            },
            targetCell: {
                type: 'string',
                description: '求和结果写入的目标单元格，如 "A11" 或 "G2"',
            },
        },
        required: ['range', 'targetCell'],
    },
};
const autoSumHandler = async (args) => {
    const { range, targetCell } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('autoSum', // NOTE: macOS未实现，仅Windows支持
        { range, targetCell }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `自动求和失败: ${response.error}` }],
                error: response.error,
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `自动求和成功！\n求和范围: ${range}\n结果写入: ${targetCell}`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `自动求和出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.autoSumHandler = autoSumHandler;
/**
 * 导出所有工作表管理相关的Tools
 */
exports.sheetTools = [
    { definition: exports.createSheetDefinition, handler: exports.createSheetHandler },
    { definition: exports.deleteSheetDefinition, handler: exports.deleteSheetHandler },
    { definition: exports.renameSheetDefinition, handler: exports.renameSheetHandler },
    { definition: exports.copySheetDefinition, handler: exports.copySheetHandler },
    { definition: exports.getSheetListDefinition, handler: exports.getSheetListHandler },
    { definition: exports.switchSheetDefinition, handler: exports.switchSheetHandler },
    { definition: exports.moveSheetDefinition, handler: exports.moveSheetHandler },
    { definition: exports.getSelectionDefinition, handler: exports.getSelectionHandler },
    { definition: exports.deleteRowDefinition, handler: exports.deleteRowHandler },
    { definition: exports.insertColumnDefinition, handler: exports.insertColumnHandler },
    { definition: exports.deleteColumnDefinition, handler: exports.deleteColumnHandler },
    { definition: exports.freezePanesDefinition, handler: exports.freezePanesHandler },
    { definition: exports.autoFillDefinition, handler: exports.autoFillHandler },
    { definition: exports.setNamedRangeDefinition, handler: exports.setNamedRangeHandler },
    { definition: exports.hideColumnDefinition, handler: exports.hideColumnHandler },
    { definition: exports.autoSumDefinition, handler: exports.autoSumHandler },
];
exports.default = exports.sheetTools;
//# sourceMappingURL=sheet.js.map
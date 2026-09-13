"use strict";
/**
 * Input: 工作簿管理工具参数
 * Output: 工作簿操作结果
 * Pos: Excel 工作簿管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.workbookTools = exports.clearRangeHandler = exports.clearRangeDefinition = exports.getCellInfoHandler = exports.getCellInfoDefinition = exports.getFormulaHandler = exports.getFormulaDefinition = exports.setCellValueHandler = exports.setCellValueDefinition = exports.getCellValueHandler = exports.getCellValueDefinition = exports.createWorkbookHandler = exports.createWorkbookDefinition = exports.closeWorkbookHandler = exports.closeWorkbookDefinition = exports.switchWorkbookHandler = exports.switchWorkbookDefinition = exports.getOpenWorkbooksHandler = exports.getOpenWorkbooksDefinition = exports.openWorkbookHandler = exports.openWorkbookDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 打开指定路径的Excel工作簿
 */
exports.openWorkbookDefinition = {
    name: 'wps_excel_open_workbook',
    description: '打开指定路径的Excel工作簿文件。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            filePath: {
                type: 'string',
                description: '工作簿文件路径',
            },
        },
        required: ['filePath'],
    },
};
const openWorkbookHandler = async (args) => {
    const { filePath } = args;
    if (!filePath) {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '请提供工作簿文件路径' }],
            error: '缺少文件路径',
        };
    }
    try {
        // The bridge reads "path"; filePath was never read, so a relative or alternate spelling of the
        // argument was silently ignored.
        const params = { path: filePath };
        const response = await wps_client_1.wpsClient.executeMethod('openWorkbook', params, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `打开工作簿失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `工作簿已打开: ${filePath}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `打开工作簿出错: ${errMsg}` }], error: errMsg };
    }
};
exports.openWorkbookHandler = openWorkbookHandler;
/**
 * 获取所有已打开的工作簿列表
 */
exports.getOpenWorkbooksDefinition = {
    name: 'wps_excel_get_open_workbooks',
    description: '获取当前所有已打开的Excel工作簿列表。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const getOpenWorkbooksHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getOpenWorkbooks', {}, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取工作簿列表失败: ${response.error}` }], error: response.error };
        }
        // The bridge returns one object per workbook ({name, path, sheets, active}); joining them
        // straight into the text printed "[object Object]" for every open workbook.
        const list = (response.data?.workbooks || []).map((entry) => {
            if (typeof entry === 'string')
                return entry;
            const parts = [entry?.name || '(未命名工作簿)'];
            // An unsaved workbook reports its name as its FullName; showing it twice reads like a bug.
            if (entry?.path && entry.path !== entry.name)
                parts.push(entry.path);
            if (typeof entry?.sheets === 'number')
                parts.push(`${entry.sheets} 个工作表`);
            if (entry?.active)
                parts.push('当前活动');
            return parts.join(' | ');
        });
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `已打开的工作簿 (${list.length}个):\n${list.join('\n') || '无'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取工作簿列表出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getOpenWorkbooksHandler = getOpenWorkbooksHandler;
/**
 * 切换到指定名称的工作簿
 */
exports.switchWorkbookDefinition = {
    name: 'wps_excel_switch_workbook',
    description: '切换到指定名称的Excel工作簿。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            name: {
                type: 'string',
                description: '工作簿名称',
            },
        },
        required: ['name'],
    },
};
const switchWorkbookHandler = async (args) => {
    const { name } = args;
    if (!name) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '请提供工作簿名称' }], error: '缺少工作簿名称' };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('switchWorkbook', { name }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `切换工作簿失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `已切换到工作簿: ${name}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `切换工作簿出错: ${errMsg}` }], error: errMsg };
    }
};
exports.switchWorkbookHandler = switchWorkbookHandler;
/**
 * 关闭指定工作簿
 */
exports.closeWorkbookDefinition = {
    name: 'wps_excel_close_workbook',
    description: '关闭指定的Excel工作簿，可选是否保存。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            name: {
                type: 'string',
                description: '工作簿名称，不填则关闭当前工作簿',
            },
            save: {
                type: 'boolean',
                description: '是否保存，默认true',
            },
        },
        required: [],
    },
};
const closeWorkbookHandler = async (args) => {
    const { name, save } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('closeWorkbook', { name, save: save !== false }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `关闭工作簿失败: ${response.error}` }], error: response.error };
        }
        // Report what actually happened, not what was asked for: closing a never-saved workbook
        // with save=true drops the save instead of raising a modal Save As dialog.
        const closed = response.data?.closed ?? name ?? '(当前工作簿)';
        const savedNote = response.data?.saved ? '（已保存）' : '（未保存）';
        const warnNote = response.data?.warning ? `\n注意: ${response.data.warning}` : '';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `工作簿已关闭: ${closed}${savedNote}${warnNote}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `关闭工作簿出错: ${errMsg}` }], error: errMsg };
    }
};
exports.closeWorkbookHandler = closeWorkbookHandler;
/**
 * 新建空白工作簿
 */
exports.createWorkbookDefinition = {
    name: 'wps_excel_create_workbook',
    description: '新建一个空白Excel工作簿。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const createWorkbookHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('createWorkbook', {}, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `新建工作簿失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '新工作簿已创建' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `新建工作簿出错: ${errMsg}` }], error: errMsg };
    }
};
exports.createWorkbookHandler = createWorkbookHandler;
/**
 * 获取指定单元格的值
 */
exports.getCellValueDefinition = {
    name: 'wps_excel_get_cell_value',
    description: '获取Excel指定单元格的值。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            sheet: { type: 'string', description: '工作表名称' },
            row: { type: 'number', description: '行号（从1开始）' },
            col: { type: 'number', description: '列号（从1开始）' },
        },
        required: ['sheet', 'row', 'col'],
    },
};
const getCellValueHandler = async (args) => {
    const { sheet, row, col } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getCellValue', { sheet, row, col }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取单元格值失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `单元格值: ${JSON.stringify(response.data?.value)}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取单元格值出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getCellValueHandler = getCellValueHandler;
/**
 * 设置指定单元格的值
 */
exports.setCellValueDefinition = {
    name: 'wps_excel_set_cell_value',
    description: '设置Excel指定单元格的值。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            sheet: { type: 'string', description: '工作表名称' },
            row: { type: 'number', description: '行号（从1开始）' },
            col: { type: 'number', description: '列号（从1开始）' },
            value: { type: 'string', description: '要设置的值' },
        },
        required: ['sheet', 'row', 'col', 'value'],
    },
};
const setCellValueHandler = async (args) => {
    const { sheet, row, col, value } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setCellValue', { sheet, row, col, value }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置单元格值失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `单元格值已设置: ${sheet}!R${row}C${col} = ${value}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置单元格值出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setCellValueHandler = setCellValueHandler;
/**
 * 获取指定单元格的公式
 */
exports.getFormulaDefinition = {
    name: 'wps_excel_get_formula',
    description: '获取Excel指定单元格的公式。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            sheet: { type: 'string', description: '工作表名称' },
            cell: { type: 'string', description: '单元格地址，如 A1、B2' },
        },
        required: ['sheet', 'cell'],
    },
};
const getFormulaHandler = async (args) => {
    const { sheet, cell } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getFormula', { sheet, cell }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取公式失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `单元格 ${cell} 的公式: ${response.data?.formula || '无公式'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取公式出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getFormulaHandler = getFormulaHandler;
/**
 * 获取单元格详细信息
 */
exports.getCellInfoDefinition = {
    name: 'wps_excel_get_cell_info',
    description: '获取单元格的详细信息（值、公式、格式等）。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            sheet: { type: 'string', description: '工作表名称' },
            cell: { type: 'string', description: '单元格地址，如 A1' },
        },
        required: ['sheet', 'cell'],
    },
};
const getCellInfoHandler = async (args) => {
    const { sheet, cell } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getCellInfo', { sheet, cell }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取单元格信息失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `单元格 ${cell} 信息:\n${JSON.stringify(response.data, null, 2)}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取单元格信息出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getCellInfoHandler = getCellInfoHandler;
/**
 * 清除指定范围
 */
exports.clearRangeDefinition = {
    name: 'wps_excel_clear_range',
    description: '清除指定范围的内容、格式或全部。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '范围地址，如 A1:C10' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前工作表' },
            type: {
                type: 'string',
                description: '清除类型：all（全部）、contents（仅内容）、formats（仅格式）',
                enum: ['all', 'contents', 'formats'],
            },
        },
        required: ['range'],
    },
};
const clearRangeHandler = async (args) => {
    const { range, sheet, type } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('clearRange', { range, sheet, type: type || 'all' }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `清除范围失败: ${response.error}` }], error: response.error };
        }
        const typeLabel = type === 'contents' ? '内容' : type === 'formats' ? '格式' : '全部';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `范围 ${range} 的${typeLabel}已清除` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `清除范围出错: ${errMsg}` }], error: errMsg };
    }
};
exports.clearRangeHandler = clearRangeHandler;
/**
 * 导出所有工作簿管理相关的Tools
 */
exports.workbookTools = [
    { definition: exports.openWorkbookDefinition, handler: exports.openWorkbookHandler },
    { definition: exports.getOpenWorkbooksDefinition, handler: exports.getOpenWorkbooksHandler },
    { definition: exports.switchWorkbookDefinition, handler: exports.switchWorkbookHandler },
    { definition: exports.closeWorkbookDefinition, handler: exports.closeWorkbookHandler },
    { definition: exports.createWorkbookDefinition, handler: exports.createWorkbookHandler },
    { definition: exports.getCellValueDefinition, handler: exports.getCellValueHandler },
    { definition: exports.setCellValueDefinition, handler: exports.setCellValueHandler },
    { definition: exports.getFormulaDefinition, handler: exports.getFormulaHandler },
    { definition: exports.getCellInfoDefinition, handler: exports.getCellInfoHandler },
    { definition: exports.clearRangeDefinition, handler: exports.clearRangeHandler },
];
exports.default = exports.workbookTools;
//# sourceMappingURL=workbook.js.map
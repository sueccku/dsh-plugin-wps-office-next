"use strict";
/**
 * Input: 批注/保护/公式/图片/超链接工具参数
 * Output: 批注删除/获取、取消保护、锁定单元格、数组公式、图片插入、超链接设置结果
 * Pos: Excel 批注保护及扩展工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel批注保护Tools - 批注、保护与扩展功能模块
 * 处理批注管理、工作表取消保护、单元格锁定、数组公式、图片插入、超链接等操作
 *
 * 包含：
 * - wps_excel_delete_cell_comment: 删除单元格批注
 * - wps_excel_get_cell_comments: 获取单元格批注
 * - wps_excel_unprotect_sheet: 取消保护工作表
 * - wps_excel_lock_cells: 锁定/解锁单元格
 * - wps_excel_set_array_formula: 设置数组公式
 * - wps_excel_insert_excel_image: 插入图片
 * - wps_excel_set_hyperlink: 设置超链接
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.commentProtectTools = exports.setHyperlinkHandler = exports.setHyperlinkDefinition = exports.insertExcelImageHandler = exports.insertExcelImageDefinition = exports.setArrayFormulaHandler = exports.setArrayFormulaDefinition = exports.lockCellsHandler = exports.lockCellsDefinition = exports.unprotectSheetHandler = exports.unprotectSheetDefinition = exports.getCellCommentsHandler = exports.getCellCommentsDefinition = exports.deleteCellCommentHandler = exports.deleteCellCommentDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 删除单元格批注
 */
exports.deleteCellCommentDefinition = {
    name: 'wps_excel_delete_cell_comment',
    description: '删除Excel单元格上的批注。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            cell: { type: 'string', description: '单元格地址，如 A1、B2' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['cell'],
    },
};
const deleteCellCommentHandler = async (args) => {
    const { cell, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteCellComment', { cell, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `删除批注失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `批注删除成功！单元格: ${cell}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `删除批注出错: ${errMsg}` }], error: errMsg };
    }
};
exports.deleteCellCommentHandler = deleteCellCommentHandler;
/**
 * 获取单元格批注
 */
exports.getCellCommentsDefinition = {
    name: 'wps_excel_get_cell_comments',
    description: '获取Excel指定范围内的所有批注。不指定范围则获取当前工作表所有批注。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '要查询的范围，如 A1:D10。不填则获取所有批注' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: [],
    },
};
const getCellCommentsHandler = async (args) => {
    const { range, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getCellComments', { range, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success || !response.data) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取批注失败: ${response.error}` }], error: response.error };
        }
        const comments = response.data.comments;
        if (!comments || comments.length === 0) {
            return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `${range ? `范围 ${range}` : '当前工作表'}没有批注` }] };
        }
        let output = `找到${comments.length}条批注：\n\n`;
        comments.forEach((c) => {
            output += `- ${c.cell}: ${c.comment}${c.author ? ` (作者: ${c.author})` : ''}\n`;
        });
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: output }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取批注出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getCellCommentsHandler = getCellCommentsHandler;
/**
 * 取消保护工作表
 */
exports.unprotectSheetDefinition = {
    name: 'wps_excel_unprotect_sheet',
    description: '取消保护当前工作表。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            password: { type: 'string', description: '保护密码（如果设置了密码保护则需要提供）' },
        },
        required: [],
    },
};
const unprotectSheetHandler = async (args) => {
    const { password } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('unprotectSheet', { password }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `取消保护工作表失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '工作表取消保护成功！' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `取消保护工作表出错: ${errMsg}` }], error: errMsg };
    }
};
exports.unprotectSheetHandler = unprotectSheetHandler;
/**
 * 锁定/解锁单元格
 */
exports.lockCellsDefinition = {
    name: 'wps_excel_lock_cells',
    description: '锁定或解锁Excel指定范围的单元格（需配合工作表保护使用）。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '要锁定/解锁的范围，如 A1:D10' },
            locked: { type: 'boolean', description: 'true为锁定，false为解锁' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['range', 'locked'],
    },
};
const lockCellsHandler = async (args) => {
    const { range, locked, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('lockCells', { range, locked, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `${locked ? '锁定' : '解锁'}单元格失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `单元格${locked ? '锁定' : '解锁'}成功！范围: ${range}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `${locked ? '锁定' : '解锁'}单元格出错: ${errMsg}` }], error: errMsg };
    }
};
exports.lockCellsHandler = lockCellsHandler;
/**
 * 设置数组公式
 */
exports.setArrayFormulaDefinition = {
    name: 'wps_excel_set_array_formula',
    description: '为Excel指定范围设置数组公式（CSE数组公式）。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '数组公式应用的范围，如 A1:A10' },
            formula: { type: 'string', description: '数组公式，如 =A1:A10*B1:B10' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['range', 'formula'],
    },
};
const setArrayFormulaHandler = async (args) => {
    const { range, formula, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setArrayFormula', { range, formula, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置数组公式失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `数组公式设置成功！范围: ${range}，公式: ${formula}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置数组公式出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setArrayFormulaHandler = setArrayFormulaHandler;
/**
 * 插入图片到Excel
 */
exports.insertExcelImageDefinition = {
    name: 'wps_excel_insert_excel_image',
    description: '在Excel中插入图片到指定位置。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            filePath: { type: 'string', description: '图片文件路径' },
            cell: { type: 'string', description: '插入位置的单元格地址，如 A1。不填则插入到当前选中位置' },
            width: { type: 'number', description: '图片宽度（像素），不填则使用原始宽度' },
            height: { type: 'number', description: '图片高度（像素），不填则使用原始高度' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['filePath'],
    },
};
const insertExcelImageHandler = async (args) => {
    const { filePath, cell, width, height, sheet } = args;
    try {
        // 跨平台参数对齐：macOS/Windows 底层均读取 params.path，需同时发送 path/imagePath 别名
        const response = await wps_client_1.wpsClient.executeMethod('insertExcelImage', { filePath, path: filePath, imagePath: filePath, cell, width, height, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入图片失败: ${response.error}` }], error: response.error };
        }
        let text = `图片插入成功！文件: ${filePath}`;
        if (cell)
            text += `，位置: ${cell}`;
        if (width || height)
            text += `，尺寸: ${width || '自动'}x${height || '自动'}`;
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入图片出错: ${errMsg}` }], error: errMsg };
    }
};
exports.insertExcelImageHandler = insertExcelImageHandler;
/**
 * 设置超链接
 */
exports.setHyperlinkDefinition = {
    name: 'wps_excel_set_hyperlink',
    description: '为Excel单元格设置超链接。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            cell: { type: 'string', description: '单元格地址，如 A1' },
            url: { type: 'string', description: '超链接URL地址' },
            text: { type: 'string', description: '显示文本，不填则显示URL' },
            sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
        },
        required: ['cell', 'url'],
    },
};
const setHyperlinkHandler = async (args) => {
    const { cell, url, text, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setHyperlink', { cell, url, text, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置超链接失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `超链接设置成功！单元格: ${cell}，URL: ${url}${text ? `，显示文本: ${text}` : ''}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置超链接出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setHyperlinkHandler = setHyperlinkHandler;
/**
 * 导出所有批注保护相关的Tools
 */
exports.commentProtectTools = [
    { definition: exports.deleteCellCommentDefinition, handler: exports.deleteCellCommentHandler },
    { definition: exports.getCellCommentsDefinition, handler: exports.getCellCommentsHandler },
    { definition: exports.unprotectSheetDefinition, handler: exports.unprotectSheetHandler },
    { definition: exports.lockCellsDefinition, handler: exports.lockCellsHandler },
    { definition: exports.setArrayFormulaDefinition, handler: exports.setArrayFormulaHandler },
    { definition: exports.insertExcelImageDefinition, handler: exports.insertExcelImageHandler },
    { definition: exports.setHyperlinkDefinition, handler: exports.setHyperlinkHandler },
];
exports.default = exports.commentProtectTools;
//# sourceMappingURL=comment-protect.js.map
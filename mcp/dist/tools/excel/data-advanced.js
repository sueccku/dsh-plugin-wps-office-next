"use strict";
/**
 * Input: 高级数据处理工具参数
 * Output: 数据处理结果
 * Pos: Excel 高级数据处理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.dataAdvancedTools = exports.subtotalHandler = exports.subtotalDefinition = exports.textToColumnsHandler = exports.textToColumnsDefinition = exports.transposeHandler = exports.transposeDefinition = exports.fillSeriesHandler = exports.fillSeriesDefinition = exports.pasteRangeHandler = exports.pasteRangeDefinition = exports.copyRangeHandler = exports.copyRangeDefinition = exports.autoFilterHandler = exports.autoFilterDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 自动筛选
 */
exports.autoFilterDefinition = {
    name: 'wps_excel_auto_filter',
    description: '对Excel指定范围应用自动筛选。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '筛选范围，如 A1:D100' },
            column: { type: 'string', description: '筛选列标识' },
            criteria: { type: 'string', description: '筛选条件' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['range'],
    },
};
const autoFilterHandler = async (args) => {
    const { range, column, criteria, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('autoFilter', { range, column, criteria, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `自动筛选失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `自动筛选已应用: ${range}${column ? '，列: ' + column : ''}${criteria ? '，条件: ' + criteria : ''}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `自动筛选出错: ${errMsg}` }], error: errMsg };
    }
};
exports.autoFilterHandler = autoFilterHandler;
/**
 * 复制范围
 */
exports.copyRangeDefinition = {
    name: 'wps_excel_copy_range',
    description: '复制Excel指定范围到目标位置。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            source: { type: 'string', description: '源范围，如 A1:C10' },
            destination: { type: 'string', description: '目标位置，如 E1' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['source', 'destination'],
    },
};
const copyRangeHandler = async (args) => {
    const { source, destination, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('copyRange', { source, destination, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `复制范围失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `已复制 ${source} 到 ${destination}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `复制范围出错: ${errMsg}` }], error: errMsg };
    }
};
exports.copyRangeHandler = copyRangeHandler;
/**
 * 粘贴范围
 */
exports.pasteRangeDefinition = {
    name: 'wps_excel_paste_range',
    description: '粘贴已复制的内容到指定位置。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            destination: { type: 'string', description: '目标位置，如 E1' },
            pasteType: {
                type: 'string',
                description: '粘贴类型',
                enum: ['all', 'values', 'formats', 'formulas'],
            },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['destination'],
    },
};
const pasteRangeHandler = async (args) => {
    const { destination, pasteType, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('pasteRange', { destination, pasteType: pasteType || 'all', sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `粘贴失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `已粘贴到 ${destination}，类型: ${pasteType || 'all'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `粘贴出错: ${errMsg}` }], error: errMsg };
    }
};
exports.pasteRangeHandler = pasteRangeHandler;
/**
 * 填充序列
 */
exports.fillSeriesDefinition = {
    name: 'wps_excel_fill_series',
    description: '自动填充序列数据。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '填充范围，如 A1:A10' },
            direction: { type: 'string', description: '填充方向', enum: ['down', 'right', 'up', 'left'] },
            type: { type: 'string', description: '填充类型', enum: ['linear', 'growth', 'date', 'auto'] },
            step: { type: 'number', description: '步长值' },
            startValue: { type: 'number', description: '序列起始值，默认1' },
            sheet: { type: 'string', description: '工作表名称' },
            sourceRange: { type: 'string', description: '自动填充的源区域（与 targetRange 配对使用）' },
            targetRange: { type: 'string', description: '自动填充的目标区域' },
        },
        // Either range (fill a series) or sourceRange+targetRange (extend a pattern) is required; the
        // bridge rejects a call that has neither.
        required: [],
    },
};
const fillSeriesHandler = async (args) => {
    const { range, direction, type, step, sheet, startValue, sourceRange, targetRange } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('fillSeries', { range, direction: direction || 'down', type: type || 'auto', step, sheet, startValue, sourceRange, targetRange }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `填充序列失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `序列填充完成: ${range}，方向: ${direction || 'down'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `填充序列出错: ${errMsg}` }], error: errMsg };
    }
};
exports.fillSeriesHandler = fillSeriesHandler;
/**
 * 转置数据
 */
exports.transposeDefinition = {
    name: 'wps_excel_transpose',
    description: '转置数据（行列互换）。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            source: { type: 'string', description: '源范围，如 A1:C3' },
            destination: { type: 'string', description: '目标位置，如 E1' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['source', 'destination'],
    },
};
const transposeHandler = async (args) => {
    const { source, destination, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('transpose', { source, destination, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `转置失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `数据已转置: ${source} → ${destination}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `转置出错: ${errMsg}` }], error: errMsg };
    }
};
exports.transposeHandler = transposeHandler;
/**
 * 分列
 */
exports.textToColumnsDefinition = {
    name: 'wps_excel_text_to_columns',
    description: '将文本按分隔符拆分到多列。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '要拆分的范围，如 A1:A100' },
            delimiter: { type: 'string', description: '分隔符，默认逗号' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['range'],
    },
};
const textToColumnsHandler = async (args) => {
    const { range, delimiter, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('textToColumns', { range, delimiter: delimiter || ',', sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `分列失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `分列完成: ${range}，分隔符: "${delimiter || ','}"` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `分列出错: ${errMsg}` }], error: errMsg };
    }
};
exports.textToColumnsHandler = textToColumnsHandler;
/**
 * 分类汇总
 */
exports.subtotalDefinition = {
    name: 'wps_excel_subtotal',
    description: '创建分类汇总。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '数据范围，如 A1:D100' },
            groupBy: { type: 'string', description: '分组依据列在 range 内的序号（从 1 开始），如 range=A1:C4 时用 "1" 按第一列分组' },
            function: { type: 'string', description: '汇总函数', enum: ['sum', 'count', 'average', 'max', 'min'] },
            columns: {
                type: 'array',
                items: { type: 'string' },
                description: '要汇总的列在 range 内的序号列表（从 1 开始），如 [3] 表示对第三列求和',
            },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['range', 'groupBy', 'function', 'columns'],
    },
};
const subtotalHandler = async (args) => {
    const { range, groupBy, function: func, columns, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('subtotal', { range, groupBy, function: func, columns, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `分类汇总失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `分类汇总完成: 按 ${groupBy} 分组，${func} 汇总 [${columns.join(', ')}]` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `分类汇总出错: ${errMsg}` }], error: errMsg };
    }
};
exports.subtotalHandler = subtotalHandler;
/**
 * 导出所有高级数据处理相关的Tools
 */
exports.dataAdvancedTools = [
    { definition: exports.autoFilterDefinition, handler: exports.autoFilterHandler },
    { definition: exports.copyRangeDefinition, handler: exports.copyRangeHandler },
    { definition: exports.pasteRangeDefinition, handler: exports.pasteRangeHandler },
    { definition: exports.fillSeriesDefinition, handler: exports.fillSeriesHandler },
    { definition: exports.transposeDefinition, handler: exports.transposeHandler },
    { definition: exports.textToColumnsDefinition, handler: exports.textToColumnsHandler },
    { definition: exports.subtotalDefinition, handler: exports.subtotalHandler },
];
exports.default = exports.dataAdvancedTools;
//# sourceMappingURL=data-advanced.js.map
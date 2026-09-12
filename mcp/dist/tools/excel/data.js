"use strict";
/**
 * Input: 数据处理工具参数
 * Output: 读写/清洗结果
 * Pos: Excel 数据处理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel数据处理Tools - 数据读写与清洗模块
 * 处理数据的读取、写入、去重、去空格、格式统一等操作
 *
 * 包含：
 * - wps_excel_read_range: 读取指定范围数据
 * - wps_excel_write_range: 写入数据到指定范围
 * - wps_excel_clean_data: 数据清洗（核心功能）
 * - wps_excel_remove_duplicates: 删除重复行
 * - wps_excel_sort_range: 对选定区域排序
 * - wps_excel_find_replace: 查找并替换内容
 * - wps_excel_insert_row: 插入行
 * - wps_excel_add_comment: 给单元格添加批注
 * - wps_excel_protect_sheet: 保护/取消保护工作表
 * - wps_excel_set_conditional_format: 设置条件格式
 * - wps_excel_protect_workbook: 保护/取消保护工作簿
 * - wps_excel_set_zoom: 设置工作表缩放比例
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.dataTools = exports.setZoomHandler = exports.setZoomDefinition = exports.protectWorkbookHandler = exports.protectWorkbookDefinition = exports.setConditionalFormatHandler = exports.setConditionalFormatDefinition = exports.protectSheetHandler = exports.protectSheetDefinition = exports.addCommentHandler = exports.addCommentDefinition = exports.insertRowHandler = exports.insertRowDefinition = exports.findReplaceHandler = exports.findReplaceDefinition = exports.sortRangeHandler = exports.sortRangeDefinition = exports.removeDuplicatesHandler = exports.removeDuplicatesDefinition = exports.cleanDataHandler = exports.cleanDataDefinition = exports.writeRangeHandler = exports.writeRangeDefinition = exports.readRangeHandler = exports.readRangeDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 读取指定范围的单元格数据
 */
exports.readRangeDefinition = {
    name: 'wps_excel_read_range',
    description: '读取Excel指定范围的单元格数据，返回二维数组格式的数据。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: {
                type: 'string',
                description: '要读取的范围，如 A1:C10、B2:D5',
            },
            sheet: {
                type: 'string',
                description: '工作表名称，不填则使用当前活动工作表',
            },
            include_header: {
                type: 'boolean',
                description: '是否将第一行作为表头返回，默认false',
            },
        },
        required: ['range'],
    },
};
const readRangeHandler = async (args) => {
    const { range, sheet, include_header } = args;
    try {
        const response = await wps_client_1.wpsClient.getRangeData(sheet, range);
        if (!response || response.length === 0) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `范围 ${range} 是空的，没有数据`,
                    },
                ],
            };
        }
        // 格式化输出
        let output = `范围 ${range} 的数据（${response.length}行 x ${response[0]?.length || 0}列）：\n\n`;
        if (include_header && response.length > 0) {
            const headers = response[0];
            output += `表头: ${headers.join(' | ')}\n`;
            output += '-'.repeat(50) + '\n';
            for (let i = 1; i < response.length; i++) {
                output += `第${i}行: ${response[i].join(' | ')}\n`;
            }
        }
        else {
            response.forEach((row, index) => {
                output += `第${index + 1}行: ${row.join(' | ')}\n`;
            });
        }
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
            content: [{ type: 'text', text: `读取数据出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.readRangeHandler = readRangeHandler;
/**
 * 向指定范围写入数据
 * 批量写入数据，比一个个单元格设置快多了
 */
exports.writeRangeDefinition = {
    name: 'wps_excel_write_range',
    description: '向Excel指定范围写入数据。数据格式为二维数组，从指定单元格开始向右下方填充。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: {
                type: 'string',
                description: '起始单元格地址，如 A1、B2',
            },
            data: {
                type: 'array',
                description: '二维数组数据，如 [["姓名","年龄"],["张三",25],["李四",30]]',
                items: {
                    type: 'array',
                    items: {
                        type: 'string',
                    },
                },
            },
            sheet: {
                type: 'string',
                description: '工作表名称，不填则使用当前活动工作表',
            },
        },
        required: ['range', 'data'],
    },
};
const writeRangeHandler = async (args) => {
    const { range, data, sheet } = args;
    // 数据校验
    if (!Array.isArray(data) || data.length === 0) {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '写入数据不能为空，请提供有效的二维数组数据' }],
            error: '数据为空',
        };
    }
    try {
        const success = await wps_client_1.wpsClient.setRangeData(sheet, range, data);
        if (success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `数据写入成功！\n起始位置: ${range}\n写入规模: ${data.length}行 x ${data[0]?.length || 0}列`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: '写入数据失败' }],
                error: '写入失败',
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `写入数据出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.writeRangeHandler = writeRangeHandler;
/**
 * 数据清洗工具
 * 一键处理脏数据，支持多种清洗操作组合
 */
exports.cleanDataDefinition = {
    name: 'wps_excel_clean_data',
    description: `数据清洗工具，支持多种清洗操作的组合。

使用场景：
- "把A列的前后空格去掉" -> 使用 trim 操作
- "把日期格式统一成年-月-日" -> 使用 unify_date 操作
- "删除空行" -> 使用 remove_empty_rows 操作

支持的操作：
- trim: 去除单元格前后空格
- remove_duplicates: 删除重复行
- unify_date: 统一日期格式为 yyyy-mm-dd
- remove_empty_rows: 删除空行`,
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: {
                type: 'string',
                description: '要清洗的数据范围，如 A1:D100',
            },
            operations: {
                type: 'array',
                description: '要执行的清洗操作列表',
                items: {
                    type: 'string',
                    enum: ['trim', 'remove_duplicates', 'unify_date', 'remove_empty_rows'],
                },
            },
            sheet: {
                type: 'string',
                description: '工作表名称，不填则使用当前活动工作表',
            },
        },
        required: ['range', 'operations'],
    },
};
const cleanDataHandler = async (args) => {
    const { range, operations, sheet } = args;
    // 校验操作列表
    const validOperations = ['trim', 'remove_duplicates', 'unify_date', 'remove_empty_rows'];
    const invalidOps = operations.filter((op) => !validOperations.includes(op));
    if (invalidOps.length > 0) {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [
                {
                    type: 'text',
                    text: `不支持的操作: ${invalidOps.join(', ')}\n支持的操作: ${validOperations.join(', ')}`,
                },
            ],
            error: '无效的操作类型',
        };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('cleanData', { range, operations, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success || !response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `数据清洗失败: ${response.error}` }],
                error: response.error,
            };
        }
        const result = response.data;
        let output = `数据清洗完成！\n范围: ${result.range}\n\n操作结果：\n`;
        result.operations.forEach((op) => {
            const status = op.success ? '成功' : '失败';
            output += `- ${op.operation}: ${status} - ${op.message}\n`;
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
            content: [{ type: 'text', text: `数据清洗出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.cleanDataHandler = cleanDataHandler;
/**
 * 删除重复行
 * 单独拎出来，因为这个功能用得太多了
 */
exports.removeDuplicatesDefinition = {
    name: 'wps_excel_remove_duplicates',
    description: '删除指定范围内的重复行。可以指定根据哪些列判断重复。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: {
                type: 'string',
                description: '要处理的数据范围，如 A1:D100',
            },
            columns: {
                type: 'array',
                description: '用于判断重复的列，如 ["A", "B"]。不填则根据所有列判断',
                items: {
                    type: 'string',
                },
            },
            has_header: {
                type: 'boolean',
                description: '第一行是否为表头，默认true',
            },
            sheet: {
                type: 'string',
                description: '工作表名称，不填则使用当前活动工作表',
            },
        },
        required: ['range'],
    },
};
const removeDuplicatesHandler = async (args) => {
    const { range, columns, has_header, sheet } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('removeDuplicates', {
            range,
            columns: columns || [],
            hasHeader: has_header !== false, // 默认true
            sheet,
        }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success || !response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `删除重复行失败: ${response.error}` }],
                error: response.error,
            };
        }
        const result = response.data;
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `删除重复行完成！
原始行数: ${result.originalCount}
删除行数: ${result.removedCount}
剩余行数: ${result.remainingCount}`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `删除重复行出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.removeDuplicatesHandler = removeDuplicatesHandler;
/**
 * 对选定区域排序
 */
exports.sortRangeDefinition = {
    name: 'wps_excel_sort_range',
    description: '对Excel选定区域按指定列排序。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '要排序的范围，如 A1:D100' },
            column: { type: 'number', description: '排序依据的列号（从1开始）' },
            ascending: { type: 'boolean', description: '是否升序，默认true' },
        },
        required: ['range', 'column'],
    },
};
const sortRangeHandler = async (args) => {
    const { range, column, ascending } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('sortRange', { range, column, ascending: ascending !== false }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `排序失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `排序完成！范围: ${range}，按第${column}列${ascending !== false ? '升序' : '降序'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `排序出错: ${errMsg}` }], error: errMsg };
    }
};
exports.sortRangeHandler = sortRangeHandler;
/**
 * 查找并替换内容
 */
exports.findReplaceDefinition = {
    name: 'wps_excel_find_replace',
    description: '在Excel中查找并替换内容。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            find: { type: 'string', description: '要查找的文本' },
            replace: { type: 'string', description: '替换为的文本' },
            matchCase: { type: 'boolean', description: '是否区分大小写，默认false' },
        },
        required: ['find', 'replace'],
    },
};
const findReplaceHandler = async (args) => {
    const { find, replace, matchCase } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('findReplace', { find, replace, matchCase: matchCase || false }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `查找替换失败: ${response.error}` }], error: response.error };
        }
        const count = response.data?.count || 0;
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `查找替换完成！将"${find}"替换为"${replace}"，共替换${count}处` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `查找替换出错: ${errMsg}` }], error: errMsg };
    }
};
exports.findReplaceHandler = findReplaceHandler;
/**
 * 插入行
 */
exports.insertRowDefinition = {
    name: 'wps_excel_insert_row',
    description: '在Excel中插入行。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            row: { type: 'number', description: '在第几行前插入（从1开始）' },
            count: { type: 'number', description: '插入行数，默认1' },
        },
        required: ['row'],
    },
};
const insertRowHandler = async (args) => {
    const { row, count } = args;
    const insertCount = count || 1;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertRows', { row, count: insertCount }, wps_1.WpsAppType.SPREADSHEET);
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
exports.insertRowHandler = insertRowHandler;
/**
 * 给单元格添加批注
 */
exports.addCommentDefinition = {
    name: 'wps_excel_add_comment',
    description: '给单元格添加批注。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            cell: { type: 'string', description: '单元格地址，如 A1、B2' },
            comment: { type: 'string', description: '批注内容' },
        },
        required: ['cell', 'comment'],
    },
};
const addCommentHandler = async (args) => {
    const { cell, comment } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addComment', { cell, comment }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `添加批注失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `批注添加成功！单元格: ${cell}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `添加批注出错: ${errMsg}` }], error: errMsg };
    }
};
exports.addCommentHandler = addCommentHandler;
/**
 * 保护/取消保护工作表
 */
exports.protectSheetDefinition = {
    name: 'wps_excel_protect_sheet',
    description: '保护或取消保护工作表。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            password: { type: 'string', description: '保护密码（可选）' },
            protect: { type: 'boolean', description: '是否保护，默认true' },
        },
        required: [],
    },
};
const protectSheetHandler = async (args) => {
    const { password, protect } = args;
    const doProtect = protect !== false;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('protectSheet', { password, protect: doProtect }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `${doProtect ? '保护' : '取消保护'}工作表失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `工作表${doProtect ? '保护' : '取消保护'}成功！` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `${doProtect ? '保护' : '取消保护'}工作表出错: ${errMsg}` }], error: errMsg };
    }
};
exports.protectSheetHandler = protectSheetHandler;
/**
 * 设置条件格式
 */
exports.setConditionalFormatDefinition = {
    name: 'wps_excel_set_conditional_format',
    description: '设置条件格式。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '要设置条件格式的范围，如 A1:D100' },
            condition: { type: 'string', description: '条件表达式，如 ">100"、"=0"、"between(1,10)"' },
            format: { type: 'string', description: '格式描述，如 "red_fill"、"bold"、"green_font"' },
        },
        required: ['range', 'condition', 'format'],
    },
};
const setConditionalFormatHandler = async (args) => {
    const { range, condition, format } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addConditionalFormat', { range, condition, format }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置条件格式失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `条件格式设置成功！范围: ${range}，条件: ${condition}，格式: ${format}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置条件格式出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setConditionalFormatHandler = setConditionalFormatHandler;
/**
 * 保护/取消保护工作簿
 */
exports.protectWorkbookDefinition = {
    name: 'wps_excel_protect_workbook',
    description: '保护或取消保护工作簿，防止结构被修改（如添加/删除工作表）。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            password: { type: 'string', description: '保护密码' },
            protect: { type: 'boolean', description: '是否保护，true为保护，false为取消保护' },
        },
        required: ['password', 'protect'],
    },
};
const protectWorkbookHandler = async (args) => {
    const { password, protect } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('protectWorkbook', { password, protect }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `${protect ? '保护' : '取消保护'}工作簿失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `工作簿${protect ? '保护' : '取消保护'}成功！` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `${protect ? '保护' : '取消保护'}工作簿出错: ${errMsg}` }], error: errMsg };
    }
};
exports.protectWorkbookHandler = protectWorkbookHandler;
/**
 * 设置工作表缩放比例
 */
exports.setZoomDefinition = {
    name: 'wps_excel_set_zoom',
    description: '设置当前工作表的缩放比例（10-400%）。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            percent: { type: 'number', description: '缩放百分比，范围10-400' },
        },
        required: ['percent'],
    },
};
const setZoomHandler = async (args) => {
    const { percent } = args;
    if (percent < 10 || percent > 400) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '缩放比例必须在10-400之间' }], error: '缩放比例超出范围' };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setZoom', // NOTE: macOS未实现，仅Windows支持
        { percent }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置缩放失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `缩放比例已设置为${percent}%` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置缩放出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setZoomHandler = setZoomHandler;
/**
 * 导出所有数据处理相关的Tools
 */
exports.dataTools = [
    { definition: exports.readRangeDefinition, handler: exports.readRangeHandler },
    { definition: exports.writeRangeDefinition, handler: exports.writeRangeHandler },
    { definition: exports.cleanDataDefinition, handler: exports.cleanDataHandler },
    { definition: exports.removeDuplicatesDefinition, handler: exports.removeDuplicatesHandler },
    { definition: exports.sortRangeDefinition, handler: exports.sortRangeHandler },
    { definition: exports.findReplaceDefinition, handler: exports.findReplaceHandler },
    { definition: exports.insertRowDefinition, handler: exports.insertRowHandler },
    { definition: exports.addCommentDefinition, handler: exports.addCommentHandler },
    { definition: exports.protectSheetDefinition, handler: exports.protectSheetHandler },
    { definition: exports.setConditionalFormatDefinition, handler: exports.setConditionalFormatHandler },
    { definition: exports.protectWorkbookDefinition, handler: exports.protectWorkbookHandler },
    { definition: exports.setZoomDefinition, handler: exports.setZoomHandler },
];
exports.default = exports.dataTools;
//# sourceMappingURL=data.js.map
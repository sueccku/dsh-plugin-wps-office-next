"use strict";
/**
 * Input: 公式类工具参数
 * Output: 公式计算与诊断结果
 * Pos: Excel 公式工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel公式相关Tools - 公式管理模块
 * 解决用户"公式不会写"痛点的核心工具集
 *
 * 包含：
 * - wps_excel_set_formula: 设置公式到指定单元格
 * - wps_excel_generate_formula: 根据自然语言生成公式（核心功能）
 * - wps_excel_diagnose_formula: 诊断公式错误，分析原因并提供修复建议
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.formulaTools = exports.zoomHandler = exports.zoomDefinition = exports.setPrintAreaHandler = exports.setPrintAreaDefinition = exports.evaluateFormulaHandler = exports.evaluateFormulaDefinition = exports.diagnoseFormulaHandler = exports.diagnoseFormulaDefinition = exports.generateFormulaHandler = exports.generateFormulaDefinition = exports.setFormulaHandler = exports.setFormulaDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * Render the computed value the bridge read back after a formula write.
 * The bridge reads it only for a single-cell target, so anything else must not be
 * reported as "计算结果: null" — that reads as a failed calculation.
 */
function describeFormulaValue(data) {
    const cellCount = typeof data?.cellCount === 'number' ? data.cellCount : 1;
    const parts = [];
    if (data && Object.prototype.hasOwnProperty.call(data, 'value')) {
        parts.push(`\n计算结果${cellCount > 1 ? '（区域首格）' : ''}: ${JSON.stringify(data.value)}`);
    }
    else {
        parts.push('\n计算结果: 未回读（可用 wps_excel_read_range 复验）');
    }
    if (cellCount > 1) {
        // Excel broadcasts one formula string to the whole target without shifting relative
        // references, and a caller who expects per-row results reads the trap as a broken SUMIF.
        parts.push(`\n注意: 已向 ${cellCount} 个单元格写入同一个公式（与 Excel 一致，不做相对引用调整）；` +
            '需要逐行/逐列递增的公式时，请逐格调用 wps_excel_set_formula，或直接用 wps_excel_write_range 写入数值。');
    }
    return parts.join('');
}
/**
 * 设置公式到指定单元格
 * 公式功能的执行端，负责将生成的公式写入单元格
 */
exports.setFormulaDefinition = {
    name: 'wps_excel_set_formula',
    description: '在指定单元格设置Excel公式。公式必须以=开头，支持所有Excel内置函数。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: {
                type: 'string',
                description: '目标单元格地址，如 A1、B2:B10',
            },
            formula: {
                type: 'string',
                description: 'Excel公式，必须以=开头，如 =SUM(A1:A10)、=VLOOKUP(A1,B:C,2,0)',
            },
            sheet: {
                type: 'string',
                description: '工作表名称，不填则使用当前活动工作表',
            },
        },
        required: ['range', 'formula'],
    },
};
const setFormulaHandler = async (args) => {
    const { range, formula, sheet } = args;
    // 公式必须以=开头
    if (!formula.startsWith('=')) {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '公式格式错误：公式必须以=开头，如 =SUM(A1:A10)' }],
            error: '公式格式错误：必须以=开头',
        };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setFormula', { range, formula, sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        // Only claim a computed value when the bridge read one back; the old line
                        // printed "计算结果: null" for every multi-cell range, which read as a failure.
                        text: `公式设置成功！\n单元格: ${range}\n公式: ${formula}${describeFormulaValue(response.data)}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `公式设置失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置公式出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setFormulaHandler = setFormulaHandler;
/**
 * 根据自然语言生成Excel公式
 * 核心功能：获取工作表上下文，辅助AI理解表结构后生成公式
 */
exports.generateFormulaDefinition = {
    name: 'wps_excel_generate_formula',
    description: `根据自然语言描述生成Excel公式。这是解决用户"公式不会写"痛点的核心工具。

使用场景：
- 用户说"帮我写个公式查价格" -> 生成 VLOOKUP/XLOOKUP
- 用户说"如果大于100就显示达标" -> 生成 IF 公式
- 用户说"统计每个部门的人数" -> 生成 COUNTIF
- 用户说"求这列的平均值" -> 生成 AVERAGE

调用此工具会返回当前工作表的上下文信息，包括表头、选中区域等，便于生成准确的公式。`,
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            description: {
                type: 'string',
                description: '用户对公式需求的自然语言描述，如"查找产品名对应的价格"、"计算A列的总和"',
            },
            target_cell: {
                type: 'string',
                description: '目标单元格地址，如 B2。不填则返回上下文让用户确认',
            },
        },
        required: ['description'],
    },
};
const generateFormulaHandler = async (args) => {
    const { description, target_cell } = args;
    try {
        // 先获取工作表上下文，这样AI才能理解表结构
        const contextResponse = await wps_client_1.wpsClient.executeMethod('getContext', {}, wps_1.WpsAppType.SPREADSHEET);
        if (!contextResponse.success || !contextResponse.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [
                    {
                        type: 'text',
                        text: '获取工作表上下文失败，请确保WPS表格已打开并且有活动工作簿',
                    },
                ],
                error: '无法获取工作表上下文',
            };
        }
        const context = contextResponse.data;
        // 返回上下文信息，让AI根据这些信息生成公式
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `工作表上下文信息：
工作簿: ${context.workbookName}
当前工作表: ${context.currentSheet}
所有工作表: ${context.allSheets.join(', ')}
当前选中: ${context.selectedCell}
数据范围: ${context.usedRangeAddress}
表头信息:
${context.headers.map((h) => `  ${h.column}列: ${h.value}`).join('\n')}

用户需求: ${description}
目标单元格: ${target_cell || '待确认'}

请根据上下文和用户描述生成公式，然后调用 wps_excel_set_formula 写入。`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取上下文出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.generateFormulaHandler = generateFormulaHandler;
/**
 * 诊断公式错误
 * 分析#REF!、#N/A、#VALUE!等错误类型，给出原因和修复建议
 */
exports.diagnoseFormulaDefinition = {
    name: 'wps_excel_diagnose_formula',
    description: `诊断公式错误，分析原因并提供修复建议。

使用场景：
- 用户说"这个公式报错了"
- 用户说"#REF! 是什么意思"
- 用户说"帮我看看公式哪里有问题"

支持诊断的错误类型：
- #REF! - 引用了不存在的单元格
- #N/A - 查找函数未找到匹配值
- #VALUE! - 参数类型错误
- #NAME? - 函数名称错误
- #DIV/0! - 除数为零
- #NUM! - 数值问题
- #NULL! - 交集为空`,
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            cell: {
                type: 'string',
                description: '包含错误公式的单元格地址，如 A1、B2',
            },
        },
        required: ['cell'],
    },
};
const diagnoseFormulaHandler = async (args) => {
    const { cell } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('diagnoseFormula', { cell }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success || !response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `诊断失败: ${response.error}` }],
                error: response.error,
            };
        }
        const diagnosis = response.data;
        // 如果没有错误
        if (!diagnosis.errorType) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `单元格 ${cell} 的公式没有错误！
公式: ${diagnosis.formula}
计算结果: ${JSON.stringify(diagnosis.currentValue)}`,
                    },
                ],
            };
        }
        // 有错误，给出详细诊断
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `公式诊断结果：

单元格: ${diagnosis.cell}
公式: ${diagnosis.formula}
错误类型: ${diagnosis.errorType}

错误原因: ${diagnosis.diagnosis}

修复建议: ${diagnosis.suggestion}

引用的单元格: ${diagnosis.precedents.length > 0 ? diagnosis.precedents.join(', ') : '无引用'}`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `诊断公式出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.diagnoseFormulaHandler = diagnoseFormulaHandler;
/**
 * 导出所有公式相关的Tools
 */
exports.evaluateFormulaDefinition = {
    name: 'wps_excel_evaluate_formula',
    description: `计算并返回公式结果（由 Excel 求值，不写入任何单元格）。

需要在某个单元格里求值并保留公式时，用 wps_excel_set_formula 写入再从该格读值。`,
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        // "cell" used to be declared here but the action evaluates with Application.Evaluate and never
        // read it, so a caller passing it now gets an explicit error instead of a silent no-op.
        properties: {
            formula: { type: 'string', description: '要计算的公式，如 =SUM(A1:A10)' },
        },
        required: ['formula'],
    },
};
const evaluateFormulaHandler = async (args) => {
    const response = await wps_client_1.wpsClient.executeMethod('evaluateFormula', args, wps_1.WpsAppType.SPREADSHEET // NOTE: macOS未实现，仅Windows支持
    );
    return { id: (0, uuid_1.v4)(), success: response.success, content: [{ type: "text", text: JSON.stringify(response.data ?? { error: response.error }) }] };
};
exports.evaluateFormulaHandler = evaluateFormulaHandler;
exports.setPrintAreaDefinition = {
    name: 'wps_excel_set_print_area',
    description: '设置打印区域',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: { range: { type: 'string', description: '打印区域，如 A1:D20' } },
        required: ['range'],
    },
};
const setPrintAreaHandler = async (args) => {
    const response = await wps_client_1.wpsClient.executeMethod('setPrintArea', args, wps_1.WpsAppType.SPREADSHEET);
    return { id: (0, uuid_1.v4)(), success: response.success, content: [{ type: "text", text: response.success ? "打印区域已设置" : "设置失败" }] };
};
exports.setPrintAreaHandler = setPrintAreaHandler;
exports.zoomDefinition = {
    name: 'wps_excel_zoom',
    description: '设置工作表缩放比例',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: { percent: { type: 'number', description: '缩放百分比，如 100' } },
        required: ['percent'],
    },
};
const zoomHandler = async (args) => {
    const response = await wps_client_1.wpsClient.executeMethod('setZoom', args, wps_1.WpsAppType.SPREADSHEET // NOTE: macOS未实现，仅Windows支持
    );
    return { id: (0, uuid_1.v4)(), success: response.success, content: [{ type: "text", text: response.success ? "缩放已设置" : "设置失败" }] };
};
exports.zoomHandler = zoomHandler;
exports.formulaTools = [
    { definition: exports.setFormulaDefinition, handler: exports.setFormulaHandler },
    { definition: exports.generateFormulaDefinition, handler: exports.generateFormulaHandler },
    { definition: exports.diagnoseFormulaDefinition, handler: exports.diagnoseFormulaHandler },
    { definition: exports.evaluateFormulaDefinition, handler: exports.evaluateFormulaHandler },
    { definition: exports.setPrintAreaDefinition, handler: exports.setPrintAreaHandler },
    { definition: exports.zoomDefinition, handler: exports.zoomHandler },
];
exports.default = exports.formulaTools;
//# sourceMappingURL=formula.js.map
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.excelAnalysisTools = exports.groupPivotDatesHandler = exports.groupPivotDatesDefinition = exports.setSeriesAxisHandler = exports.setSeriesAxisDefinition = exports.advancedFilterHandler = exports.advancedFilterDefinition = exports.addPivotCalculatedFieldHandler = exports.addPivotCalculatedFieldDefinition = exports.importCsvHandler = exports.importCsvDefinition = exports.addTrendlineHandler = exports.addTrendlineDefinition = exports.addIconSetHandler = exports.addIconSetDefinition = void 0;
/**
 * Input: Excel 高级分析/导入参数
 * Output: P1-5（2026-10-21）工具结果
 * Pos: 图标集/趋势线/CSV 导入/透视计算字段/高级筛选/组合图次轴；默认隐藏。
 */
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
const fail = (label, error) => ({
    id: (0, uuid_1.v4)(), success: false,
    content: [{ type: 'text', text: `${label}: ${error}` }],
    error,
});
exports.addIconSetDefinition = {
    name: 'wps_excel_add_icon_set',
    description: '给区域加图标集条件格式（三色交通灯/箭头/旗标等）。iconSet 用 WPS 的 IconSets 序号（1 起），reverse 反转顺序，showValue=false 只显示图标。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '区域，如 B2:B100' },
            iconSet: { type: 'number', description: '图标集序号（1 起，WPS IconSets 顺序）' },
            reverse: { type: 'boolean', description: '是否反转图标顺序' },
            showValue: { type: 'boolean', description: '是否同时显示数值，默认 true' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['range'],
    },
};
const addIconSetHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addIconSet', args, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return fail('添加图标集失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `图标集已添加；区域条件格式共 ${response.data?.count ?? 0} 条` }] };
    }
    catch (error) {
        return fail('添加图标集出错', error instanceof Error ? error.message : String(error));
    }
};
exports.addIconSetHandler = addIconSetHandler;
exports.addTrendlineDefinition = {
    name: 'wps_excel_add_trendline',
    description: '给图表某系列加趋势线（线性/指数/对数/多项式/幂/移动平均），可设置前推/后推、显示公式与 R²。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            chartIndex: { type: 'number', description: '图表序号（从1开始），默认1' },
            seriesIndex: { type: 'number', description: '系列序号（从1开始），默认1' },
            type: { type: 'number', description: '趋势线类型（WPS TrendlineType 枚举值，如 1=线性；不填用默认）' },
            forward: { type: 'number', description: '前推周期' },
            backward: { type: 'number', description: '后推周期' },
            displayEquation: { type: 'boolean', description: '显示公式' },
            displayRSquared: { type: 'boolean', description: '显示 R²' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: [],
    },
};
const addTrendlineHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addTrendline', args, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return fail('添加趋势线失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `趋势线已添加: 图表 ${response.data?.chartIndex} / 系列 ${response.data?.seriesIndex}` }] };
    }
    catch (error) {
        return fail('添加趋势线出错', error instanceof Error ? error.message : String(error));
    }
};
exports.addTrendlineHandler = addTrendlineHandler;
exports.importCsvDefinition = {
    name: 'wps_excel_import_csv',
    description: '按指定分隔符/编码导入 CSV（OpenText），保留前导零/日期识别比直接打开更可控；导入后工作簿未保存。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            path: { type: 'string', description: 'CSV 路径' },
            delimiter: { type: 'string', enum: ['comma', 'tab'], description: '分隔符，默认 comma' },
            encoding: { type: 'string', enum: ['utf8', 'gbk'], description: '编码，默认 utf8' },
        },
        required: ['path'],
    },
};
const importCsvHandler = async (args) => {
    const { path, delimiter = 'comma', encoding = 'utf8' } = args;
    if (!path)
        return fail('导入 CSV 失败', 'path 不能为空');
    try {
        const response = await wps_client_1.wpsClient.executeMethod('importCsv', { path, delimiter, encoding }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return fail('导入 CSV 失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `CSV 已导入: ${response.data?.name}\n路径: ${response.data?.path}` }] };
    }
    catch (error) {
        return fail('导入 CSV 出错', error instanceof Error ? error.message : String(error));
    }
};
exports.importCsvHandler = importCsvHandler;
exports.addPivotCalculatedFieldDefinition = {
    name: 'wps_excel_add_pivot_calculated_field',
    description: '给透视表加计算字段（CalculatedFields.Add），formula 用透视字段名，如 =销售额/数量。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            name: { type: 'string', description: '计算字段名' },
            formula: { type: 'string', description: '公式，如 =销售额/数量' },
            pivotTableName: { type: 'string', description: '透视表名；不填用活动表第一个' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['name', 'formula'],
    },
};
const addPivotCalculatedFieldHandler = async (args) => {
    const { name, formula } = args;
    if (!name || !formula)
        return fail('添加计算字段失败', 'name/formula 不能为空');
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addPivotCalculatedField', args, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return fail('添加透视计算字段失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `计算字段已添加: ${response.data?.name}（透视表 ${response.data?.pivotTable}）` }] };
    }
    catch (error) {
        return fail('添加透视计算字段出错', error instanceof Error ? error.message : String(error));
    }
};
exports.addPivotCalculatedFieldHandler = addPivotCalculatedFieldHandler;
exports.advancedFilterDefinition = {
    name: 'wps_excel_advanced_filter',
    description: '高级筛选（AdvancedFilter）：criteriaRange 写条件区；action=inPlace 原地隐藏不匹配行，copyTo 复制到目标区域（需给 copyTo）。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            range: { type: 'string', description: '数据区域（含表头）' },
            criteriaRange: { type: 'string', description: '条件区域（含条件表头）' },
            action: { type: 'string', enum: ['inPlace', 'copyTo'], description: '原地筛选或复制到别处' },
            copyTo: { type: 'string', description: 'copyTo 模式的目标起始单元格' },
            unique: { type: 'boolean', description: '是否只保留唯一记录' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['range', 'criteriaRange', 'action'],
    },
};
const advancedFilterHandler = async (args) => {
    const { range, criteriaRange, action, copyTo } = args;
    if (!range || !criteriaRange)
        return fail('高级筛选失败', 'range/criteriaRange 不能为空');
    if (action === 'copyTo' && !copyTo)
        return fail('高级筛选失败', 'copyTo 模式需要 copyTo');
    try {
        const response = await wps_client_1.wpsClient.executeMethod('advancedFilter', args, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return fail('高级筛选失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `高级筛选完成: ${response.data?.action} (${response.data?.range})` }] };
    }
    catch (error) {
        return fail('高级筛选出错', error instanceof Error ? error.message : String(error));
    }
};
exports.advancedFilterHandler = advancedFilterHandler;
exports.setSeriesAxisDefinition = {
    name: 'wps_excel_set_series_axis',
    description: '把图表某系列改成另一种图型并放到次坐标轴（组合图）：chartType 用 WPS 图表类型枚举，secondaryAxis=true 走次轴。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            chartIndex: { type: 'number', description: '图表序号（从1开始），默认1' },
            seriesIndex: { type: 'number', description: '系列序号（从1开始），默认1' },
            chartType: { type: 'number', description: 'WPS ChartType 枚举值（如 4=折线）' },
            secondaryAxis: { type: 'boolean', description: 'true 放次坐标轴' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: [],
    },
};
const setSeriesAxisHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSeriesAxis', args, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return fail('设置组合图系列失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `系列已更新: 图表 ${response.data?.chartIndex} / 系列 ${response.data?.seriesIndex}` }] };
    }
    catch (error) {
        return fail('设置组合图系列出错', error instanceof Error ? error.message : String(error));
    }
};
exports.setSeriesAxisHandler = setSeriesAxisHandler;
exports.groupPivotDatesDefinition = {
    name: 'wps_excel_group_pivot_dates',
    description: '对透视表的日期字段做自动分组（PivotField.AutoGroup）；WPS 12.1 只暴露 AutoGroup，按 WPS 默认粒度分组。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            pivotTableName: { type: 'string', description: '透视表名；不填用活动表第一个' },
            field: { type: 'string', description: '日期字段名（透视表字段名）' },
            sheet: { type: 'string', description: '工作表名称' },
        },
        required: ['field'],
    },
};
const groupPivotDatesHandler = async (args) => {
    const { field } = args;
    if (!field)
        return fail('日期分组失败', 'field 不能为空');
    try {
        const response = await wps_client_1.wpsClient.executeMethod('groupPivotDates', args, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return fail('日期分组失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `日期字段已分组: ${response.data?.field}（${response.data?.pivotTable}）` }] };
    }
    catch (error) {
        return fail('日期分组出错', error instanceof Error ? error.message : String(error));
    }
};
exports.groupPivotDatesHandler = groupPivotDatesHandler;
exports.excelAnalysisTools = [
    { definition: exports.addIconSetDefinition, handler: exports.addIconSetHandler },
    { definition: exports.addTrendlineDefinition, handler: exports.addTrendlineHandler },
    { definition: exports.importCsvDefinition, handler: exports.importCsvHandler },
    { definition: exports.addPivotCalculatedFieldDefinition, handler: exports.addPivotCalculatedFieldHandler },
    { definition: exports.advancedFilterDefinition, handler: exports.advancedFilterHandler },
    { definition: exports.setSeriesAxisDefinition, handler: exports.setSeriesAxisHandler },
    { definition: exports.groupPivotDatesDefinition, handler: exports.groupPivotDatesHandler },
];
//# sourceMappingURL=analysis.js.map
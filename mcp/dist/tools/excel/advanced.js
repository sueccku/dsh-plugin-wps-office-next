"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.advancedTools = exports.setChartLabelsHandler = exports.setChartLabelsDefinition = exports.deleteChartHandler = exports.deleteChartDefinition = exports.clearSparklineHandler = exports.clearSparklineDefinition = exports.addSparklineHandler = exports.addSparklineDefinition = exports.goalSeekHandler = exports.goalSeekDefinition = exports.refreshAllDataHandler = exports.refreshAllDataDefinition = exports.clearPivotTableHandler = exports.clearPivotTableDefinition = exports.refreshPivotTablesHandler = exports.refreshPivotTablesDefinition = exports.getPivotTablesHandler = exports.getPivotTablesDefinition = void 0;
/**
 * Input: 高级项工具的调用参数
 * Output: 透视表 / 单变量求解 / 迷你图 / 图表标签的执行结果
 * Pos: Excel 高级项工具（P2-4）。全部先用裸 COM 量过支持面；切片器与场景管理器按实测主动推迟（FIXES 43）。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P2 状态。
 */
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
const sheetParam = { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' };
const chartParam = { type: 'string', description: '图表名（如 Chart 1）或该表上的序号；只有一张图时可省略' };
function advancedFail(prefix, error) {
    return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: prefix + ': ' + error }], error: error };
}
/** 透视表列表 */
exports.getPivotTablesDefinition = {
    name: 'wps_excel_get_pivot_tables',
    description: '列出工作簿（或指定工作表）上的透视表：名字、所在区域，以及（能读到时）行字段与数据字段。使用场景：先看清有哪几张透视表、它们叫什么，再刷新或清除。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: { type: 'object', properties: { sheet: { type: 'string', description: '只看这张工作表；不填则列出整个工作簿' } } },
};
const getPivotTablesHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getPivotTables', { sheet: args.sheet }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return advancedFail('读取透视表失败', response.error);
        const tables = response.data?.tables || [];
        if (!tables.length)
            return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '这个范围内没有透视表。' }] };
        const lines = tables.map((t) => {
            const parts = ['  ' + String(t.index) + '. 「' + (t.name || '?') + '」（工作表 ' + (t.sheet || '?') + '，' + (t.range || '?') + '）'];
            if (t.rowField)
                parts.push('行字段 ' + t.rowField);
            if (t.dataField)
                parts.push('数据字段 ' + t.dataField);
            return parts.join('，');
        });
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '共 ' + tables.length + ' 张透视表:\n' + lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return advancedFail('读取透视表出错', errMsg);
    }
};
exports.getPivotTablesHandler = getPivotTablesHandler;
/** 刷新透视表 */
exports.refreshPivotTablesDefinition = {
    name: 'wps_excel_refresh_pivot_tables',
    description: '刷新透视表：给 pivotTable 只刷新那一张，不给就刷新目标工作表上的全部。使用场景：源数据改了，透视表还是旧数字。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            pivotTable: { type: 'string', description: '透视表名；不填则刷新目标工作表上的全部透视表' },
            sheet: sheetParam,
        },
    },
};
const refreshPivotTablesHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('refreshPivotTables', { sheet: args.sheet, pivotTable: args.pivotTable }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return advancedFail('刷新透视表失败', response.error);
        const names = response.data?.names || [];
        const n = response.data?.refreshed ?? names.length;
        const detail = names.length ? '：' + names.join(', ') : '';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已刷新 ' + n + ' 张透视表' + detail }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return advancedFail('刷新透视表出错', errMsg);
    }
};
exports.refreshPivotTablesHandler = refreshPivotTablesHandler;
/** 清除透视表报表 */
exports.clearPivotTableDefinition = {
    name: 'wps_excel_clear_pivot_table',
    description: '清除透视表在表上的报表区域（数据源不动）。注意：WPS 清掉报表后透视表对象会留到保存/重开，期间它仍出现在透视表列表里——工具会如实报告剩余数量，不谎称已删除。使用场景：把临时透视表从工作表上拿掉。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            pivotTable: { type: 'string', description: '要清除的透视表名' },
            sheet: sheetParam,
        },
        required: ['pivotTable'],
    },
};
const clearPivotTableHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('clearPivotTable', { sheet: args.sheet, pivotTable: args.pivotTable }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return advancedFail('清除透视表失败', response.error);
        const d = response.data || {};
        const head = '已清除透视表「' + (d.name || String(args.pivotTable)) + '」的报表区域（原区域 ' + (d.rangeBefore || '?') + '）';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: head + '\n' + (d.message || '') + '（当前工作表上还有 ' + String(d.remaining ?? 0) + ' 个透视表对象）' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return advancedFail('清除透视表出错', errMsg);
    }
};
exports.clearPivotTableHandler = clearPivotTableHandler;
/** RefreshAll */
exports.refreshAllDataDefinition = {
    name: 'wps_excel_refresh_all_data',
    description: '刷新整个工作簿的外部数据连接与透视表（相当于 Excel 的「全部刷新」）。使用场景：多个数据源都要更新一次。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: { type: 'object', properties: {} },
};
const refreshAllDataHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('refreshAllData', {}, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return advancedFail('全部刷新失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: response.data?.message || '已全部刷新' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return advancedFail('全部刷新出错', errMsg);
    }
};
exports.refreshAllDataHandler = refreshAllDataHandler;
/** 单变量求解 */
exports.goalSeekDefinition = {
    name: 'wps_excel_goal_seek',
    description: '单变量求解：反复调整 changingCell，直到 cell 的公式结果等于 goal。cell 必须是带公式的单元格。使用场景："要利润到 100 万，销量得多少"。结果是近似解，工具会回读调整后的取值。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            cell: { type: 'string', description: '带公式的目标单元格，如 B4' },
            goal: { type: 'number', description: '希望目标单元格达到的值' },
            changingCell: { type: 'string', description: '被反复调整的单元格（不能有公式），如 B2' },
            sheet: sheetParam,
        },
        required: ['cell', 'goal', 'changingCell'],
    },
};
const goalSeekHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('goalSeek', { sheet: args.sheet, cell: args.cell, goal: args.goal, changingCell: args.changingCell }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return advancedFail('单变量求解失败', response.error);
        const d = response.data || {};
        const lines = ['单变量求解: ' + (d.solved ? '求出解' : '未收敛到解')];
        lines.push('  ' + String(d.changingCell) + ' = ' + String(d.changingValue));
        lines.push('  ' + String(d.cell) + ' 现在 = ' + String(d.resultValue) + '（目标 ' + String(d.goal) + '）');
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return advancedFail('单变量求解出错', errMsg);
    }
};
exports.goalSeekHandler = goalSeekHandler;
/** 迷你图 */
exports.addSparklineDefinition = {
    name: 'wps_excel_add_sparkline',
    description: '在单元格区域里加迷你图（单元格内的微型图表）：dataRange 是数据，location 是放图的位置，两者形状要一致（如 B2:B5 → C2:C5）。使用场景：在表格旁边一行一个小趋势图，不占地方。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            dataRange: { type: 'string', description: '数据区域，如 B2:B5（每个单元格一条迷你图时按列给）' },
            location: { type: 'string', description: '放置位置，如 C2:C5；形状要与 dataRange 一致' },
            sparklineType: { type: 'string', enum: ['line', 'column', 'winloss'], description: '迷你图类型，默认 line' },
            markers: { type: 'boolean', description: '是否标出数据点（折线图）' },
            sheet: sheetParam,
        },
        required: ['dataRange', 'location'],
    },
};
const addSparklineHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addSparkline', { sheet: args.sheet, dataRange: args.dataRange, location: args.location, sparklineType: args.sparklineType, markers: args.markers }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return advancedFail('添加迷你图失败', response.error);
        const d = response.data || {};
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已在 ' + (d.location || String(args.location)) + ' 添加 ' + (d.sparklineType || 'line') + ' 迷你图（数据 ' + (d.dataRange || String(args.dataRange)) + '，共 ' + String(d.groups ?? 0) + ' 组）' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return advancedFail('添加迷你图出错', errMsg);
    }
};
exports.addSparklineHandler = addSparklineHandler;
exports.clearSparklineDefinition = {
    name: 'wps_excel_clear_sparkline',
    description: '清除指定区域上的迷你图（数据不动）。使用场景：不想要这些微型图了。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            location: { type: 'string', description: '迷你图所在区域，如 C2:C5' },
            sheet: sheetParam,
        },
        required: ['location'],
    },
};
const clearSparklineHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('clearSparkline', { sheet: args.sheet, location: args.location }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return advancedFail('清除迷你图失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: (response.data?.location || String(args.location)) + ' 上的迷你图已清除（剩余 ' + String(response.data?.groups ?? 0) + ' 组）' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return advancedFail('清除迷你图出错', errMsg);
    }
};
exports.clearSparklineHandler = clearSparklineHandler;
/** 图表：删除与标题 */
exports.deleteChartDefinition = {
    name: 'wps_excel_delete_chart',
    description: '删除工作表上的图表（不删它引用的数据）。给 chart 名字或序号；表上只有一张图时可以省略。使用场景：清掉临时图表。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: { chart: chartParam, sheet: sheetParam },
    },
};
const deleteChartHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteChart', { sheet: args.sheet, chart: args.chart }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return advancedFail('删除图表失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已删除图表 ' + (response.data?.deleted || '') + '（该表还剩 ' + String(response.data?.remaining ?? 0) + ' 张）' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return advancedFail('删除图表出错', errMsg);
    }
};
exports.deleteChartHandler = deleteChartHandler;
exports.setChartLabelsDefinition = {
    name: 'wps_excel_set_chart_labels',
    description: '给图表加标题与坐标轴标题（分类轴 = 横轴，数值轴 = 纵轴）。使用场景：裸图没人看得懂，补上「月度销售」「月份」「金额」。',
    category: tools_1.ToolCategory.SPREADSHEET,
    inputSchema: {
        type: 'object',
        properties: {
            title: { type: 'string', description: '图表标题' },
            categoryAxisTitle: { type: 'string', description: '分类轴（横轴）标题' },
            valueAxisTitle: { type: 'string', description: '数值轴（纵轴）标题' },
            chart: chartParam,
            sheet: sheetParam,
        },
    },
};
const setChartLabelsHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setChartLabels', { sheet: args.sheet, chart: args.chart, title: args.title, categoryAxisTitle: args.categoryAxisTitle, valueAxisTitle: args.valueAxisTitle }, wps_1.WpsAppType.SPREADSHEET);
        if (!response.success)
            return advancedFail('设置图表标题失败', response.error);
        const d = response.data || {};
        const applied = d.applied || [];
        const lines = ['图表 ' + (d.chart || '') + (applied.length ? ' 已更新（' + applied.join(', ') + '）' : ' 没有变化')];
        if (d.title)
            lines.push('  标题: ' + d.title);
        if (d.categoryAxisTitle)
            lines.push('  横轴: ' + d.categoryAxisTitle);
        if (d.valueAxisTitle)
            lines.push('  纵轴: ' + d.valueAxisTitle);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return advancedFail('设置图表标题出错', errMsg);
    }
};
exports.setChartLabelsHandler = setChartLabelsHandler;
exports.advancedTools = [
    { definition: exports.getPivotTablesDefinition, handler: exports.getPivotTablesHandler },
    { definition: exports.refreshPivotTablesDefinition, handler: exports.refreshPivotTablesHandler },
    { definition: exports.clearPivotTableDefinition, handler: exports.clearPivotTableHandler },
    { definition: exports.refreshAllDataDefinition, handler: exports.refreshAllDataHandler },
    { definition: exports.goalSeekDefinition, handler: exports.goalSeekHandler },
    { definition: exports.addSparklineDefinition, handler: exports.addSparklineHandler },
    { definition: exports.clearSparklineDefinition, handler: exports.clearSparklineHandler },
    { definition: exports.deleteChartDefinition, handler: exports.deleteChartHandler },
    { definition: exports.setChartLabelsDefinition, handler: exports.setChartLabelsHandler },
];
exports.default = exports.advancedTools;
//# sourceMappingURL=advanced.js.map
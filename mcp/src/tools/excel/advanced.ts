/**
 * Input: 高级项工具的调用参数
 * Output: 透视表 / 单变量求解 / 迷你图 / 图表标签的执行结果
 * Pos: Excel 高级项工具（P2-4）。全部先用裸 COM 量过支持面；切片器与场景管理器按实测主动推迟（FIXES 43）。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P2 状态。
 */
import { v4 as uuidv4 } from 'uuid';
import { impactText, type RangeImpact } from '../impact';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

const sheetParam = { type: 'string' as const, description: '工作表名或序号；不填则用当前活动工作表' };
const chartParam = { type: 'string' as const, description: '图表名（如 Chart 1）或该表上的序号；只有一张图时可省略' };

function advancedFail(prefix: string, error?: string): ToolCallResult {
  return { id: uuidv4(), success: false, content: [{ type: 'text', text: prefix + ': ' + error }], error: error };
}

/** 透视表列表 */
export const getPivotTablesDefinition: ToolDefinition = {
  name: 'wps_excel_get_pivot_tables',
  description: '列出工作簿（或指定工作表）上的透视表：名字、所在区域，以及（能读到时）行字段与数据字段。使用场景：先看清有哪几张透视表、它们叫什么，再刷新或清除。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: { type: 'object', properties: { sheet: { type: 'string', description: '只看这张工作表；不填则列出整个工作簿' } } },
};

export const getPivotTablesHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ tables?: Array<{ sheet?: string; index?: number; name?: string; range?: string; rowField?: string; dataField?: string }>; count?: number }>(
      'getPivotTables',
      { sheet: args.sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return advancedFail('读取透视表失败', response.error);
    const tables = response.data?.tables || [];
    if (!tables.length) return { id: uuidv4(), success: true, content: [{ type: 'text', text: '这个范围内没有透视表。' }] };
    const lines = tables.map((t) => {
      const parts = ['  ' + String(t.index) + '. 「' + (t.name || '?') + '」（工作表 ' + (t.sheet || '?') + '，' + (t.range || '?') + '）'];
      if (t.rowField) parts.push('行字段 ' + t.rowField);
      if (t.dataField) parts.push('数据字段 ' + t.dataField);
      return parts.join('，');
    });
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '共 ' + tables.length + ' 张透视表:\n' + lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return advancedFail('读取透视表出错', errMsg);
  }
};

/** 刷新透视表 */
export const refreshPivotTablesDefinition: ToolDefinition = {
  name: 'wps_excel_refresh_pivot_tables',
  description: '刷新透视表：给 pivotTable 只刷新那一张，不给就刷新目标工作表上的全部。使用场景：源数据改了，透视表还是旧数字。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      pivotTable: { type: 'string', description: '透视表名；不填则刷新目标工作表上的全部透视表' },
      sheet: sheetParam,
    },
  },
};

export const refreshPivotTablesHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ refreshed?: number; names?: string[] }>(
      'refreshPivotTables',
      { sheet: args.sheet, pivotTable: args.pivotTable },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return advancedFail('刷新透视表失败', response.error);
    const names = response.data?.names || [];
    const n = response.data?.refreshed ?? names.length;
    const detail = names.length ? '：' + names.join(', ') : '';
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已刷新 ' + n + ' 张透视表' + detail }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return advancedFail('刷新透视表出错', errMsg);
  }
};

/** 清除透视表报表 */
export const clearPivotTableDefinition: ToolDefinition = {
  name: 'wps_excel_clear_pivot_table',
  description: '清除透视表在表上的报表区域（数据源不动）。注意：WPS 清掉报表后透视表对象会留到保存/重开，期间它仍出现在透视表列表里——工具会如实报告剩余数量，不谎称已删除。使用场景：把临时透视表从工作表上拿掉。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      pivotTable: { type: 'string', description: '要清除的透视表名' },
      sheet: sheetParam,
    },
    required: ['pivotTable'],
  },
};

export const clearPivotTableHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ sheet?: string; name?: string; rangeBefore?: string; remaining?: number; message?: string }>(
      'clearPivotTable',
      { sheet: args.sheet, pivotTable: args.pivotTable },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return advancedFail('清除透视表失败', response.error);
    const d = response.data || {};
    const head = '已清除透视表「' + (d.name || String(args.pivotTable)) + '」的报表区域（原区域 ' + (d.rangeBefore || '?') + '）';
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: head + '\n' + (d.message || '') + '（当前工作表上还有 ' + String(d.remaining ?? 0) + ' 个透视表对象）' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return advancedFail('清除透视表出错', errMsg);
  }
};

/** RefreshAll */
export const refreshAllDataDefinition: ToolDefinition = {
  name: 'wps_excel_refresh_all_data',
  description: '刷新整个工作簿的外部数据连接与透视表（相当于 Excel 的「全部刷新」）。使用场景：多个数据源都要更新一次。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: { type: 'object', properties: {} },
};

export const refreshAllDataHandler: ToolHandler = async (_args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ workbook?: string; message?: string }>(
      'refreshAllData',
      {},
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return advancedFail('全部刷新失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: response.data?.message || '已全部刷新' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return advancedFail('全部刷新出错', errMsg);
  }
};

/** 单变量求解 */
export const goalSeekDefinition: ToolDefinition = {
  name: 'wps_excel_goal_seek',
  description: '单变量求解：反复调整 changingCell，直到 cell 的公式结果等于 goal。cell 必须是带公式的单元格。使用场景："要利润到 100 万，销量得多少"。结果是近似解，工具会回读调整后的取值。',
  category: ToolCategory.SPREADSHEET,
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

export const goalSeekHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ cell?: string; goal?: number; changingCell?: string; solved?: boolean; changingValue?: unknown; resultValue?: unknown }>(
      'goalSeek',
      { sheet: args.sheet, cell: args.cell, goal: args.goal, changingCell: args.changingCell },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return advancedFail('单变量求解失败', response.error);
    const d = response.data || {};
    const lines = ['单变量求解: ' + (d.solved ? '求出解' : '未收敛到解')];
    lines.push('  ' + String(d.changingCell) + ' = ' + String(d.changingValue));
    lines.push('  ' + String(d.cell) + ' 现在 = ' + String(d.resultValue) + '（目标 ' + String(d.goal) + '）');
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return advancedFail('单变量求解出错', errMsg);
  }
};
/** 迷你图 */
export const addSparklineDefinition: ToolDefinition = {
  name: 'wps_excel_add_sparkline',
  description: '在单元格区域里加迷你图（单元格内的微型图表）：dataRange 是数据，location 是放图的位置，两者形状要一致（如 B2:B5 → C2:C5）。使用场景：在表格旁边一行一个小趋势图，不占地方。',
  category: ToolCategory.SPREADSHEET,
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

export const addSparklineHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ location?: string; dataRange?: string; sparklineType?: string; groups?: number }>(
      'addSparkline',
      { sheet: args.sheet, dataRange: args.dataRange, location: args.location, sparklineType: args.sparklineType, markers: args.markers },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return advancedFail('添加迷你图失败', response.error);
    const d = response.data || {};
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已在 ' + (d.location || String(args.location)) + ' 添加 ' + (d.sparklineType || 'line') + ' 迷你图（数据 ' + (d.dataRange || String(args.dataRange)) + '，共 ' + String(d.groups ?? 0) + ' 组）' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return advancedFail('添加迷你图出错', errMsg);
  }
};

export const clearSparklineDefinition: ToolDefinition = {
  name: 'wps_excel_clear_sparkline',
  description: '清除指定区域上的迷你图（数据不动）。使用场景：不想要这些微型图了。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      location: { type: 'string', description: '迷你图所在区域，如 C2:C5' },
      sheet: sheetParam,
    },
    required: ['location'],
  },
};

export const clearSparklineHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ location?: string; groups?: number; impact?: RangeImpact }>(
      'clearSparkline',
      { sheet: args.sheet, location: args.location },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return advancedFail('清除迷你图失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: (response.data?.location || String(args.location)) + ' 上的迷你图已清除（剩余 ' + String(response.data?.groups ?? 0) + ' 组）' + impactText(response.data?.impact) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return advancedFail('清除迷你图出错', errMsg);
  }
};

/** 图表：删除与标题 */
export const deleteChartDefinition: ToolDefinition = {
  name: 'wps_excel_delete_chart',
  description: '删除工作表上的图表（不删它引用的数据）。给 chart 名字或序号；表上只有一张图时可以省略。使用场景：清掉临时图表。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: { chart: chartParam, sheet: sheetParam },
  },
};

export const deleteChartHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ deleted?: string; remaining?: number; impact?: RangeImpact }>(
      'deleteChart',
      { sheet: args.sheet, chart: args.chart },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return advancedFail('删除图表失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已删除图表 ' + (response.data?.deleted || '') + '（该表还剩 ' + String(response.data?.remaining ?? 0) + ' 张）' + impactText(response.data?.impact) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return advancedFail('删除图表出错', errMsg);
  }
};

export const setChartLabelsDefinition: ToolDefinition = {
  name: 'wps_excel_set_chart_labels',
  description: '给图表加标题与坐标轴标题（分类轴 = 横轴，数值轴 = 纵轴）。使用场景：裸图没人看得懂，补上「月度销售」「月份」「金额」。',
  category: ToolCategory.SPREADSHEET,
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

export const setChartLabelsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ chart?: string; applied?: string[]; title?: string; categoryAxisTitle?: string; valueAxisTitle?: string }>(
      'setChartLabels',
      { sheet: args.sheet, chart: args.chart, title: args.title, categoryAxisTitle: args.categoryAxisTitle, valueAxisTitle: args.valueAxisTitle },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return advancedFail('设置图表标题失败', response.error);
    const d = response.data || {};
    const applied = d.applied || [];
    const lines = ['图表 ' + (d.chart || '') + (applied.length ? ' 已更新（' + applied.join(', ') + '）' : ' 没有变化')];
    if (d.title) lines.push('  标题: ' + d.title);
    if (d.categoryAxisTitle) lines.push('  横轴: ' + d.categoryAxisTitle);
    if (d.valueAxisTitle) lines.push('  纵轴: ' + d.valueAxisTitle);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return advancedFail('设置图表标题出错', errMsg);
  }
};

export const advancedTools: RegisteredTool[] = [
  { definition: getPivotTablesDefinition, handler: getPivotTablesHandler },
  { definition: refreshPivotTablesDefinition, handler: refreshPivotTablesHandler },
  { definition: clearPivotTableDefinition, handler: clearPivotTableHandler },
  { definition: refreshAllDataDefinition, handler: refreshAllDataHandler },
  { definition: goalSeekDefinition, handler: goalSeekHandler },
  { definition: addSparklineDefinition, handler: addSparklineHandler },
  { definition: clearSparklineDefinition, handler: clearSparklineHandler },
  { definition: deleteChartDefinition, handler: deleteChartHandler },
  { definition: setChartLabelsDefinition, handler: setChartLabelsHandler },
];

export default advancedTools;

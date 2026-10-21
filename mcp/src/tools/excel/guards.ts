/**
 * Input: Excel 允许编辑区域 / 外部链接 / 分页符参数
 * Output: P1（2026-10-19）新增的 Excel 半步工具结果
 * Pos: 工具默认隐藏；COM 细节在 wps-com.ps1 的 allowEditRange/getWorkbookLinks/insertPageBreak。
 */
import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

export const allowEditRangeDefinition: ToolDefinition = {
  name: 'wps_excel_allow_edit_range',
  description: '在工作表保护之下放行一块可编辑区域（AllowEditRanges）。先把单元格锁好/保护工作表，再给指定区域发编辑权限；可带密码。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '允许编辑的范围，如 B2:B100' },
      title: { type: 'string', description: '区域名（可选，默认自动生成）' },
      password: { type: 'string', description: '编辑密码（可选）' },
      sheet: { type: 'string', description: '工作表名称，不填用活动表' },
    },
    required: ['range'],
  },
};

export const allowEditRangeHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { range, title, password, sheet } = args as { range: string; title?: string; password?: string; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ title: string; range: string }>(
      'allowEditRange', { range, title, password, sheet }, WpsAppType.SPREADSHEET
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置允许编辑区域失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `允许编辑区域已设置: ${response.data?.title} (${response.data?.range})` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置允许编辑区域出错: ${errMsg}` }], error: errMsg };
  }
};

export const getLinksDefinition: ToolDefinition = {
  name: 'wps_excel_get_links',
  description: '列出当前工作簿的外部链接源（LinkSources），用于源被改名/移走、更新链接弹窗的排查。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: { type: 'object', properties: {}, required: [] },
};

export const getLinksHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ count: number; links: string[] }>(
      'getWorkbookLinks', {}, WpsAppType.SPREADSHEET
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取外部链接失败: ${response.error}` }], error: response.error };
    const d = response.data || { count: 0, links: [] };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `外部链接 (${d.count}):\n${d.links.join('\n') || '无'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取外部链接出错: ${errMsg}` }], error: errMsg };
  }
};

export const insertPageBreakDefinition: ToolDefinition = {
  name: 'wps_excel_insert_page_break',
  description: '在第 row 行上方插入手动分页符（row 从1开始）；只影响打印分页，不改数据。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      row: { type: 'number', description: '分页符插在第几行上方（从1开始）' },
      sheet: { type: 'string', description: '工作表名称，不填用活动表' },
    },
    required: ['row'],
  },
};

export const insertPageBreakHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { row, sheet } = args as { row: number; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ row: number }>(
      'insertPageBreak', { row, sheet }, WpsAppType.SPREADSHEET
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入分页符失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `已在第 ${response.data?.row} 行插入分页符` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入分页符出错: ${errMsg}` }], error: errMsg };
  }
};

export const disconnectLinkDefinition: ToolDefinition = {
  name: 'wps_excel_disconnect_link',
  description: '断开指定外部链接源（BreakLink），保留当前值、切断后续更新。source 用 wps_excel_get_links 取。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: { source: { type: 'string', description: '外部链接源，如 D:\\数据\\源.xlsx' } },
    required: ['source'],
  },
};

export const disconnectLinkHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { source } = args as { source: string };
  try {
    const response = await wpsClient.executeMethod<{ source: string }>(
      'disconnectLink', { source }, WpsAppType.SPREADSHEET
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `断开链接失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `已断开链接: ${response.data?.source}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `断开链接出错: ${errMsg}` }], error: errMsg };
  }
};

export const changeLinkDefinition: ToolDefinition = {
  name: 'wps_excel_change_link',
  description: '把外部链接从旧源改到新源（ChangeLink），不重建公式。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      oldSource: { type: 'string', description: '旧链接源' },
      newSource: { type: 'string', description: '新链接源' },
    },
    required: ['oldSource', 'newSource'],
  },
};

export const changeLinkHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { oldSource, newSource } = args as { oldSource: string; newSource: string };
  try {
    const response = await wpsClient.executeMethod<{ oldSource: string; newSource: string }>(
      'changeLink', { oldSource, newSource }, WpsAppType.SPREADSHEET
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `更改链接失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `链接已改到: ${response.data?.newSource}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `更改链接出错: ${errMsg}` }], error: errMsg };
  }
};

export const getChartsDefinition: ToolDefinition = {
  name: 'wps_excel_get_charts',
  description: '列出工作表里的图表：序号、名称、位置、尺寸、图表类型；用于改图前先确认 chartIndex。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: { sheet: { type: 'string', description: '工作表名称，不填用活动表' } },
    required: [],
  },
};

export const getChartsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { sheet } = args as { sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ count: number; charts: Array<Record<string, unknown>> }>(
      'listCharts', { sheet }, WpsAppType.SPREADSHEET
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取图表失败: ${response.error}` }], error: response.error };
    const d = response.data || { count: 0, charts: [] };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `图表 (${d.count}):\n${JSON.stringify(d.charts, null, 2)}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取图表出错: ${errMsg}` }], error: errMsg };
  }
};

export const setChartAxisDefinition: ToolDefinition = {
  name: 'wps_excel_set_chart_axis',
  description: '设置图表坐标轴：标题、数字格式、最小/最大值。axis=category（横轴）或 value（纵轴）；先用 get_charts 拿 chartIndex。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      chartIndex: { type: 'number', description: '图表序号（从1开始），默认1' },
      axis: { type: 'string', enum: ['category', 'value'], description: 'category=横轴，value=纵轴' },
      hasTitle: { type: 'boolean', description: '是否显示轴标题' },
      title: { type: 'string', description: '轴标题文字' },
      numberFormat: { type: 'string', description: '轴刻度数字格式，如 yyyy-mm、#,##0' },
      minimum: { type: 'number', description: '最小值（可选）' },
      maximum: { type: 'number', description: '最大值（可选）' },
      sheet: { type: 'string', description: '工作表名称，不填用活动表' },
    },
    required: ['axis'],
  },
};

export const setChartAxisHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { chartIndex, axis, hasTitle, title, numberFormat, minimum, maximum, sheet } = args as Record<string, unknown> & { axis: string };
  try {
    const response = await wpsClient.executeMethod<{ chartIndex: number; axis: string }>(
      'setChartAxis', { chartIndex, axis, hasTitle, title, numberFormat, minimum, maximum, sheet }, WpsAppType.SPREADSHEET
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置坐标轴失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `坐标轴已设置: 图表 ${response.data?.chartIndex} / ${response.data?.axis}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置坐标轴出错: ${errMsg}` }], error: errMsg };
  }
};

export const excelGuardTools: RegisteredTool[] = [
  { definition: allowEditRangeDefinition, handler: allowEditRangeHandler },
  { definition: getLinksDefinition, handler: getLinksHandler },
  { definition: insertPageBreakDefinition, handler: insertPageBreakHandler },
  { definition: disconnectLinkDefinition, handler: disconnectLinkHandler },
  { definition: changeLinkDefinition, handler: changeLinkHandler },
  { definition: getChartsDefinition, handler: getChartsHandler },
  { definition: setChartAxisDefinition, handler: setChartAxisHandler },
];

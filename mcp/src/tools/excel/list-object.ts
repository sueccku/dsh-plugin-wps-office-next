/**
 * Input: 表（ListObject）工具的调用参数
 * Output: 建表/读表/加行删行/样式与总计行/调整范围/转回区域 的执行结果
 * Pos: Excel 表工具（P2-2）。WPS 的 ListObject 支持完整，但桥里此前一条 action 都没有，
 *      所以这一族是新 COM 代码；工具层只做翻译与呈现，语义全在桥里。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P2 状态。
 */
import { v4 as uuidv4 } from 'uuid';
import { impactText, type RangeImpact } from './impact';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

interface ListObjectInfo {
  sheet?: string;
  name?: string;
  range?: string;
  columnCount?: number;
  rows?: number;
  hasTotals?: boolean;
  tableStyle?: string;
  columns?: Array<{ index: number; name: string }>;
  structuredRefs?: string[];
  message?: string;
}

/** 一张表的多行描述：结构 + 列名 + 结构化引用（可直接写进公式） */
function describeListObject(info: ListObjectInfo): string {
  const where = info.sheet ? '（工作表 ' + info.sheet + '，' + (info.range || '?') + '）' : '';
  const lines = ['表「' + (info.name || '?') + '」' + where + '：' + (info.columnCount || 0) + ' 列 / ' + (info.rows || 0) + ' 行，总计行 ' + (info.hasTotals ? '开' : '关') + (info.tableStyle ? '，样式 ' + info.tableStyle : '')];
  const names = (info.columns || []).map((c) => c.name).filter(Boolean);
  if (names.length) lines.push('列：' + names.join(', '));
  const refs = info.structuredRefs || [];
  if (refs.length) lines.push('结构化引用：' + refs.join(', '));
  if (info.message) lines.push(info.message);
  return lines.join('\n');
}

function listObjectFail(prefix: string, error?: string): ToolCallResult {
  return { id: uuidv4(), success: false, content: [{ type: 'text', text: prefix + ': ' + error }], error: error };
}

const tableParam = { type: 'string' as const, description: '表名（如 表1、Sales）或该表在工作表上的序号（从 1 开始）' };
const sheetParam = { type: 'string' as const, description: '工作表名或序号；不填则用当前活动工作表' };
/** 建表 */
export const createListObjectDefinition: ToolDefinition = {
  name: 'wps_excel_create_list_object',
  description: '把一块区域变成「表」（ListObject）：自动带表头与筛选按钮，并可用结构化引用（表名[列名]）写公式，之后能按表加行/删行/加总计行。使用场景：数据要反复增删、要按列筛选、要用结构化引用。默认认为首行是标题。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '建表的数据区域，含表头，如 A1:C10' },
      name: { type: 'string', description: '表名（不填由 WPS 自动命名，如 表1）；这个名字要能写进公式' },
      hasHeaders: { type: 'boolean', description: '首行是否为标题，默认 true；纯数据请传 false' },
      tableStyle: { type: 'string', description: '表格样式名，如 TableStyleMedium2；不填用 WPS 默认' },
      sheet: sheetParam,
    },
    required: ['range'],
  },
};

export const createListObjectHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<ListObjectInfo>(
      'createListObject',
      { sheet: args.sheet, range: args.range, name: args.name, hasHeaders: args.hasHeaders, tableStyle: args.tableStyle },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return listObjectFail('建表失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已建表\n' + describeListObject(response.data || {}) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return listObjectFail('建表出错', errMsg);
  }
};

/** 列出表 */
export const getListObjectsDefinition: ToolDefinition = {
  name: 'wps_excel_get_list_objects',
  description: '列出工作簿（或指定工作表）里的全部「表」及其结构：名字、范围、行列数、列名、表格样式、总计行开关，以及每列可直接写进公式的结构化引用。使用场景：先看清有哪几张表、列名叫什么，再决定怎么改。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '只看这张工作表；不填则列出整个工作簿' },
    },
  },
};

export const getListObjectsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ tables?: ListObjectInfo[]; count?: number }>(
      'getListObjects',
      { sheet: args.sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return listObjectFail('读取表失败', response.error);
    const tables = response.data?.tables || [];
    if (!tables.length) {
      return { id: uuidv4(), success: true, content: [{ type: 'text', text: '这个范围内没有表（ListObject）。要建表请用 wps_excel_create_list_object。' }] };
    }
    const blocks = tables.map(describeListObject);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '共 ' + tables.length + ' 张表:\n\n' + blocks.join('\n\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return listObjectFail('读取表出错', errMsg);
  }
};

/** 追加行 */
export const addListRowDefinition: ToolDefinition = {
  name: 'wps_excel_add_list_row',
  description: '给表末尾追加一行，可同时写入这一行的值（按列顺序）。列的格式与公式会自动带上。使用场景：往结构化表格里持续追加记录。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      table: tableParam,
      values: { type: 'array', description: '按列顺序写这一行的值，如 ["华东", "A", 10]；数字请写数字（写成字符串会当文本写进单元格）。不填只加一个空行' },
      sheet: sheetParam,
    },
    required: ['table'],
  },
};

export const addListRowHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<ListObjectInfo>(
      'addListRow',
      { sheet: args.sheet, table: args.table, values: args.values },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return listObjectFail('追加行失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已追加 1 行\n' + describeListObject(response.data || {}) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return listObjectFail('追加行出错', errMsg);
  }
};

/** 删行 */
export const deleteListRowDefinition: ToolDefinition = {
  name: 'wps_excel_delete_list_row',
  description: '删除表里的第几行（表体行，从 1 开始，不含表头）。使用场景：剔除一条记录。要按条件删请先 read_range 找到行号，或者直接重写整块数据。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      table: tableParam,
      rowIndex: { type: 'number', description: '删第几行（表体行，从 1 开始，不含表头）' },
      sheet: sheetParam,
    },
    required: ['table', 'rowIndex'],
  },
};

export const deleteListRowHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<ListObjectInfo & { impact?: RangeImpact }>(
      'deleteListRow',
      { sheet: args.sheet, table: args.table, rowIndex: args.rowIndex },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return listObjectFail('删除行失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已删除该行' + impactText(response.data?.impact) + '\n' + describeListObject(response.data || {}) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return listObjectFail('删除行出错', errMsg);
  }
};
/** 改表的设置 */
export const updateListObjectDefinition: ToolDefinition = {
  name: 'wps_excel_update_list_object',
  description: '改表本身的设置：改名、换表格样式、显示或隐藏表头行、显示或隐藏筛选按钮。使用场景：表名要能写进公式、筛选按钮碍事要关掉、换一个配色。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      table: tableParam,
      name: { type: 'string', description: '新的表名（要能写进公式，不能与已有表名或命名范围重复）' },
      tableStyle: { type: 'string', description: '新的表格样式名，如 TableStyleMedium2' },
      showHeaders: { type: 'boolean', description: '是否显示表头行' },
      showAutoFilter: { type: 'boolean', description: '是否显示表头里的筛选按钮' },
      sheet: sheetParam,
    },
    required: ['table'],
  },
};

export const updateListObjectHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<ListObjectInfo>(
      'updateListObject',
      { sheet: args.sheet, table: args.table, name: args.name, tableStyle: args.tableStyle, showHeaders: args.showHeaders, showAutoFilter: args.showAutoFilter },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return listObjectFail('更新表失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已更新\n' + describeListObject(response.data || {}) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return listObjectFail('更新表出错', errMsg);
  }
};

/** 总计行 */
export const setListObjectTotalsDefinition: ToolDefinition = {
  name: 'wps_excel_set_list_object_totals',
  description: '开/关总计行，并可指定某一列的汇总方式。总计行写的是 SUBTOTAL 公式，会随筛选结果变化，这也正是它和普通求和公式的区别。使用场景：给金额列加合计。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      table: tableParam,
      show: { type: 'boolean', description: '是否显示总计行；不填则保持现状（指定 column 时自动打开）' },
      column: { type: 'string', description: '要设置汇总方式的列（列名或序号）；不填只开关总计行' },
      function: { type: 'string', enum: ['sum', 'average', 'count', 'countNums', 'max', 'min', 'stdDev', 'var', 'none'], description: '汇总方式，默认 sum' },
      sheet: sheetParam,
    },
    required: ['table'],
  },
};

export const setListObjectTotalsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<ListObjectInfo>(
      'setListObjectTotals',
      { sheet: args.sheet, table: args.table, show: args.show, column: args.column, function: args.function },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return listObjectFail('设置总计行失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '总计行已更新\n' + describeListObject(response.data || {}) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return listObjectFail('设置总计行出错', errMsg);
  }
};

/** 调整范围 */
export const resizeListObjectDefinition: ToolDefinition = {
  name: 'wps_excel_resize_list_object',
  description: '调整表覆盖的区域（长短变化），表名、样式与结构化引用都保留。使用场景：数据变长了，把表扩到新范围。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      table: tableParam,
      range: { type: 'string', description: '表的新范围，含表头，如 A1:C20' },
      sheet: sheetParam,
    },
    required: ['table', 'range'],
  },
};

export const resizeListObjectHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<ListObjectInfo>(
      'resizeListObject',
      { sheet: args.sheet, table: args.table, range: args.range },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return listObjectFail('调整表范围失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已调整范围\n' + describeListObject(response.data || {}) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return listObjectFail('调整表范围出错', errMsg);
  }
};

/** 转回区域 */
export const unlistListObjectDefinition: ToolDefinition = {
  name: 'wps_excel_unlist_list_object',
  description: '把表转回普通区域（数据与格式保留）：结构化引用、筛选按钮与表对象都会消失。使用场景：交付前清掉表对象，避免对方打开时出现意料之外的引用。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      table: tableParam,
      sheet: sheetParam,
    },
    required: ['table'],
  },
};

export const unlistListObjectHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<ListObjectInfo & { impact?: RangeImpact }>(
      'unlistListObject',
      { sheet: args.sheet, table: args.table },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return listObjectFail('转回区域失败', response.error);
    const info = response.data || {};
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '表「' + (info.name || '') + '」已转回普通区域（' + (info.range || '') + '），数据与格式保留' + impactText(response.data?.impact) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return listObjectFail('转回区域出错', errMsg);
  }
};

export const listObjectTools: RegisteredTool[] = [
  { definition: createListObjectDefinition, handler: createListObjectHandler },
  { definition: getListObjectsDefinition, handler: getListObjectsHandler },
  { definition: addListRowDefinition, handler: addListRowHandler },
  { definition: deleteListRowDefinition, handler: deleteListRowHandler },
  { definition: updateListObjectDefinition, handler: updateListObjectHandler },
  { definition: setListObjectTotalsDefinition, handler: setListObjectTotalsHandler },
  { definition: resizeListObjectDefinition, handler: resizeListObjectHandler },
  { definition: unlistListObjectDefinition, handler: unlistListObjectHandler },
];

export default listObjectTools;

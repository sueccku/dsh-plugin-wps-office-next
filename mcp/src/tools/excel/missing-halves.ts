/**
 * Input: 补全类工具的调用参数
 * Output: 表格「读/清/整」另一半能力的执行结果
 * Pos: Excel 补全工具（P2 第一波）。这些能力在桥里早已实现，只是从来没有工具出口——
 *      上游做 demo 时不需要它们，而真实任务里必然要用（读命名范围、清除格式、查位置…）。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P2 状态。
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

/**
 * 工作表结构与已用范围：让调用方不必猜 A1:Z200
 */
export const getSheetInfoDefinition: ToolDefinition = {
  name: 'wps_excel_get_sheet_info',
  description: `获取工作表的结构信息：工作簿与工作表名、已用范围地址、表头、当前单元格。

使用场景：
- "这张表有多大" / "数据到哪一行"
- 读数据之前先确定范围，而不是猜一个很大的区域

先调用它拿到 usedRange，再用 read_range 精确读取，避免把整片空白也读回来。`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
    },
  },
};

export const getSheetInfoHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      workbookName?: string;
      currentSheet?: string;
      allSheets?: string[];
      selectedCell?: string;
      usedRange?: string;
      headers?: Array<{ column: string; value: unknown }>;
    }>('getExcelContext', { sheet: args.sheet }, WpsAppType.SPREADSHEET);
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取工作表信息失败: ${response.error}` }], error: response.error };
    }
    const d = response.data || {};
    const lines = [`工作簿: ${d.workbookName || '(未知)'}`];
    if (d.currentSheet) lines.push(`当前工作表: ${d.currentSheet}`);
    if (Array.isArray(d.allSheets) && d.allSheets.length) lines.push(`全部工作表: ${d.allSheets.join(', ')}`);
    lines.push(`已用范围: ${d.usedRange || '(空表)'}`);
    if (d.selectedCell) lines.push(`当前单元格: ${d.selectedCell}`);
    const headers = (d.headers || []).filter((h) => h && h.value !== null && h.value !== undefined && h.value !== '');
    if (headers.length) lines.push(`表头: ${headers.map((h) => h.column + '=' + String(h.value)).join(', ')}`);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取工作表信息出错: ${errMsg}` }], error: errMsg };
  }
};

/** 三个自动尺寸工具。刻意写成显式字面量 action：静态提取器要看得见（spec 的键表与守卫依赖它） */
const fitParams = {
  sheet: { type: 'string' as const, description: '工作表名或序号；不填则用当前活动工作表' },
  range: { type: 'string' as const, description: '要调整的区域，如 A1:D20；不填则用整张表的已用范围' },
};

export const autoFitDefinition: ToolDefinition = {
  name: 'wps_excel_auto_fit',
  description: '按内容自动调整列宽与行高。使用场景："列宽太窄看不清"、"让表格自适应内容"。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: { type: 'object', properties: fitParams },
};

export const autoFitHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ message?: string }>(
      'autoFitAll',
      { sheet: args.sheet, range: args.range },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `自动调整失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: response.data?.message || '列宽行高已自动调整' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `自动调整出错: ${errMsg}` }], error: errMsg };
  }
};

export const autoFitColumnsDefinition: ToolDefinition = {
  name: 'wps_excel_auto_fit_columns',
  description: '按内容自动调整列宽（不动行高）。使用 column 可只调整某一列。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: { ...fitParams, column: { type: 'string', description: '只调整这一列（列名如 B，或列号）' } },
  },
};

export const autoFitColumnsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ message?: string }>(
      'autoFitColumn',
      { sheet: args.sheet, range: args.range, column: args.column },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `自动列宽失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: response.data?.message || '列宽已自动调整' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `自动列宽出错: ${errMsg}` }], error: errMsg };
  }
};

export const autoFitRowsDefinition: ToolDefinition = {
  name: 'wps_excel_auto_fit_rows',
  description: '按内容自动调整行高（不动列宽）。使用 row 可只调整某一行。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: { ...fitParams, row: { type: 'number', description: '只调整这一行（从 1 开始）' } },
  },
};

export const autoFitRowsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ message?: string }>(
      'autoFitRow',
      { sheet: args.sheet, range: args.range, row: args.row },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `自动行高失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: response.data?.message || '行高已自动调整' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `自动行高出错: ${errMsg}` }], error: errMsg };
  }
};

/** 自动换行 */
export const setWrapTextDefinition: ToolDefinition = {
  name: 'wps_excel_set_wrap_text',
  description: '设置或取消单元格的自动换行。使用场景："让长文本在单元格里换行显示"。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '目标区域，如 A1:C10' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
      wrap: { type: 'boolean', description: 'true 打开自动换行，false 关闭，默认 true' },
    },
    required: ['range'],
  },
};

export const setWrapTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ range?: string; wrapText?: boolean }>(
      'wrapText',
      { sheet: args.sheet, range: args.range, wrap: args.wrap },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置自动换行失败: ${response.error}` }], error: response.error };
    }
    const on = response.data && response.data.wrapText !== undefined ? response.data.wrapText : true;
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `${response.data?.range || String(args.range)} 的自动换行已${on ? '打开' : '关闭'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置自动换行出错: ${errMsg}` }], error: errMsg };
  }
};

/** 查找并给位置（不改动内容）——与 find_replace 的分工是：那个替换，这个只报告在哪 */
export const findInSheetDefinition: ToolDefinition = {
  name: 'wps_excel_find_in_sheet',
  description: `在工作表里查找文本并返回每一个命中的单元格地址（不改动任何内容）。

使用场景：
- "帮我找一下'华东'出现在哪些格子里"
- 先定位再决定怎么改，比直接替换安全

只统计、不修改。要替换请用 wps_excel_find_replace。`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      searchText: { type: 'string', description: '要查找的文本' },
      range: { type: 'string', description: '查找范围；不填则用已用范围' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
      matchCase: { type: 'boolean', description: '是否区分大小写，默认 false' },
    },
    required: ['searchText'],
  },
};

export const findInSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    // A plain pass-through: the action resolves the used range itself (UsedRange.Address() does work
    // in the resident host - getExcelContext relies on it). Composing getExcelContext here instead
    // used to be needed, but it bought nothing and cost the parameter contract an unread-action
    // violation (category D) because the composed call is not the action the schema describes.
    const response = await wpsClient.executeMethod<{
      searchText?: string;
      results?: Array<{ address: string; value: unknown }>;
      count?: number;
    }>(
      'findInSheet',
      { sheet: args.sheet, range: args.range, searchText: args.searchText, matchCase: args.matchCase },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `查找失败: ${response.error}` }], error: response.error };
    }
    const results = response.data?.results || [];
    if (!results.length) {
      return { id: uuidv4(), success: true, content: [{ type: 'text', text: `没有找到「${String(args.searchText)}」` }] };
    }
    const lines = results.slice(0, 50).map((r) => `${r.address}: ${String(r.value)}`);
    const more = results.length > 50 ? `\n…（共 ${results.length} 处，只列出前 50 处）` : '';
    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: `找到 ${results.length} 处「${String(args.searchText)}」:\n${lines.join('\n')}${more}` }],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `查找出错: ${errMsg}` }], error: errMsg };
  }
};

/** 命名范围：读与删（此前只有写） */
export const getNamedRangesDefinition: ToolDefinition = {
  name: 'wps_excel_get_named_ranges',
  description: '列出工作簿里的全部命名范围及其引用位置。使用场景："这个工作簿里定义了哪些名字"。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: { type: 'object', properties: {} },
};

export const getNamedRangesHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ names?: Array<{ name: string; refersTo: string }>; count?: number }>(
      'getNamedRanges',
      {},
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取命名范围失败: ${response.error}` }], error: response.error };
    }
    const names = response.data?.names || [];
    if (!names.length) {
      return { id: uuidv4(), success: true, content: [{ type: 'text', text: '这个工作簿里没有命名范围。' }] };
    }
    const lines = names.map((n) => `${n.name} = ${n.refersTo}`);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `命名范围（${names.length} 个）:\n${lines.join('\n')}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取命名范围出错: ${errMsg}` }], error: errMsg };
  }
};

export const deleteNamedRangeDefinition: ToolDefinition = {
  name: 'wps_excel_delete_named_range',
  description: '删除指定的命名范围（只删名字，不动单元格内容）。使用场景："把这个没用的名字去掉"。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: { name: { type: 'string', description: '要删除的命名范围名称' } },
    required: ['name'],
  },
};

export const deleteNamedRangeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const name = typeof args.name === 'string' ? args.name : '';
  if (!name) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: 'name 不能为空' }], error: 'name 为空' };
  }
  try {
    const response = await wpsClient.executeMethod<{ deletedName?: string }>(
      'deleteNamedRange',
      { name },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除命名范围失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `命名范围已删除: ${response.data?.deletedName || name}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除命名范围出错: ${errMsg}` }], error: errMsg };
  }
};

export const missingHalfTools: RegisteredTool[] = [
  { definition: getSheetInfoDefinition, handler: getSheetInfoHandler },
  { definition: autoFitDefinition, handler: autoFitHandler },
  { definition: autoFitColumnsDefinition, handler: autoFitColumnsHandler },
  { definition: autoFitRowsDefinition, handler: autoFitRowsHandler },
  { definition: setWrapTextDefinition, handler: setWrapTextHandler },
  { definition: findInSheetDefinition, handler: findInSheetHandler },
  { definition: getNamedRangesDefinition, handler: getNamedRangesHandler },
  { definition: deleteNamedRangeDefinition, handler: deleteNamedRangeHandler },
];

export default missingHalfTools;

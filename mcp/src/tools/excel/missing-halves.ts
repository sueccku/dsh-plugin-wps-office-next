/**
 * Input: 补全类工具的调用参数
 * Output: 表格「读/清/整」另一半能力的执行结果
 * Pos: Excel 补全工具（P2 第一波 + 余项，共 18 个）。这些能力在桥里早已实现，只是从来没有工具出口——
 *      上游做 demo 时不需要它们，而真实任务里必然要用（读命名范围、清除格式、查位置…）。
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

// ===== P2 第一波余项（10 个）：挂出桥里已有、却一直没出口的 Excel 能力 =====

/** 格式刷：只搬格式，不动值与公式 */
export const copyFormatDefinition: ToolDefinition = {
  name: 'wps_excel_copy_format',
  description: '把一块区域的格式复制到另一块区域（只复制格式，不改数值与公式）。使用场景：把 A1 的样式刷到整个 A 列、统一表头外观。要连值一起搬请用 wps_excel_write_range。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      source: { type: 'string', description: '格式来源区域，如 A1' },
      target: { type: 'string', description: '格式目标区域，如 A2:A100' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
    },
    required: ['source', 'target'],
  },
};

export const copyFormatHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ source?: string; target?: string }>(
      'copyFormat',
      { sheet: args.sheet, source: args.source, target: args.target },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '复制格式失败: ' + response.error }], error: response.error };
    }
    const source = response.data?.source || String(args.source);
    const target = response.data?.target || String(args.target);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已把 ' + source + ' 的格式复制到 ' + target + '（只改外观，值不变）' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '复制格式出错: ' + errMsg }], error: errMsg };
  }
};

/** 清除格式：内容留着，外观回到默认 */
export const clearFormatsDefinition: ToolDefinition = {
  name: 'wps_excel_clear_formats',
  description: '清除区域的格式（字体、颜色、边框、数字格式），单元格内容保留。使用场景：格式被弄乱了，恢复成默认样子。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '目标区域，如 A1:D20' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
    },
    required: ['range'],
  },
};

export const clearFormatsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ range?: string; impact?: RangeImpact }>(
      'clearFormats',
      { sheet: args.sheet, range: args.range },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '清除格式失败: ' + response.error }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: (response.data?.range || String(args.range)) + ' 的格式已清除，内容保留' + impactText(response.data?.impact) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '清除格式出错: ' + errMsg }], error: errMsg };
  }
};

/** 条件格式：读与删（新增侧早就有 wps_excel_set_conditional_format） */
const cfTypeNames: Record<number, string> = { 1: '单元格值', 2: '公式', 3: '色阶', 4: '数据条', 5: '前 10 项', 6: '图标集' };

export const getConditionalFormatsDefinition: ToolDefinition = {
  name: 'wps_excel_get_conditional_formats',
  description: '列出区域上生效的条件格式规则（序号 + 类型），用于先看清楚再改。使用场景：这个表为什么某些格子会变红；删规则之前先确认删哪一条。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '要查看的区域，如 A1:A100' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
    },
    required: ['range'],
  },
};

export const getConditionalFormatsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ range?: string; formats?: Array<{ index: number; type: number }>; count?: number }>(
      'getConditionalFormats',
      { sheet: args.sheet, range: args.range },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '读取条件格式失败: ' + response.error }], error: response.error };
    }
    const formats = response.data?.formats || [];
    const where = response.data?.range || String(args.range);
    if (!formats.length) {
      return { id: uuidv4(), success: true, content: [{ type: 'text', text: where + ' 上没有条件格式规则。' }] };
    }
    const lines = formats.map((f) => {
      const name = cfTypeNames[Number(f.type)];
      return '  ' + f.index + '. 类型 ' + f.type + (name ? '（' + name + '）' : '');
    });
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: where + ' 上的条件格式（' + formats.length + ' 条）:\n' + lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '读取条件格式出错: ' + errMsg }], error: errMsg };
  }
};

export const removeConditionalFormatDefinition: ToolDefinition = {
  name: 'wps_excel_remove_conditional_format',
  description: '删除区域上的条件格式规则。给 index 删指定的一条（序号见 wps_excel_get_conditional_formats），不填则删该区域的全部规则。只删规则，不动内容与普通格式。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '目标区域，如 A1:A100' },
      index: { type: 'number', description: '只删第几条规则（从 1 开始）；不填则删全部' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
    },
    required: ['range'],
  },
};

export const removeConditionalFormatHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ range?: string }>(
      'removeConditionalFormat',
      { sheet: args.sheet, range: args.range, index: args.index },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '删除条件格式失败: ' + response.error }], error: response.error };
    }
    const scope = args.index === undefined || args.index === null ? '全部规则' : '第 ' + String(args.index) + ' 条规则';
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: (response.data?.range || String(args.range)) + ' 的条件格式已删除（' + scope + '）' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '删除条件格式出错: ' + errMsg }], error: errMsg };
  }
};

/** 数据验证：读与删（新增侧早就有 wps_excel_set_data_validation） */
const dvTypeNames: Record<number, string> = { 1: '整数', 2: '小数', 3: '序列', 4: '日期', 5: '时间', 6: '文本长度', 7: '自定义' };

export const getDataValidationsDefinition: ToolDefinition = {
  name: 'wps_excel_get_data_validations',
  description: '读取区域上的数据验证规则（类型、来源公式、提示语）。使用场景：这个下拉框的选项是从哪来的；删规则之前先确认规则内容。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '要查看的区域，如 B2:B100' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
    },
    required: ['range'],
  },
};

export const getDataValidationsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ range?: string; type?: number; formula1?: string; formula2?: string; inputTitle?: string; inputMessage?: string }>(
      'getDataValidations',
      { sheet: args.sheet, range: args.range },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '读取数据验证失败: ' + response.error }], error: response.error };
    }
    const d = response.data || {};
    const where = d.range || String(args.range);
    const typeName = dvTypeNames[Number(d.type)];
    const lines = [where + ' 的数据验证: 类型 ' + d.type + (typeName ? '（' + typeName + '）' : '')];
    if (d.formula1) lines.push('  来源/条件: ' + d.formula1);
    if (d.formula2) lines.push('  条件 2: ' + d.formula2);
    if (d.inputTitle) lines.push('  输入标题: ' + d.inputTitle);
    if (d.inputMessage) lines.push('  输入提示: ' + d.inputMessage);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '读取数据验证出错: ' + errMsg }], error: errMsg };
  }
};

export const removeDataValidationDefinition: ToolDefinition = {
  name: 'wps_excel_remove_data_validation',
  description: '删除区域上的数据验证规则（下拉框、输入限制）。使用场景：去掉这列的下拉限制。只删规则，不动单元格内容。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '目标区域，如 B2:B100' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
    },
    required: ['range'],
  },
};

export const removeDataValidationHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ range?: string }>(
      'removeDataValidation',
      { sheet: args.sheet, range: args.range },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '删除数据验证失败: ' + response.error }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: (response.data?.range || String(args.range)) + ' 的数据验证已删除' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '删除数据验证出错: ' + errMsg }], error: errMsg };
  }
};

/** 外部链接刷新 */
export const refreshLinksDefinition: ToolDefinition = {
  name: 'wps_excel_refresh_links',
  description: '刷新工作簿引用的全部外部链接并报告条数。使用场景：数据源文件更新了，把引用拉一遍。没有链接时如实报告 0 条。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: { type: 'object', properties: {} },
};

export const refreshLinksHandler: ToolHandler = async (_args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ refreshed?: number; message?: string }>(
      'refreshLinks',
      {},
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '刷新外部链接失败: ' + response.error }], error: response.error };
    }
    const n = response.data?.refreshed || 0;
    const detail = response.data?.message ? '（' + response.data.message + '）' : '';
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已刷新外部链接 ' + n + ' 条' + detail }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '刷新外部链接出错: ' + errMsg }], error: errMsg };
  }
};

/** 合并计算：把多块来源区域汇总到目标区域 */
export const consolidateDefinition: ToolDefinition = {
  name: 'wps_excel_consolidate',
  description: '把多块来源区域按指定函数汇总写入目标区域（Excel 的合并计算）。sources 形如 [Sheet1!A1:B4, Sheet2!A1:B4]。使用场景：多张同结构表加总到一张。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      destination: { type: 'string', description: '汇总结果写入的区域左上角，如 E1' },
      sources: { type: 'array', items: { type: 'string' }, description: '来源区域列表，如 [Sheet1!A1:B4, Sheet2!A1:B4]' },
      function: { type: 'string', enum: ['sum', 'count', 'average', 'max', 'min'], description: '汇总函数，默认 sum' },
      topRow: { type: 'boolean', description: '按标签合并：来源首行是标题。默认 false（按位置逐格相加，纯数字表用这个）' },
      leftColumn: { type: 'boolean', description: '按标签合并：来源首列是标题。默认 false' },
      createLinks: { type: 'boolean', description: '与来源建立链接，默认 false' },
      sheet: { type: 'string', description: '目标工作表名或序号；不填则用当前活动工作表' },
    },
    required: ['destination', 'sources'],
  },
};

export const consolidateHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ destination?: string; sources?: string[] }>(
      'consolidate',
      {
        sheet: args.sheet,
        destination: args.destination,
        sources: args.sources,
        function: args.function,
        topRow: args.topRow === undefined ? false : args.topRow,
        leftColumn: args.leftColumn === undefined ? false : args.leftColumn,
        createLinks: args.createLinks === undefined ? false : args.createLinks,
      },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '合并计算失败: ' + response.error }], error: response.error };
    }
    const sources = response.data?.sources || [];
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已把 ' + sources.length + ' 块来源合并到 ' + (response.data?.destination || String(args.destination)) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '合并计算出错: ' + errMsg }], error: errMsg };
  }
};

/** 强制重算 */
export const calculateDefinition: ToolDefinition = {
  name: 'wps_excel_calculate',
  description: '强制重算公式。all 为 true 时重算整个工作簿，否则只重算指定工作表。使用场景：刚写入公式要立刻拿结果；表格显示的是过期值（手动计算模式）。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      all: { type: 'boolean', description: 'true 重算整个工作簿；默认 false 只重算一张表' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表（all 为 true 时忽略）' },
    },
  },
};

export const calculateHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ calculated?: string }>(
      'calculateSheet',
      { sheet: args.sheet, all: args.all },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '重算失败: ' + response.error }], error: response.error };
    }
    const what = response.data?.calculated === 'all' ? '整个工作簿' : '工作表 ' + String(response.data?.calculated || '');
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已重算 ' + what }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '重算出错: ' + errMsg }], error: errMsg };
  }
};

/** 列分组（分级显示）——行分组早有 wps_excel_group_rows */
export const groupColumnsDefinition: ToolDefinition = {
  name: 'wps_excel_group_columns',
  description: '把一段列折叠分组（分级显示）。列名如 B、E 或列号 2、5。使用场景：把中间的计算列收起来，只留结果列。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      startColumn: { type: 'string', description: '起始列（列名如 B，或列号）' },
      endColumn: { type: 'string', description: '结束列（列名如 E，或列号）' },
      sheet: { type: 'string', description: '工作表名或序号；不填则用当前活动工作表' },
    },
    required: ['startColumn', 'endColumn'],
  },
};

export const groupColumnsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ grouped?: string }>(
      'groupColumns',
      { sheet: args.sheet, startColumn: args.startColumn, endColumn: args.endColumn },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '列分组失败: ' + response.error }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已分组列 ' + (response.data?.grouped || '') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '列分组出错: ' + errMsg }], error: errMsg };
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
  { definition: copyFormatDefinition, handler: copyFormatHandler },
  { definition: clearFormatsDefinition, handler: clearFormatsHandler },
  { definition: getConditionalFormatsDefinition, handler: getConditionalFormatsHandler },
  { definition: removeConditionalFormatDefinition, handler: removeConditionalFormatHandler },
  { definition: getDataValidationsDefinition, handler: getDataValidationsHandler },
  { definition: removeDataValidationDefinition, handler: removeDataValidationHandler },
  { definition: refreshLinksDefinition, handler: refreshLinksHandler },
  { definition: consolidateDefinition, handler: consolidateHandler },
  { definition: calculateDefinition, handler: calculateHandler },
  { definition: groupColumnsDefinition, handler: groupColumnsHandler },
];

export default missingHalfTools;

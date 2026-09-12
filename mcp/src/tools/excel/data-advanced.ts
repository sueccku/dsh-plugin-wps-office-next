/**
 * Input: 高级数据处理工具参数
 * Output: 数据处理结果
 * Pos: Excel 高级数据处理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
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
 * 自动筛选
 */
export const autoFilterDefinition: ToolDefinition = {
  name: 'wps_excel_auto_filter',
  description: '对Excel指定范围应用自动筛选。',
  category: ToolCategory.SPREADSHEET,
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

export const autoFilterHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, column, criteria, sheet } = args as {
    range: string; column?: string; criteria?: string; sheet?: string;
  };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'autoFilter',
      { range, column, criteria, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `自动筛选失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `自动筛选已应用: ${range}${column ? '，列: ' + column : ''}${criteria ? '，条件: ' + criteria : ''}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `自动筛选出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 复制范围
 */
export const copyRangeDefinition: ToolDefinition = {
  name: 'wps_excel_copy_range',
  description: '复制Excel指定范围到目标位置。',
  category: ToolCategory.SPREADSHEET,
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

export const copyRangeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { source, destination, sheet } = args as {
    source: string; destination: string; sheet?: string;
  };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'copyRange',
      { source, destination, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `复制范围失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `已复制 ${source} 到 ${destination}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `复制范围出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 粘贴范围
 */
export const pasteRangeDefinition: ToolDefinition = {
  name: 'wps_excel_paste_range',
  description: '粘贴已复制的内容到指定位置。',
  category: ToolCategory.SPREADSHEET,
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

export const pasteRangeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { destination, pasteType, sheet } = args as {
    destination: string; pasteType?: string; sheet?: string;
  };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'pasteRange',
      { destination, pasteType: pasteType || 'all', sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `粘贴失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `已粘贴到 ${destination}，类型: ${pasteType || 'all'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `粘贴出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 填充序列
 */
export const fillSeriesDefinition: ToolDefinition = {
  name: 'wps_excel_fill_series',
  description: '自动填充序列数据。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '填充范围，如 A1:A10' },
      direction: { type: 'string', description: '填充方向', enum: ['down', 'right', 'up', 'left'] },
      type: { type: 'string', description: '填充类型', enum: ['linear', 'growth', 'date', 'auto'] },
      step: { type: 'number', description: '步长值' },
      sheet: { type: 'string', description: '工作表名称' },
    },
    required: ['range'],
  },
};

export const fillSeriesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, direction, type, step, sheet } = args as {
    range: string; direction?: string; type?: string; step?: number; sheet?: string;
  };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'fillSeries',
      { range, direction: direction || 'down', type: type || 'auto', step, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `填充序列失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `序列填充完成: ${range}，方向: ${direction || 'down'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `填充序列出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 转置数据
 */
export const transposeDefinition: ToolDefinition = {
  name: 'wps_excel_transpose',
  description: '转置数据（行列互换）。',
  category: ToolCategory.SPREADSHEET,
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

export const transposeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { source, destination, sheet } = args as {
    source: string; destination: string; sheet?: string;
  };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'transpose',
      { source, destination, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `转置失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `数据已转置: ${source} → ${destination}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `转置出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 分列
 */
export const textToColumnsDefinition: ToolDefinition = {
  name: 'wps_excel_text_to_columns',
  description: '将文本按分隔符拆分到多列。',
  category: ToolCategory.SPREADSHEET,
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

export const textToColumnsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, delimiter, sheet } = args as {
    range: string; delimiter?: string; sheet?: string;
  };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'textToColumns',
      { range, delimiter: delimiter || ',', sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `分列失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `分列完成: ${range}，分隔符: "${delimiter || ','}"` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `分列出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 分类汇总
 */
export const subtotalDefinition: ToolDefinition = {
  name: 'wps_excel_subtotal',
  description: '创建分类汇总。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '数据范围，如 A1:D100' },
      groupBy: { type: 'string', description: '分组列标识' },
      function: { type: 'string', description: '汇总函数', enum: ['sum', 'count', 'average', 'max', 'min'] },
      columns: {
        type: 'array',
        items: { type: 'string' },
        description: '要汇总的列标识列表',
      },
      sheet: { type: 'string', description: '工作表名称' },
    },
    required: ['range', 'groupBy', 'function', 'columns'],
  },
};

export const subtotalHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, groupBy, function: func, columns, sheet } = args as {
    range: string; groupBy: string; function: string; columns: string[]; sheet?: string;
  };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'subtotal',
      { range, groupBy, function: func, columns, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `分类汇总失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `分类汇总完成: 按 ${groupBy} 分组，${func} 汇总 [${columns.join(', ')}]` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `分类汇总出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 导出所有高级数据处理相关的Tools
 */
export const dataAdvancedTools: RegisteredTool[] = [
  { definition: autoFilterDefinition, handler: autoFilterHandler },
  { definition: copyRangeDefinition, handler: copyRangeHandler },
  { definition: pasteRangeDefinition, handler: pasteRangeHandler },
  { definition: fillSeriesDefinition, handler: fillSeriesHandler },
  { definition: transposeDefinition, handler: transposeHandler },
  { definition: textToColumnsDefinition, handler: textToColumnsHandler },
  { definition: subtotalDefinition, handler: subtotalHandler },
];

export default dataAdvancedTools;

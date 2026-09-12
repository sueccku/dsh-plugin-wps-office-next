/**
 * Input: 行列操作工具参数
 * Output: 行列插入/删除/隐藏/分组结果
 * Pos: Excel 行列操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel行列操作Tools - 行列管理模块
 * 处理行列的插入、删除、隐藏、显示、分组等操作
 *
 * 包含：
 * - wps_excel_insert_rows: 插入多行
 * - wps_excel_insert_columns: 插入多列
 * - wps_excel_delete_rows: 删除多行
 * - wps_excel_delete_columns: 删除多列
 * - wps_excel_hide_rows: 隐藏行
 * - wps_excel_show_rows: 显示行
 * - wps_excel_show_columns: 显示列
 * - wps_excel_group_rows: 分组行
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
 * 插入多行
 */
export const insertRowsDefinition: ToolDefinition = {
  name: 'wps_excel_insert_rows',
  description: '在Excel中指定位置插入一行或多行。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      row: { type: 'number', description: '在第几行前插入（从1开始）' },
      count: { type: 'number', description: '插入行数，默认1' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['row'],
  },
};

export const insertRowsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { row, count, sheet } = args as { row: number; count?: number; sheet?: string };
  const insertCount = count || 1;
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'insertRows',
      { row, count: insertCount, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入行失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `插入行完成！在第${row}行前插入了${insertCount}行` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入行出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 插入多列
 */
export const insertColumnsDefinition: ToolDefinition = {
  name: 'wps_excel_insert_columns',
  description: '在Excel中指定位置插入一列或多列。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      column: { type: 'string', description: '在哪一列前插入，如 A、B、C' },
      count: { type: 'number', description: '插入列数，默认1' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['column'],
  },
};

export const insertColumnsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { column, count, sheet } = args as { column: string; count?: number; sheet?: string };
  const insertCount = count || 1;
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'insertColumns',
      { column, count: insertCount, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入列失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `插入列完成！在${column}列前插入了${insertCount}列` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入列出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 删除多行
 */
export const deleteRowsDefinition: ToolDefinition = {
  name: 'wps_excel_delete_rows',
  description: '删除Excel中指定位置的一行或多行。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      startRow: { type: 'number', description: '起始行号（从1开始）' },
      count: { type: 'number', description: '删除行数，默认1' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['startRow'],
  },
};

export const deleteRowsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { startRow, count, sheet } = args as { startRow: number; count?: number; sheet?: string };
  const deleteCount = count || 1;
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'deleteRows',
      { startRow, count: deleteCount, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除行失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `删除行完成！从第${startRow}行开始删除了${deleteCount}行` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除行出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 删除多列
 */
export const deleteColumnsDefinition: ToolDefinition = {
  name: 'wps_excel_delete_columns',
  description: '删除Excel中指定位置的一列或多列。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      column: { type: 'string', description: '起始列字母，如 A、B、C' },
      count: { type: 'number', description: '删除列数，默认1' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['column'],
  },
};

export const deleteColumnsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { column, count, sheet } = args as { column: string; count?: number; sheet?: string };
  const deleteCount = count || 1;
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'deleteColumns',
      { column, count: deleteCount, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除列失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `删除列完成！从${column}列开始删除了${deleteCount}列` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除列出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 隐藏行
 */
export const hideRowsDefinition: ToolDefinition = {
  name: 'wps_excel_hide_rows',
  description: '隐藏Excel中指定范围的行。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      startRow: { type: 'number', description: '起始行号（从1开始）' },
      endRow: { type: 'number', description: '结束行号（从1开始）' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['startRow', 'endRow'],
  },
};

export const hideRowsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { startRow, endRow, sheet } = args as { startRow: number; endRow: number; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'hideRows',
      { startRow, endRow, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `隐藏行失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `隐藏行完成！已隐藏第${startRow}行到第${endRow}行` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `隐藏行出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 显示行
 */
export const showRowsDefinition: ToolDefinition = {
  name: 'wps_excel_show_rows',
  description: '显示Excel中已隐藏的行。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      startRow: { type: 'number', description: '起始行号（从1开始）' },
      endRow: { type: 'number', description: '结束行号（从1开始）' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['startRow', 'endRow'],
  },
};

export const showRowsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { startRow, endRow, sheet } = args as { startRow: number; endRow: number; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'showRows',
      { startRow, endRow, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `显示行失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `显示行完成！已显示第${startRow}行到第${endRow}行` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `显示行出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 显示列
 */
export const showColumnsDefinition: ToolDefinition = {
  name: 'wps_excel_show_columns',
  description: '显示Excel中已隐藏的列。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      startColumn: { type: 'string', description: '起始列字母，如 A' },
      endColumn: { type: 'string', description: '结束列字母，如 D' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['startColumn', 'endColumn'],
  },
};

export const showColumnsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { startColumn, endColumn, sheet } = args as { startColumn: string; endColumn: string; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'showColumns',
      { startColumn, endColumn, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `显示列失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `显示列完成！已显示${startColumn}列到${endColumn}列` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `显示列出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 分组行
 */
export const groupRowsDefinition: ToolDefinition = {
  name: 'wps_excel_group_rows',
  description: '对Excel中指定范围的行进行分组，便于折叠/展开管理。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      startRow: { type: 'number', description: '起始行号（从1开始）' },
      endRow: { type: 'number', description: '结束行号（从1开始）' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['startRow', 'endRow'],
  },
};

export const groupRowsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { startRow, endRow, sheet } = args as { startRow: number; endRow: number; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'groupRows',
      { startRow, endRow, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `分组行失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `分组行完成！已将第${startRow}行到第${endRow}行分组` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `分组行出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 导出所有行列操作相关的Tools
 */
export const rowColumnTools: RegisteredTool[] = [
  { definition: insertRowsDefinition, handler: insertRowsHandler },
  { definition: insertColumnsDefinition, handler: insertColumnsHandler },
  { definition: deleteRowsDefinition, handler: deleteRowsHandler },
  { definition: deleteColumnsDefinition, handler: deleteColumnsHandler },
  { definition: hideRowsDefinition, handler: hideRowsHandler },
  { definition: showRowsDefinition, handler: showRowsHandler },
  { definition: showColumnsDefinition, handler: showColumnsHandler },
  { definition: groupRowsDefinition, handler: groupRowsHandler },
];

export default rowColumnTools;

/**
 * Input: 批注/保护/公式/图片/超链接工具参数
 * Output: 批注删除/获取、取消保护、锁定单元格、数组公式、图片插入、超链接设置结果
 * Pos: Excel 批注保护及扩展工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel批注保护Tools - 批注、保护与扩展功能模块
 * 处理批注管理、工作表取消保护、单元格锁定、数组公式、图片插入、超链接等操作
 *
 * 包含：
 * - wps_excel_delete_cell_comment: 删除单元格批注
 * - wps_excel_get_cell_comments: 获取单元格批注
 * - wps_excel_unprotect_sheet: 取消保护工作表
 * - wps_excel_lock_cells: 锁定/解锁单元格
 * - wps_excel_set_array_formula: 设置数组公式
 * - wps_excel_insert_excel_image: 插入图片
 * - wps_excel_set_hyperlink: 设置超链接
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
 * 删除单元格批注
 */
export const deleteCellCommentDefinition: ToolDefinition = {
  name: 'wps_excel_delete_cell_comment',
  description: '删除Excel单元格上的批注。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      cell: { type: 'string', description: '单元格地址，如 A1、B2' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['cell'],
  },
};

export const deleteCellCommentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { cell, sheet } = args as { cell: string; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'deleteCellComment',
      { cell, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除批注失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `批注删除成功！单元格: ${cell}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除批注出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 获取单元格批注
 */
export const getCellCommentsDefinition: ToolDefinition = {
  name: 'wps_excel_get_cell_comments',
  description: '获取Excel指定范围内的所有批注。不指定范围则获取当前工作表所有批注。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '要查询的范围，如 A1:D10。不填则获取所有批注' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: [],
  },
};

export const getCellCommentsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, sheet } = args as { range?: string; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{
      comments: Array<{ cell: string; comment: string; author?: string }>;
    }>(
      'getCellComments',
      { range, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success || !response.data) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取批注失败: ${response.error}` }], error: response.error };
    }
    const comments = response.data.comments;
    if (!comments || comments.length === 0) {
      return { id: uuidv4(), success: true, content: [{ type: 'text', text: `${range ? `范围 ${range}` : '当前工作表'}没有批注` }] };
    }
    let output = `找到${comments.length}条批注：\n\n`;
    comments.forEach((c) => {
      output += `- ${c.cell}: ${c.comment}${c.author ? ` (作者: ${c.author})` : ''}\n`;
    });
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: output }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取批注出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 取消保护工作表
 */
export const unprotectSheetDefinition: ToolDefinition = {
  name: 'wps_excel_unprotect_sheet',
  description: '取消保护当前工作表。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      password: { type: 'string', description: '保护密码（如果设置了密码保护则需要提供）' },
    },
    required: [],
  },
};

export const unprotectSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { password } = args as { password?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'unprotectSheet',
      { password },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `取消保护工作表失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '工作表取消保护成功！' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `取消保护工作表出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 锁定/解锁单元格
 */
export const lockCellsDefinition: ToolDefinition = {
  name: 'wps_excel_lock_cells',
  description: '锁定或解锁Excel指定范围的单元格（需配合工作表保护使用）。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '要锁定/解锁的范围，如 A1:D10' },
      locked: { type: 'boolean', description: 'true为锁定，false为解锁' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['range', 'locked'],
  },
};

export const lockCellsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, locked, sheet } = args as { range: string; locked: boolean; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'lockCells',
      { range, locked, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `${locked ? '锁定' : '解锁'}单元格失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `单元格${locked ? '锁定' : '解锁'}成功！范围: ${range}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `${locked ? '锁定' : '解锁'}单元格出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 设置数组公式
 */
export const setArrayFormulaDefinition: ToolDefinition = {
  name: 'wps_excel_set_array_formula',
  description: '为Excel指定范围设置数组公式（CSE数组公式）。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '数组公式应用的范围，如 A1:A10' },
      formula: { type: 'string', description: '数组公式，如 =A1:A10*B1:B10' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['range', 'formula'],
  },
};

export const setArrayFormulaHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, formula, sheet } = args as { range: string; formula: string; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'setArrayFormula',
      { range, formula, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置数组公式失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `数组公式设置成功！范围: ${range}，公式: ${formula}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置数组公式出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 插入图片到Excel
 */
export const insertExcelImageDefinition: ToolDefinition = {
  name: 'wps_excel_insert_excel_image',
  description: '在Excel中插入图片到指定位置。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      filePath: { type: 'string', description: '图片文件路径' },
      cell: { type: 'string', description: '插入位置的单元格地址，如 A1。不填则插入到当前选中位置' },
      width: { type: 'number', description: '图片宽度（像素），不填则使用原始宽度' },
      height: { type: 'number', description: '图片高度（像素），不填则使用原始高度' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['filePath'],
  },
};

export const insertExcelImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { filePath, cell, width, height, sheet } = args as {
    filePath: string;
    cell?: string;
    width?: number;
    height?: number;
    sheet?: string;
  };
  try {
    // 跨平台参数对齐：macOS/Windows 底层均读取 params.path，需同时发送 path/imagePath 别名
    const response = await wpsClient.executeMethod<{ message: string }>(
      'insertExcelImage',
      { filePath, path: filePath, imagePath: filePath, cell, width, height, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入图片失败: ${response.error}` }], error: response.error };
    }
    let text = `图片插入成功！文件: ${filePath}`;
    if (cell) text += `，位置: ${cell}`;
    if (width || height) text += `，尺寸: ${width || '自动'}x${height || '自动'}`;
    return { id: uuidv4(), success: true, content: [{ type: 'text', text }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `插入图片出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 设置超链接
 */
export const setHyperlinkDefinition: ToolDefinition = {
  name: 'wps_excel_set_hyperlink',
  description: '为Excel单元格设置超链接。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      cell: { type: 'string', description: '单元格地址，如 A1' },
      url: { type: 'string', description: '超链接URL地址' },
      text: { type: 'string', description: '显示文本，不填则显示URL' },
      sheet: { type: 'string', description: '工作表名称，不填则使用当前活动工作表' },
    },
    required: ['cell', 'url'],
  },
};

export const setHyperlinkHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { cell, url, text, sheet } = args as { cell: string; url: string; text?: string; sheet?: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'setHyperlink',
      { cell, url, text, sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置超链接失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `超链接设置成功！单元格: ${cell}，URL: ${url}${text ? `，显示文本: ${text}` : ''}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置超链接出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 导出所有批注保护相关的Tools
 */
export const commentProtectTools: RegisteredTool[] = [
  { definition: deleteCellCommentDefinition, handler: deleteCellCommentHandler },
  { definition: getCellCommentsDefinition, handler: getCellCommentsHandler },
  { definition: unprotectSheetDefinition, handler: unprotectSheetHandler },
  { definition: lockCellsDefinition, handler: lockCellsHandler },
  { definition: setArrayFormulaDefinition, handler: setArrayFormulaHandler },
  { definition: insertExcelImageDefinition, handler: insertExcelImageHandler },
  { definition: setHyperlinkDefinition, handler: setHyperlinkHandler },
];

export default commentProtectTools;

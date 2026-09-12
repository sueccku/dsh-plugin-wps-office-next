/**
 * Input: 工作表管理工具参数
 * Output: 工作表操作结果
 * Pos: Excel 工作表管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel工作表管理Tools
 * 工作表的创建、删除、重命名、复制、移动、切换等操作
 *
 * 包含：
 * - wps_excel_create_sheet: 创建新工作表
 * - wps_excel_delete_sheet: 删除指定工作表
 * - wps_excel_rename_sheet: 重命名工作表
 * - wps_excel_copy_sheet: 复制工作表
 * - wps_excel_get_sheet_list: 获取工作表列表
 * - wps_excel_switch_sheet: 切换工作表
 * - wps_excel_move_sheet: 移动工作表
 * - wps_excel_get_selection: 获取当前选中区域
 * - wps_excel_delete_row: 删除指定行
 * - wps_excel_insert_column: 插入列
 * - wps_excel_delete_column: 删除指定列
 * - wps_excel_freeze_panes: 冻结/取消冻结窗格
 * - wps_excel_auto_fill: 自动填充单元格区域
 * - wps_excel_set_named_range: 设置命名范围
 * - wps_excel_hide_column: 隐藏/显示列
 * - wps_excel_auto_sum: 对指定列/行自动求和
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
 * 创建新工作表
 */
export const createSheetDefinition: ToolDefinition = {
  name: 'wps_excel_create_sheet',
  description: '在当前工作簿中创建新的工作表。可指定名称和插入位置。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '新工作表的名称',
      },
      position: {
        type: 'number',
        description: '插入位置索引（从0开始），不填则添加到末尾',
      },
    },
    required: ['name'],
  },
};

export const createSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, position } = args as {
    name: string;
    position?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      name: string;
      index: number;
    }>(
      'createSheet',
      { name, position },
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表创建成功！\n名称: ${response.data.name}\n位置: 第${response.data.index + 1}个`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 删除指定工作表
 */
export const deleteSheetDefinition: ToolDefinition = {
  name: 'wps_excel_delete_sheet',
  description: '删除当前工作簿中的指定工作表。注意：此操作不可撤销。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要删除的工作表名称',
      },
    },
    required: ['name'],
  },
};

export const deleteSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };

  try {
    const response = await wpsClient.executeMethod<{
      deleted: string;
    }>(
      'deleteSheet',
      { name },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表 "${name}" 已成功删除`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 重命名工作表
 */
export const renameSheetDefinition: ToolDefinition = {
  name: 'wps_excel_rename_sheet',
  description: '重命名当前工作簿中的指定工作表。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      oldName: {
        type: 'string',
        description: '当前工作表名称',
      },
      newName: {
        type: 'string',
        description: '新的工作表名称',
      },
    },
    required: ['oldName', 'newName'],
  },
};

export const renameSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { oldName, newName } = args as {
    oldName: string;
    newName: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      oldName: string;
      newName: string;
    }>(
      'renameSheet',
      { oldName, newName },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `重命名工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表重命名成功！\n"${oldName}" → "${newName}"`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `重命名工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 复制工作表
 */
export const copySheetDefinition: ToolDefinition = {
  name: 'wps_excel_copy_sheet',
  description: '复制当前工作簿中的指定工作表。可指定新名称和插入位置。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要复制的工作表名称',
      },
      newName: {
        type: 'string',
        description: '复制后的工作表名称，不填则自动生成',
      },
      position: {
        type: 'number',
        description: '插入位置索引（从0开始），不填则添加到末尾',
      },
    },
    required: ['name'],
  },
};

export const copySheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, newName, position } = args as {
    name: string;
    newName?: string;
    position?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      sourceName: string;
      newName: string;
      index: number;
    }>(
      'copySheet',
      { name, newName, position },
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `复制工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表复制成功！\n源工作表: ${response.data.sourceName}\n新工作表: ${response.data.newName}\n位置: 第${response.data.index + 1}个`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `复制工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 获取工作表列表
 */
export const getSheetListDefinition: ToolDefinition = {
  name: 'wps_excel_get_sheet_list',
  description: '获取当前工作簿的所有工作表列表，包含名称、索引和是否为活动工作表。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const getSheetListHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      sheets: Array<{
        name: string;
        index: number;
        active: boolean;
      }>;
      count: number;
    }>(
      'getSheetList',
      {},
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取工作表列表失败: ${response.error}` }],
        error: response.error,
      };
    }

    const { sheets, count } = response.data;
    let output = `当前工作簿共有 ${count} 个工作表：\n\n`;

    sheets.forEach((sheet) => {
      const activeFlag = sheet.active ? ' [活动]' : '';
      output += `${sheet.index + 1}. ${sheet.name}${activeFlag}\n`;
    });

    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: output }],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取工作表列表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 切换到指定工作表
 */
export const switchSheetDefinition: ToolDefinition = {
  name: 'wps_excel_switch_sheet',
  description: '切换到指定的工作表，使其成为活动工作表。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要切换到的工作表名称',
      },
    },
    required: ['name'],
  },
};

export const switchSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };

  try {
    const response = await wpsClient.executeMethod<{
      activatedSheet: string;
    }>(
      'switchSheet',
      { name },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `已切换到工作表 "${name}"`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 移动工作表到指定位置
 */
export const moveSheetDefinition: ToolDefinition = {
  name: 'wps_excel_move_sheet',
  description: '移动指定工作表到新的位置。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要移动的工作表名称',
      },
      position: {
        type: 'number',
        description: '目标位置索引（从0开始）',
      },
    },
    required: ['name', 'position'],
  },
};

export const moveSheetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, position } = args as {
    name: string;
    position: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      name: string;
      newPosition: number;
    }>(
      'moveSheet',
      { name, position },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `移动工作表失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `工作表 "${name}" 已移动到第${position + 1}个位置`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `移动工作表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 获取当前选中区域信息
 */
export const getSelectionDefinition: ToolDefinition = {
  name: 'wps_excel_get_selection',
  description: '获取当前Excel中选中区域的信息，包括范围地址、行列数等。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const getSelectionHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      address: string;
      rowCount: number;
      columnCount: number;
      sheet: string;
    }>(
      'getSelection',
      {},
      WpsAppType.SPREADSHEET
    );

    if (!response.success || !response.data) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取选中区域失败: ${response.error}` }],
        error: response.error,
      };
    }

    const { address, rowCount, columnCount, sheet } = response.data;

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `当前选中区域信息：\n工作表: ${sheet}\n范围: ${address}\n行数: ${rowCount}\n列数: ${columnCount}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取选中区域出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 删除指定行
 */
export const deleteRowDefinition: ToolDefinition = {
  name: 'wps_excel_delete_row',
  description: '删除指定行。可指定起始行号和删除行数。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      row: {
        type: 'number',
        description: '要删除的起始行号（从1开始）',
      },
      count: {
        type: 'number',
        description: '要删除的行数，默认1',
      },
    },
    required: ['row'],
  },
};

export const deleteRowHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { row, count = 1 } = args as { row: number; count?: number };

  try {
    const response = await wpsClient.executeMethod<{
      deletedRows: number;
    }>(
      'deleteRows',
      { row, count },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除行失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `已成功删除第${row}行起共${count}行`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除行出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 插入列
 */
export const insertColumnDefinition: ToolDefinition = {
  name: 'wps_excel_insert_column',
  description: '在指定位置插入列。可指定起始列号和插入列数。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      column: {
        type: 'number',
        description: '要插入的列号（从1开始）',
      },
      count: {
        type: 'number',
        description: '要插入的列数，默认1',
      },
    },
    required: ['column'],
  },
};

export const insertColumnHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { column, count = 1 } = args as { column: number; count?: number };

  try {
    const response = await wpsClient.executeMethod<{
      insertedColumns: number;
    }>(
      'insertColumns',
      { column, count },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入列失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `已在第${column}列处成功插入${count}列`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入列出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 删除指定列
 */
export const deleteColumnDefinition: ToolDefinition = {
  name: 'wps_excel_delete_column',
  description: '删除指定列。可指定起始列号和删除列数。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      column: {
        type: 'number',
        description: '要删除的起始列号（从1开始）',
      },
      count: {
        type: 'number',
        description: '要删除的列数，默认1',
      },
    },
    required: ['column'],
  },
};

export const deleteColumnHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { column, count = 1 } = args as { column: number; count?: number };

  try {
    const response = await wpsClient.executeMethod<{
      deletedColumns: number;
    }>(
      'deleteColumns',
      { column, count },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除列失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `已成功删除第${column}列起共${count}列`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除列出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 冻结/取消冻结窗格
 */
export const freezePanesDefinition: ToolDefinition = {
  name: 'wps_excel_freeze_panes',
  description: '冻结/取消冻结窗格。可指定冻结的行和列位置。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      row: {
        type: 'number',
        description: '冻结到第几行（从1开始），不填则不冻结行',
      },
      column: {
        type: 'number',
        description: '冻结到第几列（从1开始），不填则不冻结列',
      },
      freeze: {
        type: 'boolean',
        description: '是否冻结，默认true。设为false则取消冻结',
      },
    },
  },
};

export const freezePanesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { row, column, freeze = true } = args as {
    row?: number;
    column?: number;
    freeze?: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      frozen: boolean;
    }>(
      'freezePanes',
      { row, column, freeze },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `冻结窗格操作失败: ${response.error}` }],
        error: response.error,
      };
    }

    const action = freeze ? '冻结' : '取消冻结';
    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `窗格${action}成功！${row ? `\n冻结行: ${row}` : ''}${column ? `\n冻结列: ${column}` : ''}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `冻结窗格操作出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 自动填充单元格区域
 */
export const autoFillDefinition: ToolDefinition = {
  name: 'wps_excel_auto_fill',
  description: '自动填充单元格区域。根据源区域的数据模式自动填充到目标区域。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sourceRange: {
        type: 'string',
        description: '源数据区域，如 "A1:A5"',
      },
      targetRange: {
        type: 'string',
        description: '目标填充区域，如 "A1:A20"',
      },
    },
    required: ['sourceRange', 'targetRange'],
  },
};

export const autoFillHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { sourceRange, targetRange } = args as {
    sourceRange: string;
    targetRange: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      filled: boolean;
    }>(
      'fillSeries',
      { sourceRange, targetRange },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `自动填充失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `自动填充成功！\n源区域: ${sourceRange}\n目标区域: ${targetRange}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `自动填充出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置命名范围
 */
export const setNamedRangeDefinition: ToolDefinition = {
  name: 'wps_excel_set_named_range',
  description: '设置命名范围。为指定单元格区域创建或更新命名范围。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '命名范围的名称',
      },
      range: {
        type: 'string',
        description: '单元格区域，如 "A1:D10"',
      },
    },
    required: ['name', 'range'],
  },
};

export const setNamedRangeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, range } = args as {
    name: string;
    range: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      name: string;
      range: string;
    }>(
      'createNamedRange',
      { name, range },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置命名范围失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `命名范围设置成功！\n名称: ${name}\n范围: ${range}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置命名范围出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 隐藏/显示列
 */
export const hideColumnDefinition: ToolDefinition = {
  name: 'wps_excel_hide_column',
  description: '隐藏或显示指定列。可指定起始列号、列数和隐藏/显示状态。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      column: {
        type: 'number',
        description: '起始列号（从1开始）',
      },
      count: {
        type: 'number',
        description: '列数，默认1',
      },
      hide: {
        type: 'boolean',
        description: '是否隐藏，true为隐藏，false为显示',
      },
    },
    required: ['column', 'count', 'hide'],
  },
};

export const hideColumnHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { column, count = 1, hide } = args as {
    column: number;
    count: number;
    hide: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      column: number;
      count: number;
      hidden: boolean;
    }>(
      'hideColumns',
      { column, count, hide },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `${hide ? '隐藏' : '显示'}列失败: ${response.error}` }],
        error: response.error,
      };
    }

    const action = hide ? '隐藏' : '显示';
    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `已成功${action}第${column}列起共${count}列`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `隐藏/显示列出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 对指定列/行自动求和
 */
export const autoSumDefinition: ToolDefinition = {
  name: 'wps_excel_auto_sum',
  description: '对指定范围的列或行自动求和，并将结果写入目标单元格。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要求和的数据范围，如 "A1:A10" 或 "B2:F2"',
      },
      targetCell: {
        type: 'string',
        description: '求和结果写入的目标单元格，如 "A11" 或 "G2"',
      },
    },
    required: ['range', 'targetCell'],
  },
};

export const autoSumHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, targetCell } = args as {
    range: string;
    targetCell: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      range: string;
      targetCell: string;
      result: number;
    }>(
      'autoSum', // NOTE: macOS未实现，仅Windows支持
      { range, targetCell },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `自动求和失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `自动求和成功！\n求和范围: ${range}\n结果写入: ${targetCell}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `自动求和出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有工作表管理相关的Tools
 */
export const sheetTools: RegisteredTool[] = [
  { definition: createSheetDefinition, handler: createSheetHandler },
  { definition: deleteSheetDefinition, handler: deleteSheetHandler },
  { definition: renameSheetDefinition, handler: renameSheetHandler },
  { definition: copySheetDefinition, handler: copySheetHandler },
  { definition: getSheetListDefinition, handler: getSheetListHandler },
  { definition: switchSheetDefinition, handler: switchSheetHandler },
  { definition: moveSheetDefinition, handler: moveSheetHandler },
  { definition: getSelectionDefinition, handler: getSelectionHandler },
  { definition: deleteRowDefinition, handler: deleteRowHandler },
  { definition: insertColumnDefinition, handler: insertColumnHandler },
  { definition: deleteColumnDefinition, handler: deleteColumnHandler },
  { definition: freezePanesDefinition, handler: freezePanesHandler },
  { definition: autoFillDefinition, handler: autoFillHandler },
  { definition: setNamedRangeDefinition, handler: setNamedRangeHandler },
  { definition: hideColumnDefinition, handler: hideColumnHandler },
  { definition: autoSumDefinition, handler: autoSumHandler },
];

export default sheetTools;

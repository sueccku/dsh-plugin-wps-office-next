/**
 * Input: 格式化工具参数
 * Output: 格式化操作结果
 * Pos: Excel 基础格式化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel格式化Tools - 单元格格式/边框/合并/列宽/行高
 *
 * 包含：
 * - wps_excel_set_cell_format: 设置单元格格式（字体、颜色、背景等）
 * - wps_excel_set_cell_style: 应用预定义样式到单元格
 * - wps_excel_set_border: 设置单元格边框样式
 * - wps_excel_set_number_format: 设置单元格数字格式
 * - wps_excel_merge_cells: 合并指定范围的单元格
 * - wps_excel_unmerge_cells: 拆分合并的单元格
 * - wps_excel_set_column_width: 设置列宽
 * - wps_excel_set_row_height: 设置行高
 * - wps_excel_hide_row: 隐藏/显示行
 * - wps_excel_set_data_validation: 设置数据验证规则
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

// ============================================================
// 1. wps_excel_set_cell_format - 设置单元格格式
// ============================================================

export const setCellFormatDefinition: ToolDefinition = {
  name: 'wps_excel_set_cell_format',
  description: '设置Excel单元格格式，包括字体、颜色、背景色、粗体、斜体、字号等。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要设置格式的范围，如 A1:C10、B2:D5',
      },
      format: {
        type: 'object',
        description: '格式设置对象，可包含 bold(粗体)、italic(斜体)、fontSize(字号)、fontName(字体名)、fontColor(字体颜色，如#FF0000)、bgColor(背景颜色)、underline(下划线)、strikethrough(删除线)、horizontalAlignment(水平对齐: left/center/right)、verticalAlignment(垂直对齐: top/center/bottom)、wrapText(自动换行)',
        properties: {
          bold: { type: 'boolean', description: '是否粗体' },
          italic: { type: 'boolean', description: '是否斜体' },
          fontSize: { type: 'number', description: '字号大小' },
          fontName: { type: 'string', description: '字体名称，如 微软雅黑、Arial' },
          fontColor: { type: 'string', description: '字体颜色，十六进制如 #FF0000' },
          bgColor: { type: 'string', description: '背景颜色，十六进制如 #FFFF00' },
          underline: { type: 'boolean', description: '是否下划线' },
          strikethrough: { type: 'boolean', description: '是否删除线' },
          horizontalAlignment: { type: 'string', description: '水平对齐方式', enum: ['left', 'center', 'right'] },
          verticalAlignment: { type: 'string', description: '垂直对齐方式', enum: ['top', 'center', 'bottom'] },
          wrapText: { type: 'boolean', description: '是否自动换行' },
        },
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range', 'format'],
  },
};

export const setCellFormatHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, format, sheet } = args as {
    range: string;
    format: Record<string, unknown>;
    sheet?: string;
  };

  try {
    const response = await wpsClient.executeMethod(
      'setCellFormat',
      { range, format, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置单元格格式失败: ${response.error}` }],
        error: response.error,
      };
    }

    const formatDesc = Object.entries(format)
      .map(([key, value]) => `${key}: ${value}`)
      .join(', ');

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `单元格格式设置成功！\n范围: ${range}\n格式: ${formatDesc}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置单元格格式出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 2. wps_excel_set_cell_style - 应用预定义样式
// ============================================================

export const setCellStyleDefinition: ToolDefinition = {
  name: 'wps_excel_set_cell_style',
  description: '应用预定义样式到Excel单元格，如标题、强调、输入、输出等内置样式。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要应用样式的范围，如 A1:C10',
      },
      style: {
        type: 'string',
        description: '预定义样式名称，如 标题、强调、好、差、适中、输入、输出、计算、检查单元格、解释性文本、汇总 等',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range', 'style'],
  },
};

export const setCellStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, style, sheet } = args as {
    range: string;
    style: string;
    sheet?: string;
  };

  try {
    const response = await wpsClient.executeMethod(
      'setCellStyle',
      { range, style, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `应用样式失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `样式应用成功！\n范围: ${range}\n样式: ${style}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `应用样式出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 3. wps_excel_set_border - 设置单元格边框
// ============================================================

export const setBorderDefinition: ToolDefinition = {
  name: 'wps_excel_set_border',
  description: '设置Excel单元格边框样式，支持不同粗细、位置和颜色。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要设置边框的范围，如 A1:C10',
      },
      borderStyle: {
        type: 'string',
        description: '边框线条样式',
        enum: ['thin', 'medium', 'thick', 'double', 'none'],
      },
      position: {
        type: 'string',
        description: '边框位置，默认 all（全部边框）',
        enum: ['all', 'top', 'bottom', 'left', 'right', 'outline'],
      },
      color: {
        type: 'string',
        description: '边框颜色，十六进制如 #000000，默认黑色',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range', 'borderStyle'],
  },
};

export const setBorderHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, borderStyle, position, color, sheet } = args as {
    range: string;
    borderStyle: string;
    position?: string;
    color?: string;
    sheet?: string;
  };

  try {
    const response = await wpsClient.executeMethod(
      'setBorder',
      {
        range,
        borderStyle,
        position: position || 'all',
        color: color || '#000000',
        sheet,
      },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置边框失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `边框设置成功！\n范围: ${range}\n样式: ${borderStyle}\n位置: ${position || 'all'}\n颜色: ${color || '#000000'}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置边框出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 4. wps_excel_set_number_format - 设置数字格式
// ============================================================

export const setNumberFormatDefinition: ToolDefinition = {
  name: 'wps_excel_set_number_format',
  description: `设置Excel单元格的数字格式。

常用格式代码：
- #,##0.00 - 千分位+2位小数
- 0.00% - 百分比
- yyyy-mm-dd - 日期
- ¥#,##0.00 - 人民币
- $#,##0.00 - 美元
- 0.00E+00 - 科学计数法
- @ - 文本格式`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要设置格式的范围，如 A1:C10',
      },
      format: {
        type: 'string',
        description: '数字格式代码，如 #,##0.00、0.00%、yyyy-mm-dd',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range', 'format'],
  },
};

export const setNumberFormatHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, format, sheet } = args as {
    range: string;
    format: string;
    sheet?: string;
  };

  try {
    const response = await wpsClient.executeMethod(
      'setNumberFormat',
      { range, format, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置数字格式失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `数字格式设置成功！\n范围: ${range}\n格式: ${format}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置数字格式出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 5. wps_excel_merge_cells - 合并单元格
// ============================================================

export const mergeCellsDefinition: ToolDefinition = {
  name: 'wps_excel_merge_cells',
  description: '合并Excel指定范围的单元格。合并后保留左上角单元格的值。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要合并的范围，如 A1:C3',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range'],
  },
};

export const mergeCellsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, sheet } = args as {
    range: string;
    sheet?: string;
  };

  try {
    const response = await wpsClient.executeMethod(
      'mergeCells',
      { range, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `合并单元格失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `单元格合并成功！\n范围: ${range}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `合并单元格出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 6. wps_excel_unmerge_cells - 拆分合并的单元格
// ============================================================

export const unmergeCellsDefinition: ToolDefinition = {
  name: 'wps_excel_unmerge_cells',
  description: '拆分Excel中已合并的单元格，恢复为独立的单元格。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要拆分的合并单元格范围，如 A1:C3',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range'],
  },
};

export const unmergeCellsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, sheet } = args as {
    range: string;
    sheet?: string;
  };

  try {
    const response = await wpsClient.executeMethod(
      'unmergeCells',
      { range, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `拆分单元格失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `单元格拆分成功！\n范围: ${range}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `拆分单元格出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 7. wps_excel_set_column_width - 设置列宽
// ============================================================

export const setColumnWidthDefinition: ToolDefinition = {
  name: 'wps_excel_set_column_width',
  description: '设置Excel指定列的列宽。支持单列或连续多列。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      column: {
        type: 'string',
        description: '列标识，如 A（单列）或 A:C（连续多列）',
      },
      width: {
        type: 'number',
        description: '列宽数值（字符宽度单位），如 15、20',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['column', 'width'],
  },
};

export const setColumnWidthHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { column, width, sheet } = args as {
    column: string;
    width: number;
    sheet?: string;
  };

  if (width <= 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '列宽必须大于0' }],
      error: '列宽无效',
    };
  }

  try {
    const response = await wpsClient.executeMethod(
      'setColumnWidth',
      { column, width, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置列宽失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `列宽设置成功！\n列: ${column}\n宽度: ${width}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置列宽出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 8. wps_excel_set_row_height - 设置行高
// ============================================================

export const setRowHeightDefinition: ToolDefinition = {
  name: 'wps_excel_set_row_height',
  description: '设置Excel指定行的行高。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      row: {
        type: 'number',
        description: '行号，从1开始',
      },
      height: {
        type: 'number',
        description: '行高数值（磅为单位），如 20、30',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['row', 'height'],
  },
};

export const setRowHeightHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { row, height, sheet } = args as {
    row: number;
    height: number;
    sheet?: string;
  };

  if (row < 1) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '行号必须大于等于1' }],
      error: '行号无效',
    };
  }

  if (height <= 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '行高必须大于0' }],
      error: '行高无效',
    };
  }

  try {
    const response = await wpsClient.executeMethod(
      'setRowHeight',
      { row, height, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置行高失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `行高设置成功！\n行: ${row}\n高度: ${height}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置行高出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 9. wps_excel_hide_row - 隐藏/显示行
// ============================================================

export const hideRowDefinition: ToolDefinition = {
  name: 'wps_excel_hide_row',
  description: '隐藏或显示Excel指定行。可一次操作连续多行。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      row: {
        type: 'number',
        description: '起始行号，从1开始',
      },
      count: {
        type: 'number',
        description: '连续行数，默认1',
      },
      hide: {
        type: 'boolean',
        description: '是否隐藏，true=隐藏 false=显示，默认true',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['row'],
  },
};

export const hideRowHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { row, count = 1, hide = true, sheet } = args as {
    row: number;
    count?: number;
    hide?: boolean;
    sheet?: string;
  };

  if (row < 1) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '行号必须大于等于1' }],
      error: '行号无效',
    };
  }

  try {
    const response = await wpsClient.executeMethod(
      'hideRows',
      { row, count, hide, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `${hide ? '隐藏' : '显示'}行失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `行${hide ? '隐藏' : '显示'}成功！\n起始行: ${row}\n行数: ${count}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `${hide ? '隐藏' : '显示'}行出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 10. wps_excel_set_data_validation - 设置数据验证规则
// ============================================================

export const setDataValidationDefinition: ToolDefinition = {
  name: 'wps_excel_set_data_validation',
  description: '设置Excel单元格的数据验证规则，如下拉列表、数值范围、日期范围等。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要设置验证的范围，如 A1:A100',
      },
      type: {
        type: 'string',
        description: '验证类型：list(下拉列表)、whole(整数)、decimal(小数)、date(日期)、textLength(文本长度)、custom(自定义)',
      },
      formula: {
        type: 'string',
        description: '验证公式。list类型用逗号分隔值如"选项1,选项2,选项3"；数值类型如"1,100"表示范围；custom类型为Excel公式',
      },
      sheet: {
        type: 'string',
        description: '工作表名称，不填则使用当前活动工作表',
      },
    },
    required: ['range', 'type', 'formula'],
  },
};

export const setDataValidationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, type, formula, sheet } = args as {
    range: string;
    type: string;
    formula: string;
    sheet?: string;
  };

  try {
    const response = await wpsClient.executeMethod(
      'addDataValidation',
      { range, type, formula, sheet },
      WpsAppType.SPREADSHEET
    );

    if (!response.success) {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置数据验证失败: ${response.error}` }],
        error: response.error,
      };
    }

    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `数据验证设置成功！\n范围: ${range}\n类型: ${type}\n规则: ${formula}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置数据验证出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 导出所有格式化相关的Tools
// ============================================================

export const excelFormatTools: RegisteredTool[] = [
  { definition: setCellFormatDefinition, handler: setCellFormatHandler },
  { definition: setCellStyleDefinition, handler: setCellStyleHandler },
  { definition: setBorderDefinition, handler: setBorderHandler },
  { definition: setNumberFormatDefinition, handler: setNumberFormatHandler },
  { definition: mergeCellsDefinition, handler: mergeCellsHandler },
  { definition: unmergeCellsDefinition, handler: unmergeCellsHandler },
  { definition: setColumnWidthDefinition, handler: setColumnWidthHandler },
  { definition: setRowHeightDefinition, handler: setRowHeightHandler },
  { definition: hideRowDefinition, handler: hideRowHandler },
  { definition: setDataValidationDefinition, handler: setDataValidationHandler },
];

export default excelFormatTools;

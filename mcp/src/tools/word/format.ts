/**
 * Input: Word 格式化参数
 * Output: 样式与字体设置结果
 * Pos: Word 格式化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word格式化Tools - 排版格式化模块
 * 处理文档样式、字体、目录等格式化需求
 *
 * 包含：
 * - wps_word_apply_style: 应用样式到选中区域
 * - wps_word_set_font: 设置字体格式
 * - wps_word_generate_toc: 生成目录
 * - wps_word_insert_bookmark: 插入书签
 * - wps_word_set_page_setup: 设置页面布局
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
 * 应用样式到选中区域
 * 快速应用Word内置样式，比如标题1、标题2、正文等
 */
export const applyStyleDefinition: ToolDefinition = {
  name: 'wps_word_apply_style',
  description: `应用Word样式到当前选中区域或指定范围。

支持的常用样式：
- 标题1、标题2、标题3...（或 Heading 1, Heading 2...）
- 正文、正文首行缩进
- 引用、强调
- 列表段落

使用场景：
- "把这段设成标题1"
- "应用正文样式"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      style_name: {
        type: 'string',
        description: '样式名称，如 "标题 1"、"正文"、"Heading 1"',
      },
      range: {
        type: 'object',
        description: '指定范围，不填则应用到当前选中区域',
        properties: {
          start: {
            type: 'number',
            description: '起始位置（字符索引）',
          },
          end: {
            type: 'number',
            description: '结束位置（字符索引）',
          },
        },
      },
    },
    required: ['style_name'],
  },
};

export const applyStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { style_name, range } = args as {
    style_name: string;
    range?: { start: number; end: number };
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      affectedText: string;
    }>(
      'applyStyle',
      { styleName: style_name, range },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `样式应用成功！\n样式: ${style_name}\n影响的文本: ${response.data.affectedText}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `应用样式失败: ${response.error}` }],
        error: response.error,
      };
    }
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

/**
 * 设置字体格式
 */
export const setFontDefinition: ToolDefinition = {
  name: 'wps_word_set_font',
  description: `设置字体格式，包括字体名称、字号、加粗、斜体、颜色等。

使用场景：
- "把标题改成微软雅黑24号加粗"
- "把这段文字改成红色"
- "全文字体改成宋体小四"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      font_name: {
        type: 'string',
        description: '字体名称，如 "微软雅黑"、"宋体"、"Arial"',
      },
      font_size: {
        type: 'number',
        description: '字号，如 12、14、24',
      },
      bold: {
        type: 'boolean',
        description: '是否加粗',
      },
      italic: {
        type: 'boolean',
        description: '是否斜体',
      },
      underline: {
        type: 'boolean',
        description: '是否下划线',
      },
      color: {
        type: 'string',
        description: '字体颜色，支持颜色名称(red/blue/green)或十六进制(#FF0000)',
      },
      range: {
        type: 'string',
        description: '应用范围，可选值: "selection"(当前选中), "all"(全文)。默认selection',
        enum: ['selection', 'all'],
      },
    },
    required: [],
  },
};

export const setFontHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { font_name, font_size, bold, italic, underline, color, range } = args as {
    font_name?: string;
    font_size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    range?: string;
  };

  // 至少要设置一个属性吧
  if (!font_name && !font_size && bold === undefined && italic === undefined &&
      underline === undefined && !color) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '请至少指定一个字体属性（如 font_name、font_size、bold 等）' }],
      error: '没有指定任何字体属性',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      settings: Record<string, unknown>;
    }>(
      'setFont',
      {
        fontName: font_name,
        fontSize: font_size,
        bold,
        italic,
        underline,
        color,
        range: range || 'selection',
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const settings = response.data.settings;
      let settingStr = '';
      if (settings.fontName) settingStr += `字体: ${settings.fontName}\n`;
      if (settings.fontSize) settingStr += `字号: ${settings.fontSize}\n`;
      if (settings.bold !== undefined) settingStr += `加粗: ${settings.bold ? '是' : '否'}\n`;
      if (settings.italic !== undefined) settingStr += `斜体: ${settings.italic ? '是' : '否'}\n`;
      if (settings.color) settingStr += `颜色: ${settings.color}\n`;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `字体格式设置成功！\n${settingStr}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置字体失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置字体出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 生成目录
 * 自动根据文档中的标题样式生成目录
 */
export const generateTocDefinition: ToolDefinition = {
  name: 'wps_word_generate_toc',
  description: `根据文档中的标题样式自动生成目录。

前提条件：文档中的标题必须使用"标题1"、"标题2"等样式。

使用场景：
- "帮我生成目录"
- "在文档开头插入目录"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      position: {
        type: 'string',
        description: '插入位置，可选值: "start"(文档开头), "cursor"(当前光标位置)。默认start',
        enum: ['start', 'cursor'],
      },
      levels: {
        type: 'number',
        description: '目录包含的标题级别数，如 3 表示包含标题1-3。默认3',
      },
      include_page_numbers: {
        type: 'boolean',
        description: '是否包含页码，默认true',
      },
    },
    required: [],
  },
};

export const generateTocHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { position, levels, include_page_numbers } = args as {
    position?: string;
    levels?: number;
    include_page_numbers?: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      levels: number;
    }>(
      'generateTOC',
      {
        position: position || 'start',
        levels: levels || 3,
        includePageNumbers: include_page_numbers !== false,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `目录生成成功！\n包含标题级别: 1-${response.data.levels}\n位置: ${position === 'cursor' ? '当前光标位置' : '文档开头'}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `生成目录失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `生成目录出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 插入书签
 * 在当前光标位置插入书签，方便交叉引用和导航
 */
export const insertBookmarkDefinition: ToolDefinition = {
  name: 'wps_word_insert_bookmark',
  description: `在当前光标位置或选中区域插入书签。

书签可用于：
- 交叉引用
- 超链接跳转目标
- 文档内快速导航

使用场景：
- "在这里插入一个书签"
- "标记这个位置为'章节开头'"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '书签名称，不能包含空格，建议使用英文或下划线连接',
      },
    },
    required: ['name'],
  },
};

export const insertBookmarkHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertBookmark',
      { name },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `书签插入成功！\n书签名称: ${name}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入书签失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入书签出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置页面布局
 * 调整页面方向、边距等页面设置
 */
export const setPageSetupDefinition: ToolDefinition = {
  name: 'wps_word_set_page_setup',
  description: `设置文档页面布局，包括页面方向和边距。

使用场景：
- "把页面改成横向"
- "设置上下边距为2厘米"
- "调整页面为A4横向，边距2cm"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      orientation: {
        type: 'string',
        description: '页面方向: "portrait"(纵向) 或 "landscape"(横向)',
        enum: ['portrait', 'landscape'],
      },
      marginTop: {
        type: 'number',
        description: '上边距（磅值）',
      },
      marginBottom: {
        type: 'number',
        description: '下边距（磅值）',
      },
      marginLeft: {
        type: 'number',
        description: '左边距（磅值）',
      },
      marginRight: {
        type: 'number',
        description: '右边距（磅值）',
      },
    },
    required: [],
  },
};

export const setPageSetupHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { orientation, marginTop, marginBottom, marginLeft, marginRight } = args as {
    orientation?: string;
    marginTop?: number;
    marginBottom?: number;
    marginLeft?: number;
    marginRight?: number;
  };

  if (!orientation && marginTop === undefined && marginBottom === undefined &&
      marginLeft === undefined && marginRight === undefined) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '请至少指定一个页面设置属性（如 orientation、marginTop 等）' }],
      error: '没有指定任何页面设置属性',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      settings: Record<string, unknown>;
    }>(
      'setPageSetup',
      { orientation, marginTop, marginBottom, marginLeft, marginRight },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const s = response.data.settings;
      let desc = '';
      if (s.orientation) desc += `页面方向: ${s.orientation === 'landscape' ? '横向' : '纵向'}\n`;
      if (s.marginTop !== undefined) desc += `上边距: ${s.marginTop}pt\n`;
      if (s.marginBottom !== undefined) desc += `下边距: ${s.marginBottom}pt\n`;
      if (s.marginLeft !== undefined) desc += `左边距: ${s.marginLeft}pt\n`;
      if (s.marginRight !== undefined) desc += `右边距: ${s.marginRight}pt\n`;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `页面布局设置成功！\n${desc}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页面布局失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页面布局出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有格式化相关的Tools
 */
export const formatTools: RegisteredTool[] = [
  { definition: applyStyleDefinition, handler: applyStyleHandler },
  { definition: setFontDefinition, handler: setFontHandler },
  { definition: generateTocDefinition, handler: generateTocHandler },
  { definition: insertBookmarkDefinition, handler: insertBookmarkHandler },
  { definition: setPageSetupDefinition, handler: setPageSetupHandler },
];

export default formatTools;

/**
 * Input: Word 格式化参数
 * Output: 样式与字体设置结果
 * Pos: Word 格式化工具实现。一旦我被修改，请更新我的头部注释，并在 docs/FIXES.md 记账。
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
import { resolveStyleName } from './style-names';
import { WpsAppType } from '../../types/wps';

/**
 * 应用样式到选中区域
 * 快速应用Word内置样式，比如标题1、标题2、正文等
 */
export const applyStyleDefinition: ToolDefinition = {
  name: 'wps_word_apply_style',
  description: `应用Word样式到当前选中区域或指定范围。

支持的常用样式（**用中文名最稳**）：
- 标题 1 … 标题 9（注意中间有空格）、标题、副标题
- 正文、正文文本、正文首行缩进
- 引用、明显引用、强调、明显强调、题注、列表段落

英文别名会自动翻译：Heading 1 → 标题 1、Normal → 正文、Title → 标题、Subtitle → 副标题、Quote → 引用；
「标题1」（少空格）也会自动补成「标题 1」。**中文版 WPS 里没有英文样式名**，传别的英文名不会生效。

使用场景：
- "把这段设成标题1"
- "应用正文样式"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      styleName: {
        type: 'string',
        description:
          '样式名称，如「标题 1」（有空格）、「正文」、「副标题」。英文别名可用（Heading 1 / Normal / Title / Subtitle / Quote），会自动翻译成中文内置名。',
      },
      range: {
        type: ['object', 'string'],
        description:
          '指定范围（0 基字符偏移，end 不含），不填则应用到当前选中区域。也可以直接传字符串 "all" 表示整篇文档' +
          '（与 wps_word_set_font 的 range 语义一致）。注意：段落样式会作用于与范围相交的**整段**——想改第 N 段，' +
          '就用 wps_word_get_paragraphs 里那一段的 start/end。',
        properties: {
          start: {
            type: 'number',
            description: '起始位置（字符索引，0 基）',
          },
          end: {
            type: 'number',
            description: '结束位置（字符索引，不含）',
          },
        },
        required: ['start', 'end'],
      },
    },
    required: ['styleName'],
  },
};

export const applyStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { styleName, range } = args as {
    styleName: string;
    range?: { start: number; end: number } | 'all';
  };

  // FIXES 88（L2）：中文 WPS 的样式表里没有英文名（实测 NameInternational 全空、Item 抛错），
  // 而描述里曾把 Heading 1 当示例。这里把英文别名与「少空格」写法翻译成 WPS 认的名字再发出去。
  const resolved = resolveStyleName(styleName);

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      affectedText: string;
      range?: { start: number; end: number };
      affectedParagraphs?: number[];
      affectedParagraphCount?: number;
    }>(
      'applyStyle',
      { styleName: resolved.name, range },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const d = response.data;
      const paragraphs = Array.isArray(d.affectedParagraphs) ? d.affectedParagraphs : [];
      const lines = [
        '样式应用成功！',
        `样式: ${styleName}`,
        // affectedText 现在是赋值**之前**的快照（以前是赋值之后读的，范围已被 Word 扩张，所以那个值是错的）
        `影响的文本: ${d.affectedText}`,
      ];
      if (paragraphs.length) {
        lines.push(`影响的段落: 第 ${paragraphs.join('、')} 段`);
      }
      if (typeof d.range?.start === 'number' && typeof d.range?.end === 'number') {
        lines.push(`实际作用范围: ${d.range.start}-${d.range.end}`);
      }
      if (paragraphs.length > 1) {
        lines.push(
          `注意：段落样式作用于与范围相交的整段，本次连带影响了 ${paragraphs.length} 个段落。范围只能决定从哪一段开始。`
        );
      }
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    } else {
      // FIXES 88（L2）：样式名不存在时把「实际发送的名字」与可执行的下一步说清楚，
      // 否则模型只会照着 schema 的示例继续试错（E_FAIL 本身不告诉它问题在哪）。
      const hint = resolved.translated
        ? `（已把「${styleName}」翻译为「${resolved.name}」）下一步：用 wps_word_get_paragraphs 看现有段落的样式名，或换一个文档里确实存在的样式。`
        : '下一步：样式名必须与文档里的完全一致（中文名中间有空格，如「标题 1」）；可用 wps_word_get_paragraphs 看当前段落在用的样式名。';
      const errText = `应用样式失败: ${response.error}\n请求的样式: 「${styleName}」→ 实际发送: 「${resolved.name}」\n${hint}`;
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: errText }],
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
      fontName: {
        type: 'string',
        description: '字体名称，如 "微软雅黑"、"宋体"、"Arial"',
      },
      fontSize: {
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
        // D1 实测：这里原来只声明 object，而描述里写"传 \"all\" 表示全文"——
        // 参数形状守卫按声明拒绝字符串，于是 Skill 与描述推荐的 `set_font {range:"all"}`
        // 直接报「参数 range 的形状不对」。桥（FIXES 86）本来就接受 "all"，所以是声明漏了联合类型。
        type: ['object', 'string'],
        description:
          '作用范围（0 基字符偏移，end 不含）。用 wps_word_get_paragraphs 拿每段的 start/end 再传进来，' +
          '例如 {start: 12, end: 40} 只给这一段加粗。传 "all" 表示全文；不传表示当前选中内容（没有选中时是空范围，什么都不会变）',
        properties: {
          start: { type: 'number', description: '起始字符位置（0 基）' },
          end: { type: 'number', description: '结束字符位置（不含）' },
        },
        required: ['start', 'end'],
      },
    },
    required: [],
  },
};

export const setFontHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { fontName, fontSize, bold, italic, underline, color, range } = args as {
    fontName?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    range?: { start: number; end: number } | 'all';
  };

  // 至少要设置一个属性吧
  if (!fontName && !fontSize && bold === undefined && italic === undefined &&
      underline === undefined && !color) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '请至少指定一个字体属性（如 fontName、fontSize、bold 等）' }],
      error: '没有指定任何字体属性',
    };
  }

  try {
    // 范围显式解析：对象直接透传（桥会做边界校验），"all" 表示全文，未指定则由桥用当前选区。
    // 旧版这里恒发 range: range || 'selection'，而桥是 if ... -eq "all" else 选区 —— "selection" 永远落进 else，
    // 于是"设了但没生效"（FIXES 86 的 B1）。现在没有任何隐式兜底。
    const resolvedRange =
      range === 'all' ? 'all' : range && typeof range === 'object' ? range : undefined;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      settings: Record<string, unknown>;
      range?: { start: number; end: number; resolve: string; characters: number };
      readBack?: Record<string, unknown>;
    }>(
      'setFont',
      {
        fontName: fontName,
        fontSize: fontSize,
        bold,
        italic,
        underline,
        color,
        range: resolvedRange,
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const settings = response.data.settings;
      const info = response.data.range;
      let settingStr = '';
      if (settings.fontName) settingStr += `字体: ${settings.fontName}\n`;
      if (settings.fontSize) settingStr += `字号: ${settings.fontSize}\n`;
      if (settings.bold !== undefined) settingStr += `加粗: ${settings.bold ? '是' : '否'}\n`;
      if (settings.italic !== undefined) settingStr += `斜体: ${settings.italic ? '是' : '否'}\n`;
      if (settings.color) settingStr += `颜色: ${settings.color}\n`;
      // 如实回报**实际作用范围**：范围为空时上面那些"是"其实什么也没改（桥会同时发 warning）。
      const scopeStr = info
        ? `作用范围: ${info.start}-${info.end}（${info.resolve === 'characterRange' ? '指定字符范围' : info.resolve === 'all' ? '全文' : '当前选区'}），${info.characters} 个字符\n`
        : '';
      const emptyHint =
        info && info.characters === 0
          ? '\n注意：作用范围里没有任何字符，格式没有实际落点。请用 range:{start,end} 指定文字范围（先用 wps_word_get_paragraphs 取坐标），或先在文档里选中内容。'
          : '';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `字体格式设置成功！\n${settingStr}${scopeStr}${emptyHint}`,
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
      includePageNumbers: {
        type: 'boolean',
        description: '是否包含页码，默认true',
      },
      action: {
        type: 'string',
        description: 'insert（默认，新插一个目录）或 update（只刷新已有目录，不新增）。没有目录时 update 会如实报错并提示改用 insert',
        enum: ['insert', 'update'],
      },
    },
    required: [],
  },
};

export const generateTocHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { position, levels, includePageNumbers, action } = args as {
    position?: string;
    levels?: number;
    includePageNumbers?: boolean;
    action?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      levels: number;
      action?: string;
      tablesOfContents?: number;
      updated?: number;
      entries?: number;
    }>(
      'generateTOC',
      {
        position: position || 'start',
        levels: levels || 3,
        includePageNumbers: includePageNumbers !== false,
        action: action || 'insert',
      },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const d = response.data;
      // update 模式不新插目录，所以不要报"插到了哪里"—— 如实回报刷新了几份、共几条条目。
      if ((d.action || 'insert') === 'update') {
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `目录已刷新！\n刷新的目录数: ${d.updated ?? 0}\n条目数: ${d.entries ?? 0}`,
            },
          ],
        };
      }
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `目录生成成功！\n包含标题级别: 1-${d.levels}\n位置: ${position === 'cursor' ? '当前光标位置' : '文档开头'}`,
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
  description: `设置**某一节**的页面布局：页面方向、纸张大小、页边距。

使用场景：
- "把页面改成横向"
- "设置上下边距为2厘米"
- "调整页面为A4横向，边距2cm"
- "第3节改成横向A3，前面的节不动"（多节混合版式的标书/施工方案）

注意：
- 只影响**点名的那一节**（不填 section 就是第 1 节）。要"整篇都改"得逐节调用，或者先看有几节。
- **横向必须和纸张一起给**（orientation + paperSize）。只给 orientation 时 WPS 会把纸张变成"自定义"，
  不再是 A4/A3；桥内部会先设纸张、再设方向——顺序反了方向会被静默吃回去（S7 实测）。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      section: {
        type: 'number',
        description: '第几节（从 1 开始）；不填则用第 1 节。节数用 wps_word_get_header_footer 或 get_page_setup 逐节探',
      },
      orientation: {
        type: 'string',
        description: '页面方向: "portrait"(纵向) 或 "landscape"(横向)；建议与 paperSize 同时给',
        enum: ['portrait', 'landscape'],
      },
      paperSize: {
        type: 'string',
        description: '纸张大小：A4 / A3 / Letter / Legal',
        enum: ['A4', 'A3', 'Letter', 'Legal'],
      },
      topMargin: {
        type: 'number',
        description: '上边距（磅值）',
      },
      bottomMargin: {
        type: 'number',
        description: '下边距（磅值）',
      },
      leftMargin: {
        type: 'number',
        description: '左边距（磅值）',
      },
      rightMargin: {
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
  const { section, orientation, paperSize, topMargin, bottomMargin, leftMargin, rightMargin } = args as {
    section?: number;
    orientation?: string;
    paperSize?: string;
    topMargin?: number;
    bottomMargin?: number;
    leftMargin?: number;
    rightMargin?: number;
  };

  // 只给 section 也算合法调用：那是"读这一节现在的页面设置"（顺手拿到全文节数），
  // 多节文档里探节数就靠它——get_page_setup 只报自己那一节、不报总数。
  if (!orientation && !paperSize && topMargin === undefined && bottomMargin === undefined &&
      leftMargin === undefined && rightMargin === undefined && section === undefined) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '请至少指定一个页面设置属性（如 orientation、paperSize、topMargin 等），或者给一个 section 来读该节现状' }],
      error: '没有指定任何页面设置属性',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      section?: number;
      sectionCount?: number;
      settings: Record<string, unknown>;
    }>(
      'setPageSetup',
      // The bridge's keys are topMargin/bottomMargin/leftMargin/rightMargin; the schema keeps the
      // marginX spelling for callers.
      { section, orientation, paperSize, topMargin: topMargin, bottomMargin: bottomMargin, leftMargin: leftMargin, rightMargin: rightMargin },
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const s = response.data.settings;
      let desc = '';
      desc += `第 ${response.data.section ?? 1} 节（全文共 ${response.data.sectionCount ?? '?'} 节）\n`;
      if (s.orientation) desc += `页面方向: ${s.orientation === 'landscape' ? '横向' : '纵向'}\n`;
      if (s.paperSize !== undefined) desc += `纸张: ${s.paperSize}\n`;
      if (s.pageWidth !== undefined && s.pageHeight !== undefined) desc += `页面: ${s.pageWidth}pt x ${s.pageHeight}pt\n`;
      if (s.topMargin !== undefined) desc += `上边距: ${s.topMargin}pt\n`;
      if (s.bottomMargin !== undefined) desc += `下边距: ${s.bottomMargin}pt\n`;
      if (s.leftMargin !== undefined) desc += `左边距: ${s.leftMargin}pt\n`;
      if (s.rightMargin !== undefined) desc += `右边距: ${s.rightMargin}pt\n`;

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
 * 读取页面布局
 * 与 set_page_setup 对称：改之前先看现状，改完再核对生效了没有。
 */
export const getPageSetupDefinition: ToolDefinition = {
  name: 'wps_word_get_page_setup',
  description: `读取某一节的页面设置：纸张方向、纸张大小、页边距（磅）与页面宽高。使用场景：
- "现在页边距是多少"
- 改页面设置之前先看一眼，改完再核对生效了没有`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      section: {
        type: 'number',
        description: '第几节（从 1 开始）；不填则用第 1 节',
      },
    },
  },
};

export const getPageSetupHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { section } = args as { section?: number };
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      section?: number;
      orientation?: string;
      paperSize?: string | number;
      topMargin?: number;
      bottomMargin?: number;
      leftMargin?: number;
      rightMargin?: number;
      pageWidth?: number;
      pageHeight?: number;
    }>('getPageSetup', { section }, WpsAppType.WRITER);

    if (response.success && response.data) {
      const d = response.data;
      const parts = [
        `第 ${d.section} 节页面设置：`,
        `方向: ${d.orientation === 'landscape' ? '横向' : '纵向'}`,
        `纸张: ${d.paperSize ?? '未知'}`,
        `页面: ${d.pageWidth}pt x ${d.pageHeight}pt`,
        `边距: 上 ${d.topMargin}pt / 下 ${d.bottomMargin}pt / 左 ${d.leftMargin}pt / 右 ${d.rightMargin}pt`,
      ];
      return { id: uuidv4(), success: true, content: [{ type: 'text', text: parts.join('\n') }] };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `读取页面设置失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `读取页面设置出错: ${errMsg}` }],
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
  { definition: setPageSetupDefinition, handler: setPageSetupHandler },
  { definition: getPageSetupDefinition, handler: getPageSetupHandler },
  { definition: insertBookmarkDefinition, handler: insertBookmarkHandler },
];

export default formatTools;

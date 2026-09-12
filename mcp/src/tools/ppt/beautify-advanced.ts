/**
 * Input: PPT 高级美化操作参数
 * Output: 高级美化操作结果
 * Pos: PPT 高级美化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT高级美化Tools - 高级美化模块
 * 处理配色方案、自动美化、KPI卡片、装饰元素等高级美化操作
 *
 * 包含：
 * - wps_ppt_apply_color_scheme: 应用配色方案
 * - wps_ppt_auto_beautify_slide: 自动美化单页幻灯片
 * - wps_ppt_beautify_all_slides: 批量美化所有幻灯片
 * - wps_ppt_create_kpi_cards: 创建KPI指标卡片
 * - wps_ppt_create_styled_table: 创建带样式的表格
 * - wps_ppt_add_title_decoration: 添加标题装饰元素
 * - wps_ppt_add_page_indicator: 添加页码指示器
 * - wps_ppt_set_background_gradient: 设置渐变背景
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
 * 应用配色方案
 * 为幻灯片或全文档应用统一配色方案
 */
export const applyColorSchemeDefinition: ToolDefinition = {
  name: 'wps_ppt_apply_color_scheme',
  description: `为幻灯片应用统一配色方案。

支持的配色方案：
- ocean: 海洋蓝（#0078D4主色）
- forest: 森林绿（#107C10主色）
- sunset: 日落橙（#FF8C00主色）
- lavender: 薰衣草紫（#7B68EE主色）
- corporate: 企业蓝灰（#2B579A主色）
- warm: 暖色调（#E74C3C主色）
- cool: 冷色调（#3498DB主色）
- monochrome: 单色黑白灰

使用场景：
- "给PPT换个海洋蓝配色"
- "用企业风格的配色方案"
- "统一所有页面的颜色"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始），不填则应用到所有页',
      },
      scheme: {
        type: 'string',
        description: '配色方案名称',
      },
    },
    required: ['scheme'],
  },
};

export const applyColorSchemeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, scheme } = args as {
    slideIndex?: number;
    scheme: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      appliedSlides: number;
    }>(
      'applyColorScheme',
      { slideIndex, scheme },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const schemeNames: Record<string, string> = {
        ocean: '海洋蓝',
        forest: '森林绿',
        sunset: '日落橙',
        lavender: '薰衣草紫',
        corporate: '企业蓝灰',
        warm: '暖色调',
        cool: '冷色调',
        monochrome: '单色黑白灰',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `配色方案应用成功！\n方案: ${schemeNames[scheme] || scheme}\n范围: ${slideIndex ? `第 ${slideIndex} 页` : '所有幻灯片'}\n已应用页数: ${response.data.appliedSlides}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `应用配色方案失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `应用配色方案出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 自动美化单页幻灯片
 * 智能分析并优化单页幻灯片的排版、字体、间距和配色
 */
export const autoBeautifySlideDefinition: ToolDefinition = {
  name: 'wps_ppt_auto_beautify_slide',
  description: `自动美化指定的单页幻灯片。

智能分析幻灯片内容并自动执行以下优化：
- 统一字体风格
- 优化元素间距和对齐
- 调整文本大小和行距
- 应用协调的配色
- 调整形状和图片布局

使用场景：
- "自动美化第3页"
- "优化这页PPT的排版"
- "一键美化当前幻灯片"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const autoBeautifySlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as {
    slideIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      operations: Array<{
        operation: string;
        count: number;
        details?: string;
      }>;
    }>(
      'autoBeautifySlide',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      let output = `幻灯片自动美化完成！\n幻灯片: 第 ${slideIndex} 页\n\n优化详情：\n`;

      const opNames: Record<string, string> = {
        unify_font: '统一字体',
        optimize_spacing: '优化间距',
        align_elements: '对齐元素',
        adjust_colors: '调整配色',
        fix_layout: '修正布局',
        adjust_text_size: '调整文字大小',
      };

      response.data.operations.forEach((op) => {
        output += `- ${opNames[op.operation] || op.operation}: 处理了 ${op.count} 个元素${op.details ? ` (${op.details})` : ''}\n`;
      });

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `自动美化失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `自动美化出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 批量美化所有幻灯片
 * 对演示文稿中所有幻灯片应用统一美化风格
 */
export const beautifyAllSlidesDefinition: ToolDefinition = {
  name: 'wps_ppt_beautify_all_slides',
  description: `批量美化演示文稿中的所有幻灯片。

支持的风格：
- business: 商务风（稳重、专业）
- tech: 科技风（现代、简约）
- creative: 创意风（活泼、多彩）
- minimal: 极简风（留白、克制）
- academic: 学术风（严谨、规范）

使用场景：
- "把整个PPT都美化一下"
- "用科技风格统一美化所有页面"
- "一键优化整个演示文稿"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      style: {
        type: 'string',
        description: '美化风格，不填则默认使用 business 风格',
      },
    },
    required: [],
  },
};

export const beautifyAllSlidesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { style } = args as {
    style?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideCount: number;
      style: string;
    }>(
      'beautifyAllSlides',
      { style: style || 'business' },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const styleNames: Record<string, string> = {
        business: '商务风',
        tech: '科技风',
        creative: '创意风',
        minimal: '极简风',
        academic: '学术风',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `全部幻灯片美化完成！\n风格: ${styleNames[response.data.style] || response.data.style}\n处理页数: ${response.data.slideCount} 页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `批量美化失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `批量美化出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 创建KPI指标卡片
 * 在幻灯片上生成数据指标卡片，适合展示关键业绩数据
 */
export const createKpiCardsDefinition: ToolDefinition = {
  name: 'wps_ppt_create_kpi_cards',
  description: `在幻灯片上创建KPI指标卡片。

每个KPI卡片数据项包含：
- title: 指标名称（如"销售额"、"增长率"）
- value: 指标数值（如"¥128万"、"+15.3%"）
- trend: 趋势方向（up/down/flat），可选
- color: 卡片颜色（十六进制），可选

使用场景：
- "创建一组KPI展示卡片"
- "在第2页展示销售数据指标"
- "添加业绩看板"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      data: {
        type: 'array',
        description: 'KPI数据数组，每项包含title（指标名称）、value（指标数值）、trend（趋势:up/down/flat，可选）、color（颜色，可选）',
        items: {
          type: 'object',
          properties: {
            title: { type: 'string', description: '指标名称' },
            value: { type: 'string', description: '指标数值' },
            trend: { type: 'string', description: '趋势方向', enum: ['up', 'down', 'flat'] },
            color: { type: 'string', description: '卡片颜色（十六进制）' },
          },
          required: ['title', 'value'],
        },
      },
    },
    required: ['slideIndex', 'data'],
  },
};

export const createKpiCardsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, data } = args as {
    slideIndex: number;
    data: Array<{
      title: string;
      value: string;
      trend?: string;
      color?: string;
    }>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      cardCount: number;
    }>(
      'createKpiCards',
      { slideIndex, data },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      let output = `KPI卡片创建成功！\n幻灯片: 第 ${slideIndex} 页\n卡片数量: ${response.data.cardCount} 个\n\n指标列表：\n`;
      data.forEach((item) => {
        const trendIcon = item.trend === 'up' ? '↑' : item.trend === 'down' ? '↓' : '→';
        output += `- ${item.title}: ${item.value} ${item.trend ? trendIcon : ''}\n`;
      });

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建KPI卡片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建KPI卡片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 创建带样式的表格
 * 插入预设样式的精美表格
 */
export const createStyledTableDefinition: ToolDefinition = {
  name: 'wps_ppt_create_styled_table',
  description: `在幻灯片中创建带预设样式的精美表格。

表格数据格式：二维数组，第一行为表头。
例如: [["姓名","分数"],["张三","95"],["李四","88"]]

支持的样式：
- blue: 蓝色主题（默认）
- green: 绿色主题
- orange: 橙色主题
- gray: 灰色主题
- dark: 深色主题
- striped: 斑马纹样式

使用场景：
- "创建一个蓝色主题的数据表格"
- "插入带样式的表格"
- "在第3页添加成绩表"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      data: {
        type: 'array',
        description: '表格数据，二维数组格式，第一行为表头',
        items: {
          type: 'array',
          items: { type: 'string' },
        },
      },
      style: {
        type: 'string',
        description: '表格样式，不填默认使用 blue',
      },
    },
    required: ['slideIndex', 'data'],
  },
};

export const createStyledTableHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, data, style } = args as {
    slideIndex: number;
    data: string[][];
    style?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      rows: number;
      cols: number;
    }>(
      'createStyledTable',
      { slideIndex, data, style: style || 'blue' },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const styleNames: Record<string, string> = {
        blue: '蓝色主题',
        green: '绿色主题',
        orange: '橙色主题',
        gray: '灰色主题',
        dark: '深色主题',
        striped: '斑马纹',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `带样式表格创建成功！\n幻灯片: 第 ${slideIndex} 页\n规格: ${response.data.rows} 行 × ${response.data.cols} 列\n样式: ${styleNames[style || 'blue'] || style}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建带样式表格失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建带样式表格出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 添加标题装饰元素
 * 为幻灯片标题添加装饰性元素（下划线、色块、图标等）
 */
export const addTitleDecorationDefinition: ToolDefinition = {
  name: 'wps_ppt_add_title_decoration',
  description: `为幻灯片标题添加装饰性元素。

支持的装饰样式：
- underline: 标题下划线装饰
- block: 色块背景装饰
- sidebar: 侧边竖条装饰
- bracket: 括号框装饰
- gradient_bar: 渐变条装饰

使用场景：
- "给标题加个装饰线"
- "美化标题样式"
- "添加侧边装饰条"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      style: {
        type: 'string',
        description: '装饰样式，不填默认使用 underline',
      },
    },
    required: ['slideIndex'],
  },
};

export const addTitleDecorationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, style } = args as {
    slideIndex: number;
    style?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'addTitleDecoration',
      { slideIndex, style: style || 'underline' },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const styleNames: Record<string, string> = {
        underline: '下划线装饰',
        block: '色块背景装饰',
        sidebar: '侧边竖条装饰',
        bracket: '括号框装饰',
        gradient_bar: '渐变条装饰',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `标题装饰添加成功！\n幻灯片: 第 ${slideIndex} 页\n装饰样式: ${styleNames[style || 'underline'] || style}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加标题装饰失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加标题装饰出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 添加页码指示器
 * 在幻灯片上添加页码/进度指示器
 */
export const addPageIndicatorDefinition: ToolDefinition = {
  name: 'wps_ppt_add_page_indicator',
  description: `为演示文稿添加页码指示器。

支持的位置：
- bottom-right: 右下角（默认）
- bottom-center: 底部居中
- bottom-left: 左下角

使用场景：
- "给PPT加上页码"
- "在右下角添加页码指示"
- "添加页面进度指示器"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      position: {
        type: 'string',
        description: '页码位置',
        enum: ['bottom-right', 'bottom-center', 'bottom-left'],
      },
    },
    required: [],
  },
};

export const addPageIndicatorHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { position } = args as {
    position?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideCount: number;
    }>(
      'addPageIndicator',
      { position: position || 'bottom-right' },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const posNames: Record<string, string> = {
        'bottom-right': '右下角',
        'bottom-center': '底部居中',
        'bottom-left': '左下角',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `页码指示器添加成功！\n位置: ${posNames[position || 'bottom-right'] || position}\n已添加页数: ${response.data.slideCount} 页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加页码指示器失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加页码指示器出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置渐变背景
 * 为幻灯片设置渐变色背景
 */
export const setBackgroundGradientDefinition: ToolDefinition = {
  name: 'wps_ppt_set_background_gradient',
  description: `为幻灯片设置渐变色背景。

渐变配置对象（gradient）属性：
- type: 渐变类型（linear/radial），默认linear
- angle: 渐变角度（0-360度），仅linear有效，默认180
- colors: 渐变颜色数组（如 ["#1a1a2e", "#16213e", "#0f3460"]）
- stops: 颜色停靠点数组（如 [0, 0.5, 1]），与colors对应

使用场景：
- "设置蓝色渐变背景"
- "给第2页加个从深到浅的渐变"
- "设置PPT背景为渐变色"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      gradient: {
        type: 'object',
        description: '渐变配置对象，包含type（linear/radial）、angle（角度）、colors（颜色数组）、stops（停靠点数组）',
      },
    },
    required: ['slideIndex', 'gradient'],
  },
};

export const setBackgroundGradientHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, gradient } = args as {
    slideIndex: number;
    gradient: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setBackgroundGradient',
      { slideIndex, gradient },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const gradientType = gradient.type === 'radial' ? '径向渐变' : '线性渐变';
      const colors = Array.isArray(gradient.colors) ? (gradient.colors as string[]).join(' → ') : '';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `渐变背景设置成功！\n幻灯片: 第 ${slideIndex} 页\n渐变类型: ${gradientType}${gradient.angle ? `\n渐变角度: ${gradient.angle}°` : ''}${colors ? `\n颜色: ${colors}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置渐变背景失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置渐变背景出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有高级美化相关的Tools
 */
export const beautifyAdvancedTools: RegisteredTool[] = [
  { definition: applyColorSchemeDefinition, handler: applyColorSchemeHandler },
  { definition: autoBeautifySlideDefinition, handler: autoBeautifySlideHandler },
  { definition: beautifyAllSlidesDefinition, handler: beautifyAllSlidesHandler },
  { definition: createKpiCardsDefinition, handler: createKpiCardsHandler },
  { definition: createStyledTableDefinition, handler: createStyledTableHandler },
  { definition: addTitleDecorationDefinition, handler: addTitleDecorationHandler },
  { definition: addPageIndicatorDefinition, handler: addPageIndicatorHandler },
  { definition: setBackgroundGradientDefinition, handler: setBackgroundGradientHandler },
];

export default beautifyAdvancedTools;

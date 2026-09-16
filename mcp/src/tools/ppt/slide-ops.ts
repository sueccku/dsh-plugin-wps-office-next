/**
 * Input: 幻灯片操作工具参数
 * Output: 幻灯片操作结果
 * Pos: PPT 幻灯片操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 幻灯片操作Tools - 删除/复制/移动/查询/切换/布局/备注
 *
 * 包含：
 * - wps_ppt_delete_slide: 删除指定幻灯片
 * - wps_ppt_duplicate_slide: 复制幻灯片
 * - wps_ppt_move_slide: 移动幻灯片到指定位置
 * - wps_ppt_get_slide_count: 获取幻灯片总数
 * - wps_ppt_get_slide_info: 获取指定幻灯片的详细信息
 * - wps_ppt_switch_slide: 切换到指定幻灯片
 * - wps_ppt_set_slide_layout: 设置幻灯片版式布局
 * - wps_ppt_get_slide_notes: 获取幻灯片备注内容
 * - wps_ppt_set_slide_notes: 设置幻灯片备注
 * - wps_ppt_add_shape: 添加形状
 * - wps_ppt_set_shape_style: 设置形状样式
 * - wps_ppt_add_textbox: 添加文本框
 * - wps_ppt_set_slide_title: 设置幻灯片标题
 * - wps_ppt_insert_image: 插入图片
 * - wps_ppt_set_shape_text: 设置形状文字
 * - wps_ppt_set_animation: 设置元素动画
 * - wps_ppt_set_background: 设置幻灯片背景
 * - wps_ppt_set_slide_size: 设置幻灯片尺寸
 * - wps_ppt_set_transition: 设置幻灯片切换效果
 * - wps_ppt_add_chart: 在幻灯片中插入图表
 * - wps_ppt_set_shape_fill: 设置形状填充颜色
 * - wps_ppt_add_speaker_notes: 添加演讲者备注
 */

import { v4 as uuidv4 } from 'uuid';
import { impactText, type RangeImpact } from '../impact';
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
// 1. wps_ppt_delete_slide - 删除指定幻灯片
// ============================================================

export const deleteSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_delete_slide',
  description: `删除指定的幻灯片。

使用场景：
- "删除第3页幻灯片"
- "把最后一页删掉"
- "移除多余的页面"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '要删除的幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const deleteSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      impact?: RangeImpact;
    }>(
      'deleteSlide',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片删除成功！\n已删除: 第 ${slideIndex} 页${impactText(response.data?.impact)}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 2. wps_ppt_duplicate_slide - 复制幻灯片
// ============================================================



// ============================================================
// 3. wps_ppt_move_slide - 移动幻灯片到指定位置
// ============================================================

export const moveSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_move_slide',
  description: `移动幻灯片到指定位置。

使用场景：
- "把第5页移到第2页"
- "把最后一页移到开头"
- "调整幻灯片顺序"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      fromIndex: {
        type: 'number',
        description: '原位置索引（从1开始）',
      },
      toIndex: {
        type: 'number',
        description: '目标位置索引（从1开始）',
      },
    },
    required: ['fromIndex', 'toIndex'],
  },
};

export const moveSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { fromIndex, toIndex } = args as {
    fromIndex: number;
    toIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'moveSlide',
      { fromIndex, toIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片移动成功！\n从第 ${fromIndex} 页 → 移到第 ${toIndex} 页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `移动幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `移动幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 4. wps_ppt_get_slide_count - 获取幻灯片总数
// ============================================================

export const getSlideCountDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_count',
  description: `获取演示文稿中的幻灯片总数。

使用场景：
- "一共有多少页幻灯片"
- "PPT有几页"
- "查看幻灯片数量"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getSlideCountHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      count: number;
    }>(
      'getSlideCount',
      {},
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `当前演示文稿共有 ${response.data.count} 页幻灯片`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取幻灯片总数失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取幻灯片总数出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 5. wps_ppt_get_slide_info - 获取指定幻灯片的详细信息
// ============================================================

export const getSlideInfoDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_info',
  description: `获取指定幻灯片的详细信息，包括布局、元素列表等。

使用场景：
- "查看第3页的信息"
- "这页有什么内容"
- "获取幻灯片详情"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const getSlideInfoHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideIndex: number;
      layout: string;
      shapesCount: number;
      shapes: Array<{
        name: string;
        type: string;
        text?: string;
      }>;
    }>(
      'getSlideInfo',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const info = response.data;
      let output = `第 ${info.slideIndex} 页幻灯片信息：\n`;
      output += `布局: ${info.layout}\n`;
      output += `元素数量: ${info.shapesCount}\n`;

      if (info.shapes && info.shapes.length > 0) {
        output += `\n元素列表：\n`;
        info.shapes.forEach((shape, i) => {
          output += `  ${i + 1}. [${shape.type}] ${shape.name}`;
          if (shape.text) {
            output += ` - "${shape.text.substring(0, 50)}${shape.text.length > 50 ? '...' : ''}"`;
          }
          output += '\n';
        });
      }

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取幻灯片信息失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取幻灯片信息出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 6. wps_ppt_switch_slide - 切换到指定幻灯片
// ============================================================

export const switchSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_switch_slide',
  description: `切换到指定的幻灯片页面。

使用场景：
- "切换到第5页"
- "跳到最后一页"
- "显示第1页"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '目标幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const switchSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'switchSlide',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已切换到第 ${slideIndex} 页幻灯片`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 7. wps_ppt_set_slide_layout - 设置幻灯片版式布局
// ============================================================

export const setSlideLayoutDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_layout',
  description: `设置幻灯片的版式布局。

支持的布局类型：
- title: 标题页
- title_content: 标题+内容
- blank: 空白页
- two_column: 两栏内容
- comparison: 对比布局
- section_header: 节标题
- title_only: 仅标题

使用场景：
- "把这页改成空白布局"
- "设置为两栏内容"
- "改成标题页版式"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
      layout: {
        type: 'string',
        description: '布局名称',
      },
    },
    required: ['slideIndex', 'layout'],
  },
};

export const setSlideLayoutHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, layout } = args as {
    slideIndex: number;
    layout: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideLayout',
      { slideIndex, layout },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const layoutName: Record<string, string> = {
        title: '标题页',
        title_content: '标题+内容',
        blank: '空白页',
        two_column: '两栏内容',
        comparison: '对比布局',
        section_header: '节标题',
        title_only: '仅标题',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片布局设置成功！\n幻灯片: 第 ${slideIndex} 页\n布局: ${layoutName[layout] || layout}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片布局失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片布局出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 8. wps_ppt_get_slide_notes - 获取幻灯片备注内容
// ============================================================

export const getSlideNotesDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_notes',
  description: `获取指定幻灯片的备注内容。

使用场景：
- "查看第3页的备注"
- "读取演讲备注"
- "这页有什么备注"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const getSlideNotesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      notes: string;
    }>(
      'getSlideNotes',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const notes = response.data.notes;
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: notes
              ? `第 ${slideIndex} 页备注内容：\n${notes}`
              : `第 ${slideIndex} 页没有备注内容`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取幻灯片备注失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取幻灯片备注出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 9. wps_ppt_set_slide_notes - 设置幻灯片备注
// ============================================================

export const setSlideNotesDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_notes',
  description: `设置幻灯片的备注内容，用于演讲提示。

使用场景：
- "给第1页添加备注"
- "写一些演讲提示"
- "修改备注内容"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
      notes: {
        type: 'string',
        description: '备注内容',
      },
    },
    required: ['slideIndex', 'notes'],
  },
};

export const setSlideNotesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, notes } = args as {
    slideIndex: number;
    notes: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideNotes',
      { slideIndex, notes },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片备注设置成功！\n幻灯片: 第 ${slideIndex} 页\n备注内容: "${notes.substring(0, 80)}${notes.length > 80 ? '...' : ''}"`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片备注失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片备注出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 10. wps_ppt_add_shape - 添加形状
// ============================================================

export const addShapeDefinition: ToolDefinition = {
  name: 'wps_ppt_add_shape',
  description: `在幻灯片中添加形状。

支持的形状类型：
- rectangle: 矩形
- oval: 椭圆
- triangle: 三角形
- diamond: 菱形
- pentagon: 五边形
- hexagon: 六边形
- arrow: 箭头
- star: 星形
- heart: 心形
- cloud: 云形

使用场景：
- "在第1页添加一个矩形"
- "插入一个蓝色的圆形"
- "添加一个带文字的箭头"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始），默认1',
      },
      type: {
        type: 'string',
        description: '形状类型',
        enum: ['rectangle', 'oval', 'triangle', 'diamond', 'pentagon', 'hexagon', 'arrow', 'star', 'heart', 'cloud'],
      },
      left: {
        type: 'number',
        description: '左边距（像素），默认100',
      },
      top: {
        type: 'number',
        description: '上边距（像素），默认100',
      },
      width: {
        type: 'number',
        description: '宽度（像素），默认100',
      },
      height: {
        type: 'number',
        description: '高度（像素），默认100',
      },
      text: {
        type: 'string',
        description: '形状内的文本',
      },
      fillColor: {
        type: 'string',
        description: '填充颜色，十六进制如 #FF0000',
      },
    },
    required: [],
  },
};

export const addShapeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, type, left, top, width, height, text, fillColor } = args as {
    slideIndex?: number;
    type?: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
    text?: string;
    fillColor?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
      slideIndex: number;
    }>(
      'addShape',
      {
        slideIndex: slideIndex || 1,
        type: type || 'rectangle',
        left: left || 100,
        top: top || 100,
        width: width || 100,
        height: height || 100,
        text,
        fillColor,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const shapeNameMap: Record<string, string> = {
        rectangle: '矩形', oval: '椭圆', triangle: '三角形', diamond: '菱形',
        pentagon: '五边形', hexagon: '六边形', arrow: '箭头', star: '星形',
        heart: '心形', cloud: '云形',
      };
      const typeName = shapeNameMap[type || 'rectangle'] || type;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状添加成功！\n幻灯片: 第 ${response.data.slideIndex} 页\n形状类型: ${typeName}\n形状名称: ${response.data.name}${text ? `\n文本: ${text}` : ''}${fillColor ? `\n填充色: ${fillColor}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加形状失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加形状出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 11. wps_ppt_set_shape_style - 设置形状样式
// ============================================================

export const setShapeStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_style',
  description: `设置幻灯片中形状的样式，包括填充颜色、边框颜色和边框粗细。

使用场景：
- "把矩形改成红色"
- "设置形状的边框为蓝色"
- "修改形状样式"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始），默认1',
      },
      name: {
        type: 'string',
        description: '形状名称（通过getSlideInfo获取）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（与name二选一）',
      },
      fillColor: {
        type: 'string',
        description: '填充颜色，十六进制如 #FF0000',
      },
      lineColor: {
        type: 'string',
        description: '边框颜色，十六进制如 #000000',
      },
      lineWidth: {
        type: 'number',
        description: '边框粗细（磅）',
      },
    },
    required: [],
  },
};

export const setShapeStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, name, shapeIndex, fillColor, lineColor, lineWidth } = args as {
    slideIndex?: number;
    name?: string;
    shapeIndex?: number;
    fillColor?: string;
    lineColor?: string;
    lineWidth?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
    }>(
      'setShapeStyle',
      {
        slideIndex: slideIndex || 1,
        name,
        shapeIndex,
        fillColor,
        lineColor,
        lineWidth,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      let output = `形状样式设置成功！\n形状: ${response.data.name}`;
      if (fillColor) output += `\n填充色: ${fillColor}`;
      if (lineColor) output += `\n边框色: ${lineColor}`;
      if (lineWidth) output += `\n边框粗细: ${lineWidth}pt`;

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状样式失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状样式出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 12. wps_ppt_add_textbox - 添加文本框
// ============================================================

export const addTextboxDefinition: ToolDefinition = {
  name: 'wps_ppt_add_textbox',
  description: `在幻灯片中添加文本框。

使用场景：
- "在第1页添加一个文本框"
- "插入一个写着标题的文本框"
- "添加文本框并设置字号"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始），默认1',
      },
      left: {
        type: 'number',
        description: '左边距（像素），默认100',
      },
      top: {
        type: 'number',
        description: '上边距（像素），默认100',
      },
      width: {
        type: 'number',
        description: '宽度（像素），默认200',
      },
      height: {
        type: 'number',
        description: '高度（像素），默认50',
      },
      text: {
        type: 'string',
        description: '文本框内容',
      },
      fontSize: {
        type: 'number',
        description: '字号大小',
      },
      fontName: {
        type: 'string',
        description: '字体名称，如 "微软雅黑"',
      },
    },
    required: [],
  },
};

export const addTextboxHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, left, top, width, height, text, fontSize, fontName } = args as {
    slideIndex?: number;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
    text?: string;
    fontSize?: number;
    fontName?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name?: string;
      slideIndex?: number;
    }>(
      'addTextBox',
      {
        slideIndex: slideIndex || 1,
        left: left || 100,
        top: top || 100,
        width: width || 200,
        height: height || 50,
        text,
        fontSize,
        fontName,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      // Report what was asked for: the action's payload does not always carry these back, and the
      // message used to print "第 undefined 页 / 名称: undefined".
      const targetSlide = response.data?.slideIndex ?? slideIndex ?? 1;
      const shapeName = response.data?.name ? `\n名称: ${response.data.name}` : '';
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文本框添加成功！\n幻灯片: 第 ${targetSlide} 页${shapeName}${text ? `\n内容: "${text}"` : ''}${fontSize ? `\n字号: ${fontSize}` : ''}${fontName ? `\n字体: ${fontName}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加文本框失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加文本框出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 13. wps_ppt_set_slide_title - 设置幻灯片标题
// ============================================================

export const setSlideTitleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_title',
  description: `设置幻灯片的标题文本。

使用场景：
- "把第1页标题改成'年度总结'"
- "设置标题为'项目进展'"
- "修改幻灯片标题"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始），默认1',
      },
      title: {
        type: 'string',
        description: '标题文本',
      },
    },
    required: ['title'],
  },
};

export const setSlideTitleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, title } = args as {
    slideIndex?: number;
    title: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideIndex: number;
      title: string;
    }>(
      'setSlideTitle',
      {
        slideIndex: slideIndex || 1,
        title,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片标题设置成功！\n幻灯片: 第 ${response.data.slideIndex} 页\n标题: "${response.data.title}"`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片标题失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片标题出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 14. wps_ppt_insert_image - 插入图片
// ============================================================



// ============================================================
// 15. wps_ppt_set_shape_text - 设置形状文字
// ============================================================

export const setShapeTextDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_text',
  description: `设置幻灯片中指定形状的文字内容。

使用场景：
- "把第1页的第2个形状文字改成'销售额'"
- "修改形状里的文字"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      shapeIndex: { type: 'number', description: '形状索引（从1开始）' },
      text: { type: 'string', description: '要设置的文字内容' },
    },
    required: ['slideIndex', 'shapeIndex', 'text'],
  },
};

export const setShapeTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, text } = args as {
    slideIndex: number;
    shapeIndex: number;
    text: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
    }>(
      'setShapeText',
      { slideIndex, shapeIndex, text },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `形状文字设置成功！\n形状: ${response.data.name}\n文字: "${text}"` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状文字失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状文字出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 16. wps_ppt_set_animation - 设置元素动画
// ============================================================



// ============================================================
// 17. wps_ppt_set_background - 设置幻灯片背景
// ============================================================



// ============================================================
// 18. wps_ppt_set_slide_size - 设置幻灯片尺寸
// ============================================================

export const setSlideSizeDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_size',
  description: `设置演示文稿的幻灯片尺寸。

常用尺寸：
- 标准(4:3): 宽960, 高720
- 宽屏(16:9): 宽960, 高540
- 宽屏(16:10): 宽960, 高600
- A4横版: 宽1123, 高794
- A4竖版: 宽794, 高1123

使用场景：
- "把PPT改成16:9宽屏"
- "设置幻灯片为A4尺寸"
- "修改幻灯片大小"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      width: {
        type: 'number',
        description: '幻灯片宽度（像素）',
      },
      height: {
        type: 'number',
        description: '幻灯片高度（像素）',
      },
    },
    required: ['width', 'height'],
  },
};

export const setSlideSizeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { width, height } = args as {
    width: number;
    height: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideSize', // NOTE: macOS未实现，仅Windows支持
      { width, height },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片尺寸设置成功！\n宽度: ${width}px\n高度: ${height}px`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片尺寸失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片尺寸出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 19. wps_ppt_set_transition - 设置幻灯片切换效果
// ============================================================



// ============================================================
// 20. wps_ppt_add_chart - 在幻灯片中插入图表
// ============================================================



// ============================================================
// 21. wps_ppt_set_shape_fill - 设置形状填充颜色
// ============================================================

export const setShapeFillDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_fill',
  description: `设置幻灯片中指定形状的填充颜色。

使用场景：
- "把第1页的第2个形状填充为红色"
- "修改形状背景色为#00FF00"
- "设置形状的填充颜色"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      shapeIndex: { type: 'number', description: '形状索引（从1开始）' },
      color: { type: 'string', description: '填充颜色，十六进制如 #FF0000' },
    },
    required: ['slideIndex', 'shapeIndex', 'color'],
  },
};

export const setShapeFillHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, color } = args as {
    slideIndex: number;
    shapeIndex: number;
    color: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name?: string;
    }>(
      'setShapeFill', // NOTE: macOS未实现，仅Windows支持
      { slideIndex, shapeIndex, color },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状填充颜色设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n填充颜色: ${color}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状填充失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状填充出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 22. wps_ppt_add_speaker_notes - 添加演讲者备注
// ============================================================



// ============================================================
// 导出所有幻灯片操作相关的Tools
// ============================================================

export const slideOpsTools: RegisteredTool[] = [
  { definition: deleteSlideDefinition, handler: deleteSlideHandler },
  { definition: moveSlideDefinition, handler: moveSlideHandler },
  { definition: getSlideCountDefinition, handler: getSlideCountHandler },
  { definition: getSlideInfoDefinition, handler: getSlideInfoHandler },
  { definition: switchSlideDefinition, handler: switchSlideHandler },
  { definition: setSlideLayoutDefinition, handler: setSlideLayoutHandler },
  { definition: getSlideNotesDefinition, handler: getSlideNotesHandler },
  { definition: setSlideNotesDefinition, handler: setSlideNotesHandler },
  { definition: addShapeDefinition, handler: addShapeHandler },
  { definition: setShapeStyleDefinition, handler: setShapeStyleHandler },
  { definition: addTextboxDefinition, handler: addTextboxHandler },
  { definition: setSlideTitleDefinition, handler: setSlideTitleHandler },
  { definition: setShapeTextDefinition, handler: setShapeTextHandler },
  { definition: setSlideSizeDefinition, handler: setSlideSizeHandler },
  { definition: setShapeFillDefinition, handler: setShapeFillHandler },
];

export default slideOpsTools;

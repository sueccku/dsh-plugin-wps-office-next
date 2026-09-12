/**
 * Input: PPT 形状操作参数（删除、获取、位置、阴影、渐变、边框、透明度、对齐、分布、组合）
 * Output: 形状操作结果
 * Pos: PPT 形状基础工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT形状基础Tools - 形状管理模块
 * 处理形状的删除、获取、定位、样式设置、对齐分布和组合操作
 *
 * 包含：
 * - wps_ppt_delete_shape: 删除形状
 * - wps_ppt_get_shapes: 获取幻灯片中的形状列表
 * - wps_ppt_set_shape_position: 设置形状位置和大小
 * - wps_ppt_set_shape_shadow: 设置形状阴影
 * - wps_ppt_set_shape_gradient: 设置形状渐变填充
 * - wps_ppt_set_shape_border: 设置形状边框
 * - wps_ppt_set_shape_transparency: 设置形状透明度
 * - wps_ppt_align_shapes: 对齐多个形状
 * - wps_ppt_distribute_shapes: 等距分布多个形状
 * - wps_ppt_group_shapes: 组合多个形状
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

// ==================== 1. 删除形状 ====================

export const deleteShapeDefinition: ToolDefinition = {
  name: 'wps_ppt_delete_shape',
  description: `删除幻灯片中指定的形状。

使用场景：
- "删除第2页的第3个形状"
- "移除这个形状"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
    },
    required: ['slideIndex', 'shapeIndex'],
  },
};

export const deleteShapeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex } = args as {
    slideIndex: number;
    shapeIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'deleteShape',
      { slideIndex, shapeIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状删除成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除形状失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除形状出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 2. 获取形状列表 ====================

export const getShapesDefinition: ToolDefinition = {
  name: 'wps_ppt_get_shapes',
  description: `获取幻灯片中所有形状的列表信息。

返回每个形状的类型、位置、大小等属性。

使用场景：
- "这页PPT有哪些形状"
- "列出第3页的所有元素"`,
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

export const getShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as {
    slideIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      shapes: Array<{
        index: number;
        name: string;
        type: string;
        left: number;
        top: number;
        width: number;
        height: number;
      }>;
    }>(
      'getShapes',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const shapes = response.data.shapes;
      let output = `幻灯片第 ${slideIndex} 页的形状列表（共 ${shapes.length} 个）：\n\n`;
      shapes.forEach((s) => {
        output += `[${s.index}] ${s.name} (${s.type})\n`;
        output += `    位置: (${s.left}, ${s.top}) 大小: ${s.width} x ${s.height}\n`;
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
        content: [{ type: 'text', text: `获取形状列表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取形状列表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 3. 设置形状位置 ====================

export const setShapePositionDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_position',
  description: `设置幻灯片中指定形状的位置和大小。

使用场景：
- "把这个形状移到左上角"
- "调整形状大小"
- "设置形状位置为(100, 200)"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      left: {
        type: 'number',
        description: '左边距（磅）',
      },
      top: {
        type: 'number',
        description: '上边距（磅）',
      },
      width: {
        type: 'number',
        description: '宽度（磅），可选',
      },
      height: {
        type: 'number',
        description: '高度（磅），可选',
      },
    },
    required: ['slideIndex', 'shapeIndex', 'left', 'top'],
  },
};

export const setShapePositionHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, left, top, width, height } = args as {
    slideIndex: number;
    shapeIndex: number;
    left: number;
    top: number;
    width?: number;
    height?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setShapePosition',
      { slideIndex, shapeIndex, left, top, width, height },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `形状位置设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n位置: (${left}, ${top})`;
      if (width !== undefined) text += `\n宽度: ${width}`;
      if (height !== undefined) text += `\n高度: ${height}`;

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状位置失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状位置出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 4. 设置形状阴影 ====================

export const setShapeShadowDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_shadow',
  description: `设置幻灯片中指定形状的阴影效果。

shadow对象属性：
- enabled: 是否启用阴影 (boolean)
- color: 阴影颜色，如 "#000000"
- blur: 模糊半径（磅）
- offsetX: 水平偏移（磅）
- offsetY: 垂直偏移（磅）
- opacity: 透明度 (0-1)

使用场景：
- "给这个形状加阴影"
- "设置阴影效果"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      shadow: {
        type: 'object',
        description: '阴影配置对象',
        properties: {
          enabled: { type: 'boolean', description: '是否启用阴影' },
          color: { type: 'string', description: '阴影颜色' },
          blur: { type: 'number', description: '模糊半径（磅）' },
          offsetX: { type: 'number', description: '水平偏移（磅）' },
          offsetY: { type: 'number', description: '垂直偏移（磅）' },
          opacity: { type: 'number', description: '透明度 (0-1)' },
        },
      },
    },
    required: ['slideIndex', 'shapeIndex', 'shadow'],
  },
};

export const setShapeShadowHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, shadow } = args as {
    slideIndex: number;
    shapeIndex: number;
    shadow: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setShapeShadow',
      { slideIndex, shapeIndex, shadow },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状阴影设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状阴影失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状阴影出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 5. 设置形状渐变填充 ====================

export const setShapeGradientDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_gradient',
  description: `设置幻灯片中指定形状的渐变填充效果。

gradient对象属性：
- type: 渐变类型，"linear"(线性) 或 "radial"(径向)
- angle: 渐变角度（0-360，仅线性渐变）
- stops: 渐变色标数组 [{color: "#FF0000", position: 0}, {color: "#0000FF", position: 1}]

使用场景：
- "给形状加渐变色"
- "设置从红到蓝的渐变"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      gradient: {
        type: 'object',
        description: '渐变配置对象',
        properties: {
          type: { type: 'string', description: '渐变类型: linear 或 radial', enum: ['linear', 'radial'] },
          angle: { type: 'number', description: '渐变角度（0-360）' },
          stops: {
            type: 'array',
            description: '渐变色标数组',
            items: {
              type: 'object',
              properties: {
                color: { type: 'string', description: '颜色值' },
                position: { type: 'number', description: '位置 (0-1)' },
              },
            },
          },
        },
      },
    },
    required: ['slideIndex', 'shapeIndex', 'gradient'],
  },
};

export const setShapeGradientHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, gradient } = args as {
    slideIndex: number;
    shapeIndex: number;
    gradient: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setShapeGradient',
      { slideIndex, shapeIndex, gradient },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状渐变填充设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状渐变失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状渐变出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 6. 设置形状边框 ====================

export const setShapeBorderDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_border',
  description: `设置幻灯片中指定形状的边框样式。

border对象属性：
- enabled: 是否启用边框 (boolean)
- color: 边框颜色，如 "#000000"
- weight: 边框粗细（磅）
- style: 边框样式，如 "solid"(实线)、"dash"(虚线)、"dot"(点线)

使用场景：
- "给形状加边框"
- "设置红色虚线边框"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      border: {
        type: 'object',
        description: '边框配置对象',
        properties: {
          enabled: { type: 'boolean', description: '是否启用边框' },
          color: { type: 'string', description: '边框颜色' },
          weight: { type: 'number', description: '边框粗细（磅）' },
          style: { type: 'string', description: '边框样式', enum: ['solid', 'dash', 'dot', 'dash_dot', 'dash_dot_dot'] },
        },
      },
    },
    required: ['slideIndex', 'shapeIndex', 'border'],
  },
};

export const setShapeBorderHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, border } = args as {
    slideIndex: number;
    shapeIndex: number;
    border: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setShapeBorder',
      { slideIndex, shapeIndex, border },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状边框设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状边框失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状边框出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 7. 设置形状透明度 ====================

export const setShapeTransparencyDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_transparency',
  description: `设置幻灯片中指定形状的透明度。

使用场景：
- "把这个形状设为半透明"
- "设置形状透明度为50%"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      transparency: {
        type: 'number',
        description: '透明度值 (0-100)，0为完全不透明，100为完全透明',
      },
    },
    required: ['slideIndex', 'shapeIndex', 'transparency'],
  },
};

export const setShapeTransparencyHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, transparency } = args as {
    slideIndex: number;
    shapeIndex: number;
    transparency: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setShapeTransparency',
      { slideIndex, shapeIndex, transparency },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状透明度设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n透明度: ${transparency}%`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状透明度失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状透明度出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 8. 对齐多个形状 ====================

export const alignShapesDefinition: ToolDefinition = {
  name: 'wps_ppt_align_shapes',
  description: `对齐幻灯片中的多个形状。

支持的对齐方式：
- left: 左对齐
- center: 水平居中
- right: 右对齐
- top: 顶部对齐
- middle: 垂直居中
- bottom: 底部对齐

使用场景：
- "把这几个形状左对齐"
- "让第1、3、5个形状垂直居中"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndices: {
        type: 'array',
        items: { type: 'number' },
        description: '要对齐的形状索引数组（从1开始）',
      },
      alignment: {
        type: 'string',
        description: '对齐方式',
        enum: ['left', 'center', 'right', 'top', 'middle', 'bottom'],
      },
    },
    required: ['slideIndex', 'shapeIndices', 'alignment'],
  },
};

export const alignShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndices, alignment } = args as {
    slideIndex: number;
    shapeIndices: number[];
    alignment: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      count?: number;
    }>(
      'alignShapes',
      { slideIndex, shapeIndices, alignment },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const alignName: Record<string, string> = {
        left: '左对齐',
        center: '水平居中',
        right: '右对齐',
        top: '顶部对齐',
        middle: '垂直居中',
        bottom: '底部对齐',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状对齐完成！\n幻灯片: 第 ${slideIndex} 页\n对齐方式: ${alignName[alignment] || alignment}\n形状数量: ${shapeIndices.length} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `对齐形状失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `对齐形状出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 9. 等距分布多个形状 ====================

export const distributeShapesDefinition: ToolDefinition = {
  name: 'wps_ppt_distribute_shapes',
  description: `等距分布幻灯片中的多个形状。

支持的分布方向：
- horizontal: 水平等距分布
- vertical: 垂直等距分布

使用场景：
- "让这些形状水平等距排列"
- "垂直均匀分布这些元素"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndices: {
        type: 'array',
        items: { type: 'number' },
        description: '要分布的形状索引数组（从1开始）',
      },
      direction: {
        type: 'string',
        description: '分布方向',
        enum: ['horizontal', 'vertical'],
      },
    },
    required: ['slideIndex', 'shapeIndices', 'direction'],
  },
};

export const distributeShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndices, direction } = args as {
    slideIndex: number;
    shapeIndices: number[];
    direction: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'distributeShapes',
      { slideIndex, shapeIndices, direction },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const dirName = direction === 'horizontal' ? '水平等距' : '垂直等距';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状分布完成！\n幻灯片: 第 ${slideIndex} 页\n分布方向: ${dirName}\n形状数量: ${shapeIndices.length} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `分布形状失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `分布形状出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 10. 组合多个形状 ====================

export const groupShapesDefinition: ToolDefinition = {
  name: 'wps_ppt_group_shapes',
  description: `将幻灯片中的多个形状组合为一个组。

使用场景：
- "把这几个形状组合在一起"
- "将第1、2、3个形状编组"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndices: {
        type: 'array',
        items: { type: 'number' },
        description: '要组合的形状索引数组（从1开始，至少2个）',
      },
    },
    required: ['slideIndex', 'shapeIndices'],
  },
};

export const groupShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndices } = args as {
    slideIndex: number;
    shapeIndices: number[];
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      groupIndex?: number;
    }>(
      'groupShapes',
      { slideIndex, shapeIndices },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `形状组合成功！\n幻灯片: 第 ${slideIndex} 页\n组合形状: ${shapeIndices.length} 个`;
      if (response.data?.groupIndex) {
        text += `\n组合索引: ${response.data.groupIndex}`;
      }

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `组合形状失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `组合形状出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有形状基础相关的Tools
 */
export const shapeBasicTools: RegisteredTool[] = [
  { definition: deleteShapeDefinition, handler: deleteShapeHandler },
  { definition: getShapesDefinition, handler: getShapesHandler },
  { definition: setShapePositionDefinition, handler: setShapePositionHandler },
  { definition: setShapeShadowDefinition, handler: setShapeShadowHandler },
  { definition: setShapeGradientDefinition, handler: setShapeGradientHandler },
  { definition: setShapeBorderDefinition, handler: setShapeBorderHandler },
  { definition: setShapeTransparencyDefinition, handler: setShapeTransparencyHandler },
  { definition: alignShapesDefinition, handler: alignShapesHandler },
  { definition: distributeShapesDefinition, handler: distributeShapesHandler },
  { definition: groupShapesDefinition, handler: groupShapesHandler },
];

export default shapeBasicTools;

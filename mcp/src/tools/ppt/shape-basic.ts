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

// ==================== 5. 设置形状渐变填充 ====================

// ==================== 6. 设置形状边框 ====================

// ==================== 7. 设置形状透明度 ====================

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
    // shapeIndices is optional: without it every shape on the slide is aligned.
    required: ['slideIndex', 'alignment'],
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
      shapes?: number;
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
            // shapeIndices is optional: without it the action aligns every shape on the slide, and
            // dereferencing it here used to fail with "Cannot read properties of undefined".
            text: `形状对齐完成！\n幻灯片: 第 ${slideIndex} 页\n对齐方式: ${alignName[alignment] || alignment}\n形状数量: ${shapeIndices ? shapeIndices.length + ' 个' : (response.data?.shapes ?? '全部')}`,
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
/** 形状效果：阴影 / 边框 / 渐变 / 透明度（P4-2 由四个碎片 setter 合并而来） */
export const setShapeEffectDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_effect',
  description: '给形状设置视觉效果：阴影、边框、渐变填充、填充透明度；只改给出来的项，其余保持原样。使用场景：给标题条加阴影、给方框加描边、给卡片做渐变。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      shapeIndex: { type: 'number', description: '形状序号（从 1 开始）；给了 name 就用 name' },
      name: { type: 'string', description: '形状名称' },
      slideIndex: { type: 'number', description: '第几页（从 1 开始），默认 1' },
      presentationName: { type: 'string', description: '演示文稿名；不填用当前文稿' },
      shadowEnabled: { type: 'boolean', description: '是否显示阴影' },
      shadowColor: { type: 'string', description: '阴影颜色，十六进制如 #333333' },
      shadowTransparency: { type: 'number', description: '阴影透明度 0-1' },
      shadowBlur: { type: 'number', description: '阴影模糊半径' },
      shadowOffsetX: { type: 'number', description: '阴影水平偏移' },
      shadowOffsetY: { type: 'number', description: '阴影垂直偏移' },
      borderEnabled: { type: 'boolean', description: '是否显示边框' },
      borderColor: { type: 'string', description: '边框颜色，十六进制如 #1A365D' },
      borderWidth: { type: 'number', description: '边框粗细（磅）' },
      borderStyle: { type: 'string', enum: ['solid', 'dash', 'dot', 'dash_dot', 'dash_dot_dot'], description: '边框线型' },
      gradientColor1: { type: 'string', description: '渐变起始色' },
      gradientColor2: { type: 'string', description: '渐变结束色' },
      transparency: { type: 'number', description: '填充透明度 0-1' },
    },
  },
};

export const setShapeEffectHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ name?: string; applied?: string[] }>(
      'setShapeEffect',
      {
        presentationName: args.presentationName, slideIndex: args.slideIndex, name: args.name, shapeIndex: args.shapeIndex,
        shadowEnabled: args.shadowEnabled, shadowColor: args.shadowColor, shadowTransparency: args.shadowTransparency,
        shadowBlur: args.shadowBlur, shadowOffsetX: args.shadowOffsetX, shadowOffsetY: args.shadowOffsetY,
        borderEnabled: args.borderEnabled, borderColor: args.borderColor, borderWidth: args.borderWidth, borderStyle: args.borderStyle,
        gradientColor1: args.gradientColor1, gradientColor2: args.gradientColor2, transparency: args.transparency,
      },
      WpsAppType.PRESENTATION
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: '设置形状效果失败: ' + response.error }], error: response.error };
    }
    const applied = response.data?.applied || [];
    const head = '形状 ' + (response.data?.name || '') + (applied.length ? ' 已更新（' + applied.join(', ') + '）' : ' 没有变化');
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: head }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: '设置形状效果出错: ' + errMsg }], error: errMsg };
  }
};

export const shapeBasicTools: RegisteredTool[] = [
  { definition: setShapeEffectDefinition, handler: setShapeEffectHandler },
  { definition: deleteShapeDefinition, handler: deleteShapeHandler },
  { definition: getShapesDefinition, handler: getShapesHandler },
  { definition: setShapePositionDefinition, handler: setShapePositionHandler },
  { definition: alignShapesDefinition, handler: alignShapesHandler },
  { definition: distributeShapesDefinition, handler: distributeShapesHandler },
  { definition: groupShapesDefinition, handler: groupShapesHandler },
];

export default shapeBasicTools;

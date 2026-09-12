/**
 * Input: PPT 背景、页脚、形状操作参数
 * Output: 背景/页脚/形状操作结果
 * Pos: PPT 背景与页面信息工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT背景Tools - 背景、页面信息、高级形状模块
 * 处理幻灯片背景设置、页码/页脚/日期时间、形状复制与层级调整
 *
 * 包含：
 * - wps_ppt_set_slide_background: 设置幻灯片背景（通用）
 * - wps_ppt_set_background_color: 设置幻灯片背景颜色
 * - wps_ppt_set_background_image: 设置幻灯片背景图片
 * - wps_ppt_set_slide_number: 设置幻灯片页码显示
 * - wps_ppt_set_ppt_footer: 设置PPT页脚
 * - wps_ppt_set_ppt_date_time: 设置PPT日期时间
 * - wps_ppt_duplicate_shape: 复制形状
 * - wps_ppt_set_shape_z_order: 设置形状层级顺序
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

// ==================== 背景 (3) ====================

/**
 * 设置幻灯片背景（通用）
 * 支持纯色、渐变、图片等多种背景类型
 */
export const setSlideBackgroundDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_background',
  description: `设置幻灯片背景，支持多种背景类型。

支持的背景类型（background.type）：
- solid: 纯色背景，需提供 color 字段
- gradient: 渐变背景，需提供 colors 数组
- image: 图片背景，需提供 imagePath 字段
- pattern: 图案背景，需提供 pattern 和 color 字段

使用场景：
- "给第1页设置蓝色背景"
- "设置渐变背景"
- "用图片做背景"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      background: {
        type: 'object',
        description: '背景配置对象，包含 type/color/colors/imagePath/pattern 等字段',
      },
    },
    required: ['slideIndex', 'background'],
  },
};

export const setSlideBackgroundHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, background } = args as {
    slideIndex: number;
    background: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideBackground',
      { slideIndex, background },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const bgType = (background.type as string) || 'solid';
      const typeName: Record<string, string> = {
        solid: '纯色',
        gradient: '渐变',
        image: '图片',
        pattern: '图案',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片背景设置成功！\n幻灯片: 第 ${slideIndex} 页\n背景类型: ${typeName[bgType] || bgType}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片背景失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片背景出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置幻灯片背景颜色
 * 快速设置纯色背景的便捷方法
 */
export const setBackgroundColorDefinition: ToolDefinition = {
  name: 'wps_ppt_set_background_color',
  description: `设置幻灯片背景为指定颜色。

使用场景：
- "把第1页背景改成蓝色"
- "设置背景颜色为 #FF5733"
- "背景换成白色"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      color: {
        type: 'string',
        description: '十六进制颜色值，如 "#FF0000"、"#FFFFFF"',
      },
    },
    required: ['slideIndex', 'color'],
  },
};

export const setBackgroundColorHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, color } = args as {
    slideIndex: number;
    color: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setBackgroundColor',
      { slideIndex, color },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `背景颜色设置成功！\n幻灯片: 第 ${slideIndex} 页\n颜色: ${color}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置背景颜色失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置背景颜色出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置幻灯片背景图片
 * 使用图片文件作为幻灯片背景
 */
export const setBackgroundImageDefinition: ToolDefinition = {
  name: 'wps_ppt_set_background_image',
  description: `设置幻灯片背景为指定图片。

使用场景：
- "用这张图片做背景"
- "设置第2页的背景图片"
- "把图片设为幻灯片背景"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      imagePath: {
        type: 'string',
        description: '图片文件的完整路径',
      },
    },
    required: ['slideIndex', 'imagePath'],
  },
};

export const setBackgroundImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, imagePath } = args as {
    slideIndex: number;
    imagePath: string;
  };

  try {
    // 跨平台参数对齐：macOS/Windows 底层均读取 params.path，需同时发送 path/filePath 别名
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setBackgroundImage',
      { slideIndex, imagePath, path: imagePath, filePath: imagePath },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `背景图片设置成功！\n幻灯片: 第 ${slideIndex} 页\n图片: ${imagePath}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置背景图片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置背景图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 页面信息 (3) ====================

/**
 * 设置幻灯片页码显示
 * 控制页码的显示/隐藏和起始编号
 */
export const setSlideNumberDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_number',
  description: `设置幻灯片页码的显示状态和起始编号。

使用场景：
- "显示页码"
- "隐藏幻灯片编号"
- "页码从第2页开始编号"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      show: {
        type: 'boolean',
        description: '是否显示页码',
      },
      startFrom: {
        type: 'number',
        description: '页码起始编号，默认为1',
      },
    },
    required: ['show'],
  },
};

export const setSlideNumberHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { show, startFrom } = args as {
    show: boolean;
    startFrom?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideNumber',
      { show, startFrom },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `幻灯片页码设置成功！\n显示状态: ${show ? '显示' : '隐藏'}`;
      if (startFrom !== undefined) {
        text += `\n起始编号: ${startFrom}`;
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
        content: [{ type: 'text', text: `设置页码失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页码出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置PPT页脚
 * 控制页脚文本的显示/隐藏
 */
export const setPptFooterDefinition: ToolDefinition = {
  name: 'wps_ppt_set_ppt_footer',
  description: `设置演示文稿页脚文本。

使用场景：
- "添加页脚'公司名称'"
- "设置页脚为'机密文件'"
- "隐藏页脚"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '页脚文本内容',
      },
      show: {
        type: 'boolean',
        description: '是否显示页脚，默认为true',
      },
    },
    required: ['text'],
  },
};

export const setPptFooterHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, show } = args as {
    text: string;
    show?: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setPptFooter',
      { text, show: show !== false },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `页脚设置成功！\n页脚文本: ${text}\n显示状态: ${show !== false ? '显示' : '隐藏'}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页脚失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页脚出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置PPT日期时间
 * 控制日期时间占位符的显示和格式
 */
export const setPptDateTimeDefinition: ToolDefinition = {
  name: 'wps_ppt_set_ppt_date_time',
  description: `设置演示文稿日期时间显示。

支持的日期格式（format）：
- "YYYY-MM-DD": 如 2026-03-21
- "YYYY/MM/DD": 如 2026/03/21
- "MM/DD/YYYY": 如 03/21/2026
- "DD/MM/YYYY": 如 21/03/2026

使用场景：
- "显示日期时间"
- "设置自动更新日期"
- "隐藏日期"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      show: {
        type: 'boolean',
        description: '是否显示日期时间',
      },
      autoUpdate: {
        type: 'boolean',
        description: '是否自动更新日期时间，默认为true',
      },
      format: {
        type: 'string',
        description: '日期时间格式，如 "YYYY-MM-DD"',
      },
    },
    required: ['show'],
  },
};

export const setPptDateTimeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { show, autoUpdate, format } = args as {
    show: boolean;
    autoUpdate?: boolean;
    format?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setPptDateTime',
      { show, autoUpdate, format },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `日期时间设置成功！\n显示状态: ${show ? '显示' : '隐藏'}`;
      if (show) {
        text += `\n自动更新: ${autoUpdate !== false ? '是' : '否'}`;
        if (format) {
          text += `\n格式: ${format}`;
        }
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
        content: [{ type: 'text', text: `设置日期时间失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置日期时间出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 高级形状 (2) ====================

/**
 * 复制形状
 * 在同一幻灯片内复制指定形状
 */
export const duplicateShapeDefinition: ToolDefinition = {
  name: 'wps_ppt_duplicate_shape',
  description: `复制幻灯片中的指定形状。

使用场景：
- "复制这个形状"
- "把第1页的第2个形状复制一份"
- "克隆这个元素"`,
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
        description: '要复制的形状索引（从1开始）',
      },
    },
    required: ['slideIndex', 'shapeIndex'],
  },
};

export const duplicateShapeHandler: ToolHandler = async (
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
      newShapeIndex?: number;
    }>(
      'duplicateShape',
      { slideIndex, shapeIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `形状复制成功！\n幻灯片: 第 ${slideIndex} 页\n源形状: 第 ${shapeIndex} 个`;
      if (response.data?.newShapeIndex) {
        text += `\n新形状索引: ${response.data.newShapeIndex}`;
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
        content: [{ type: 'text', text: `复制形状失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `复制形状出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置形状层级顺序
 * 调整形状在幻灯片中的前后层级
 */
export const setShapeZOrderDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_z_order',
  description: `设置形状在幻灯片中的层级顺序（Z轴排列）。

支持的层级操作（order）：
- front: 置于顶层
- back: 置于底层
- forward: 上移一层
- backward: 下移一层

使用场景：
- "把这个形状移到最前面"
- "将形状置于底层"
- "上移一层"`,
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
      order: {
        type: 'string',
        description: '层级操作类型',
        enum: ['front', 'back', 'forward', 'backward'],
      },
    },
    required: ['slideIndex', 'shapeIndex', 'order'],
  },
};

export const setShapeZOrderHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, order } = args as {
    slideIndex: number;
    shapeIndex: number;
    order: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setShapeZOrder',
      { slideIndex, shapeIndex, order },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const orderName: Record<string, string> = {
        front: '置于顶层',
        back: '置于底层',
        forward: '上移一层',
        backward: '下移一层',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状层级调整成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n操作: ${orderName[order] || order}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `调整形状层级失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `调整形状层级出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有背景、页面信息、高级形状相关的Tools
 */
export const backgroundTools: RegisteredTool[] = [
  { definition: setSlideBackgroundDefinition, handler: setSlideBackgroundHandler },
  { definition: setBackgroundColorDefinition, handler: setBackgroundColorHandler },
  { definition: setBackgroundImageDefinition, handler: setBackgroundImageHandler },
  { definition: setSlideNumberDefinition, handler: setSlideNumberHandler },
  { definition: setPptFooterDefinition, handler: setPptFooterHandler },
  { definition: setPptDateTimeDefinition, handler: setPptDateTimeHandler },
  { definition: duplicateShapeDefinition, handler: duplicateShapeHandler },
  { definition: setShapeZOrderDefinition, handler: setShapeZOrderHandler },
];

export default backgroundTools;

/**
 * Input: PPT 图片操作参数（插入、删除、样式设置、幻灯片导出为图片）
 * Output: 图片操作结果
 * Pos: PPT 图片工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT图片Tools - 图片管理模块
 * 处理图片的插入、删除、样式设置以及幻灯片导出为位图操作
 *
 * 包含：
 * - wps_ppt_insert_ppt_image: 插入图片到幻灯片
 * - wps_ppt_delete_ppt_image: 删除幻灯片中的图片
 * - wps_ppt_set_image_style: 设置图片样式
 * - wps_ppt_export_slide_as_image: 将指定幻灯片导出为位图图片（PNG/JPG/GIF/BMP）
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

// ==================== 1. 插入图片 ====================

export const insertPptImageDefinition: ToolDefinition = {
  name: 'wps_ppt_insert_ppt_image',
  description: `插入图片到幻灯片中。

支持常见图片格式（PNG、JPG、BMP、GIF等）。
可以指定图片的位置和大小，不指定则使用默认值。

使用场景：
- "在第2页插入一张图片"
- "把这张图片放到PPT里"
- "在幻灯片右下角插入logo"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      filePath: {
        type: 'string',
        description: '图片文件路径',
      },
      left: {
        type: 'number',
        description: '左边距（磅），可选',
      },
      top: {
        type: 'number',
        description: '上边距（磅），可选',
      },
      width: {
        type: 'number',
        description: '宽度（磅），可选，不指定则按原始比例',
      },
      height: {
        type: 'number',
        description: '高度（磅），可选，不指定则按原始比例',
      },
    },
    required: ['slideIndex', 'filePath'],
  },
};

export const insertPptImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, filePath, left, top, width, height } = args as {
    slideIndex: number;
    filePath: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  };

  try {
    // 跨平台参数对齐：macOS/Windows 底层均读取 params.path，需同时发送 path/imagePath 别名
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      imageIndex?: number;
    }>(
      'insertPptImage',
      { slideIndex, filePath, path: filePath, imagePath: filePath, left, top, width, height },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `图片插入成功！\n幻灯片: 第 ${slideIndex} 页\n文件: ${filePath}`;
      if (left !== undefined && top !== undefined) text += `\n位置: (${left}, ${top})`;
      if (width !== undefined) text += `\n宽度: ${width}`;
      if (height !== undefined) text += `\n高度: ${height}`;
      if (response.data?.imageIndex) text += `\n图片索引: ${response.data.imageIndex}`;

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入图片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 2. 删除图片 ====================

export const deletePptImageDefinition: ToolDefinition = {
  name: 'wps_ppt_delete_ppt_image',
  description: `删除幻灯片中指定的图片。

使用场景：
- "删除第2页的第1张图片"
- "移除这张图片"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      imageIndex: {
        type: 'number',
        description: '图片索引（从1开始）',
      },
    },
    required: ['slideIndex', 'imageIndex'],
  },
};

export const deletePptImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, imageIndex } = args as {
    slideIndex: number;
    imageIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'deletePptImage',
      { slideIndex, imageIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图片删除成功！\n幻灯片: 第 ${slideIndex} 页\n图片: 第 ${imageIndex} 张`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除图片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 3. 设置图片样式 ====================

export const setImageStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_image_style',
  description: `设置幻灯片中指定图片的样式。

style对象属性：
- border: 边框设置 {enabled: boolean, color: string, weight: number}
- shadow: 阴影设置 {enabled: boolean, color: string, blur: number, offsetX: number, offsetY: number}
- opacity: 透明度 (0-100)
- rotation: 旋转角度 (0-360)
- cropTop/cropBottom/cropLeft/cropRight: 裁剪比例 (0-1)

使用场景：
- "给图片加边框"
- "设置图片阴影效果"
- "旋转图片45度"
- "裁剪图片"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      imageIndex: {
        type: 'number',
        description: '图片索引（从1开始）',
      },
      style: {
        type: 'object',
        description: '图片样式配置对象',
        properties: {
          border: {
            type: 'object',
            description: '边框设置',
            properties: {
              enabled: { type: 'boolean', description: '是否启用边框' },
              color: { type: 'string', description: '边框颜色' },
              weight: { type: 'number', description: '边框粗细（磅）' },
            },
          },
          shadow: {
            type: 'object',
            description: '阴影设置',
            properties: {
              enabled: { type: 'boolean', description: '是否启用阴影' },
              color: { type: 'string', description: '阴影颜色' },
              blur: { type: 'number', description: '模糊半径' },
              offsetX: { type: 'number', description: '水平偏移' },
              offsetY: { type: 'number', description: '垂直偏移' },
            },
          },
          opacity: { type: 'number', description: '透明度 (0-100)' },
          rotation: { type: 'number', description: '旋转角度 (0-360)' },
          cropTop: { type: 'number', description: '顶部裁剪比例 (0-1)' },
          cropBottom: { type: 'number', description: '底部裁剪比例 (0-1)' },
          cropLeft: { type: 'number', description: '左侧裁剪比例 (0-1)' },
          cropRight: { type: 'number', description: '右侧裁剪比例 (0-1)' },
        },
      },
    },
    required: ['slideIndex', 'imageIndex', 'style'],
  },
};

export const setImageStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, imageIndex, style } = args as {
    slideIndex: number;
    imageIndex: number;
    style: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setImageStyle',
      { slideIndex, imageIndex, style },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图片样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n图片: 第 ${imageIndex} 张`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置图片样式失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置图片样式出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 4. 导出幻灯片为图片 ====================

export const exportSlideAsImageDefinition: ToolDefinition = {
  name: 'wps_ppt_export_slide_as_image',
  description: `将指定幻灯片导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。

调用底层 WPS PowerPoint 原生接口 Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight)，
实现 1:1 像素级还原，避免通过 PDF 中转再转图片造成的版式/字体/形状失真问题。

支持的 format（FilterName）取值：
- PNG（默认，推荐用于截图与无损展示）
- JPG / JPEG（自动按 JPG 滤镜处理，体积更小）
- GIF（限 256 色，适合简单图形）
- BMP（无压缩位图，文件最大）

使用场景：
- "把第3页 PPT 导出成 PNG 给我"
- "导出整个演示文稿每一页为 1920x1080 的 JPG"
- "把封面页保存为高清图片用于网页"

注意：
- outputPath 必须是绝对路径
- macOS 上建议输出到 ~/Downloads 或用户可写目录，避免沙箱权限拒绝
- 不指定 width/height 时使用 1280x720（16:9 默认尺寸）`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片序号（从1开始）',
      },
      outputPath: {
        type: 'string',
        description: '输出图片文件的绝对路径（如 /Users/xxx/Downloads/slide1.png）',
      },
      format: {
        type: 'string',
        enum: ['PNG', 'JPG', 'JPEG', 'GIF', 'BMP'],
        description: '图片格式，默认 PNG',
      },
      width: {
        type: 'number',
        description: '输出图片宽度（像素），默认 1280',
      },
      height: {
        type: 'number',
        description: '输出图片高度（像素），默认 720',
      },
    },
    required: ['slideIndex', 'outputPath'],
  },
};

export const exportSlideAsImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, outputPath, format, width, height } = args as {
    slideIndex: number;
    outputPath: string;
    format?: string;
    width?: number;
    height?: number;
  };

  // 归一化 format：JPEG 在 WPS COM 中实际对应 JPG 滤镜
  const rawFormat = (format || 'PNG').toUpperCase();
  const filterName = rawFormat === 'JPEG' ? 'JPG' : rawFormat;
  const finalWidth = width || 1280;
  const finalHeight = height || 720;

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message?: string;
      slideIndex?: number;
      outputPath?: string;
      format?: string;
    }>(
      'exportSlideAsImage',
      {
        slideIndex,
        outputPath,
        // 跨平台参数对齐：macOS/Windows 底层兼容 path/outputPath 双别名
        path: outputPath,
        format: filterName,
        width: finalWidth,
        height: finalHeight,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const text =
        `幻灯片导出图片成功！\n` +
        `幻灯片: 第 ${slideIndex} 页\n` +
        `格式: ${filterName}\n` +
        `尺寸: ${finalWidth} x ${finalHeight}\n` +
        `输出路径: ${outputPath}`;

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `导出幻灯片为图片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `导出幻灯片为图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 5. 原位替换图片（换图不动版式） ====================

export const replacePptImageDefinition: ToolDefinition = {
  name: 'wps_ppt_replace_ppt_image',
  description: `原位替换幻灯片中的某张图片：保留原图的位置、尺寸、旋转角度，删除旧图后在同一矩形内插入新图。用于在保持版式不变的前提下把模板里的旧图换成自己的图。

使用场景：
- "把第5页的第2个形状（图片）换成 D:/figs/图15.png，位置大小不变"
- "替换封面主视觉图，但保持原来的排版"
- "把这页的配图换掉，别动版式"

说明：
- 先用 wps_ppt_get_shapes 查到目标图片的形状索引(shapeIndex)
- 新图会被拉伸/适配到旧图原有的位置与尺寸，从而不破坏版式`,
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
        description: '要替换的图片形状索引（从1开始，通过 get_shapes 获取）',
      },
      name: {
        type: 'string',
        description: '要替换的图片形状名称（与 shapeIndex 二选一）',
      },
      filePath: {
        type: 'string',
        description: '新图片文件的完整路径',
      },
    },
    required: ['slideIndex', 'filePath'],
  },
};

export const replacePptImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, name, filePath } = args as {
    slideIndex: number;
    shapeIndex?: number;
    name?: string;
    filePath: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
      left: number;
      top: number;
      width: number;
      height: number;
    }>(
      'replacePptImage',
      { slideIndex, shapeIndex, name, filePath, path: filePath, imagePath: filePath },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图片已原位替换（位置尺寸不变）！\n幻灯片: 第 ${slideIndex} 页\n新图: ${filePath}\n位置: (${Math.round(response.data.left)}, ${Math.round(response.data.top)})  尺寸: ${Math.round(response.data.width)} x ${Math.round(response.data.height)}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `替换图片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `替换图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有图片相关的Tools
 */
export const imageTools: RegisteredTool[] = [
  { definition: insertPptImageDefinition, handler: insertPptImageHandler },
  { definition: deletePptImageDefinition, handler: deletePptImageHandler },
  { definition: setImageStyleDefinition, handler: setImageStyleHandler },
  { definition: exportSlideAsImageDefinition, handler: exportSlideAsImageHandler },
  { definition: replacePptImageDefinition, handler: replacePptImageHandler },
];

export default imageTools;

/**
 * Input: PPT 高级美化操作参数
 * Output: 高级美化操作结果
 * Pos: PPT 高级美化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT高级美化Tools - 高级美化模块
 * 处理配色方案、自动美化、KPI卡片、装饰元素等高级美化操作
 *
 * 包含：
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
  { definition: setBackgroundGradientDefinition, handler: setBackgroundGradientHandler },
];

export default beautifyAdvancedTools;

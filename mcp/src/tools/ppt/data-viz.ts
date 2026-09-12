/**
 * Input: PPT 数据可视化与布局操作参数
 * Output: 数据可视化与布局操作结果
 * Pos: PPT 数据可视化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT数据可视化Tools - 进度条/仪表盘/迷你图/环形图/布局模块
 *
 * 包含：
 * - wps_ppt_create_progress_bar: 创建进度条
 * - wps_ppt_create_gauge: 创建仪表盘
 * - wps_ppt_create_mini_charts: 创建迷你图表
 * - wps_ppt_create_donut_chart: 创建环形图
 * - wps_ppt_auto_layout: 自动排版
 * - wps_ppt_smart_distribute: 智能分布
 * - wps_ppt_create_grid: 创建网格布局
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

// ==================== 数据可视化 ====================

/**
 * 创建进度条
 * 在幻灯片中创建可视化进度条
 */
export const createProgressBarDefinition: ToolDefinition = {
  name: 'wps_ppt_create_progress_bar',
  description: `在幻灯片中创建进度条。

可以指定进度值和标签文字，生成美观的进度条可视化。

使用场景：
- "添加一个80%的进度条"
- "创建项目完成度进度条"
- "做一个进度条展示KPI"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      value: {
        type: 'number',
        description: '进度值（0-100）',
      },
      label: {
        type: 'string',
        description: '可选的标签文字，如 "项目进度"、"完成率"',
      },
    },
    required: ['slideIndex', 'value'],
  },
};

export const createProgressBarHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, value, label } = args as {
    slideIndex: number;
    value: number;
    label?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      shapeId?: number;
    }>(
      'createProgressBar',
      { slideIndex, value, label: label || '' },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `进度条创建成功！\n幻灯片: 第 ${slideIndex} 页\n进度: ${value}%${label ? `\n标签: ${label}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建进度条失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建进度条出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 创建仪表盘
 * 在幻灯片中创建仪表盘图表
 */
export const createGaugeDefinition: ToolDefinition = {
  name: 'wps_ppt_create_gauge',
  description: `在幻灯片中创建仪表盘图表。

用于展示KPI、完成度等数值指标。

使用场景：
- "创建一个仪表盘展示得分"
- "添加速度计样式的图表"
- "做一个KPI仪表盘"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      value: {
        type: 'number',
        description: '当前值',
      },
      max: {
        type: 'number',
        description: '最大值，默认100',
      },
    },
    required: ['slideIndex', 'value'],
  },
};

export const createGaugeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, value, max } = args as {
    slideIndex: number;
    value: number;
    max?: number;
  };

  const maxValue = max || 100;

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      shapeId?: number;
    }>(
      'createGauge',
      { slideIndex, value, max: maxValue },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `仪表盘创建成功！\n幻灯片: 第 ${slideIndex} 页\n当前值: ${value}\n最大值: ${maxValue}\n百分比: ${Math.round((value / maxValue) * 100)}%`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建仪表盘失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建仪表盘出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 创建迷你图表
 * 在幻灯片中创建迷你图表（sparkline风格）
 */
export const createMiniChartsDefinition: ToolDefinition = {
  name: 'wps_ppt_create_mini_charts',
  description: `在幻灯片中创建迷你图表。

基于数据数组生成紧凑的sparkline风格图表。

使用场景：
- "创建一组迷你趋势图"
- "添加sparkline图表"
- "用小图表展示数据趋势"`,
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
        description: '数据数组，如 [{label:"Q1",values:[10,20,30]},{label:"Q2",values:[15,25,35]}]',
        items: { type: 'object' },
      },
    },
    required: ['slideIndex', 'data'],
  },
};

export const createMiniChartsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, data } = args as {
    slideIndex: number;
    data: Array<Record<string, unknown>>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      chartCount: number;
    }>(
      'createMiniCharts',
      { slideIndex, data },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `迷你图表创建成功！\n幻灯片: 第 ${slideIndex} 页\n图表数量: ${response.data.chartCount} 个\n数据组数: ${data.length}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建迷你图表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建迷你图表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 创建环形图
 * 在幻灯片中创建环形图（Donut Chart）
 */
export const createDonutChartDefinition: ToolDefinition = {
  name: 'wps_ppt_create_donut_chart',
  description: `在幻灯片中创建环形图。

用于展示占比、分布等数据。

使用场景：
- "创建一个环形图展示市场份额"
- "做一个甜甜圈图表"
- "用环形图展示各部门占比"`,
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
        description: '数据数组，如 [{label:"产品A",value:40},{label:"产品B",value:30},{label:"其他",value:30}]',
        items: { type: 'object' },
      },
    },
    required: ['slideIndex', 'data'],
  },
};

export const createDonutChartHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, data } = args as {
    slideIndex: number;
    data: Array<Record<string, unknown>>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      shapeId?: number;
    }>(
      'createDonutChart',
      { slideIndex, data },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `环形图创建成功！\n幻灯片: 第 ${slideIndex} 页\n数据项: ${data.length} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建环形图失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建环形图出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 布局操作 ====================

/**
 * 自动排版
 * 智能调整幻灯片中所有元素的布局
 */
export const autoLayoutDefinition: ToolDefinition = {
  name: 'wps_ppt_auto_layout',
  description: `自动调整幻灯片中所有元素的布局。

智能分析元素位置和大小，自动优化排版。

使用场景：
- "自动排版这一页"
- "整理一下布局"
- "让元素自动对齐排列"`,
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

export const autoLayoutHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as {
    slideIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      adjustedCount?: number;
    }>(
      'autoLayout',
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
            text: `自动排版完成！\n幻灯片: 第 ${slideIndex} 页${response.data?.adjustedCount ? `\n调整元素: ${response.data.adjustedCount} 个` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `自动排版失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `自动排版出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 智能分布
 * 将指定形状进行等距分布排列
 */
export const smartDistributeDefinition: ToolDefinition = {
  name: 'wps_ppt_smart_distribute',
  description: `将指定形状进行等距分布排列。

自动计算间距，让选中的形状均匀分布。

使用场景：
- "让这些形状等距排列"
- "均匀分布选中的对象"
- "智能排列这些元素"`,
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
        description: '要分布的形状索引数组，如 [1,2,3,4]',
        items: { type: 'number' },
      },
    },
    required: ['slideIndex', 'shapeIndices'],
  },
};

export const smartDistributeHandler: ToolHandler = async (
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
    }>(
      'smartDistribute',
      { slideIndex, shapeIndices },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `智能分布完成！\n幻灯片: 第 ${slideIndex} 页\n处理形状: ${shapeIndices.length} 个 (索引: ${shapeIndices.join(', ')})`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `智能分布失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `智能分布出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 创建网格布局
 * 在幻灯片中创建网格布局结构
 */
export const createGridDefinition: ToolDefinition = {
  name: 'wps_ppt_create_grid',
  description: `在幻灯片中创建网格布局。

按行列数自动生成均匀分布的网格占位区域。

使用场景：
- "创建一个2x3的网格布局"
- "做一个九宫格"
- "建立4列的网格排版"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      rows: {
        type: 'number',
        description: '行数',
      },
      cols: {
        type: 'number',
        description: '列数',
      },
    },
    required: ['slideIndex', 'rows', 'cols'],
  },
};

export const createGridHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, rows, cols } = args as {
    slideIndex: number;
    rows: number;
    cols: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      cellCount?: number;
    }>(
      'createGrid',
      { slideIndex, rows, cols },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `网格布局创建成功！\n幻灯片: 第 ${slideIndex} 页\n布局: ${rows} 行 x ${cols} 列\n单元格: ${rows * cols} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建网格布局失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建网格布局出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有数据可视化与布局Tools
 */
export const dataVizTools: RegisteredTool[] = [
  // 数据可视化
  { definition: createProgressBarDefinition, handler: createProgressBarHandler },
  { definition: createGaugeDefinition, handler: createGaugeHandler },
  { definition: createMiniChartsDefinition, handler: createMiniChartsHandler },
  { definition: createDonutChartDefinition, handler: createDonutChartHandler },
  // 布局操作
  { definition: autoLayoutDefinition, handler: autoLayoutHandler },
  { definition: smartDistributeDefinition, handler: smartDistributeHandler },
  { definition: createGridDefinition, handler: createGridHandler },
];

export default dataVizTools;

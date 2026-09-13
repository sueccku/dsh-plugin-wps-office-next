/**
 * Input: PPT 图表与流程图操作参数
 * Output: 图表/流程图操作结果
 * Pos: PPT 图表与流程图工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT图表Tools - 图表、流程图、组织架构图与时间线管理模块
 * 处理PPT中图表插入与样式设置、流程图/组织架构图/时间线创建
 *
 * 包含：
 * - wps_ppt_insert_ppt_chart: 插入图表
 * - wps_ppt_set_ppt_chart_data: 设置图表数据
 * - wps_ppt_set_ppt_chart_style: 设置图表样式
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

// ==================== 1. 插入图表 ====================

/**
 * 插入图表
 * 在幻灯片中插入数据图表
 */
export const insertPptChartDefinition: ToolDefinition = {
  name: 'wps_ppt_insert_ppt_chart',
  description: `在幻灯片中插入数据图表。

支持的图表类型：
- bar: 柱状图
- line: 折线图
- pie: 饼图
- area: 面积图
- scatter: 散点图
- doughnut: 环形图
- radar: 雷达图

使用场景：
- "在第2页插入一个柱状图"
- "添加销售数据的饼图"
- "插入折线图展示趋势"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      chartType: {
        type: 'string',
        description: '图表类型',
      },
      title: {
        type: 'string',
        description: '图表标题',
      },
      left: {
        type: 'number',
        description: '图表左边距（磅），默认自动居中',
      },
      top: {
        type: 'number',
        description: '图表上边距（磅），默认自动居中',
      },
    },
    // Chart data is not injectable: filling a chart means opening its embedded workbook, which a
    // resident COM host can leave behind. The parameter was removed rather than silently ignored.
    required: ['slideIndex', 'chartType'],
  },
};

export const insertPptChartHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, chartType, title, left, top } = args as {
    slideIndex: number;
    chartType: string;
    title?: string;
    left?: number;
    top?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      chartIndex: number;
    }>(
      'insertPptChart',
      { slideIndex, chartType, title, left, top },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const chartTypeName: Record<string, string> = {
        bar: '柱状图',
        line: '折线图',
        pie: '饼图',
        area: '面积图',
        scatter: '散点图',
        doughnut: '环形图',
        radar: '雷达图',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图表插入成功！\n幻灯片: 第 ${slideIndex} 页\n类型: ${chartTypeName[chartType] || chartType}\n图表索引: ${response.data.chartIndex}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入图表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入图表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 2. 设置图表数据 ====================

/**
 * 设置图表数据
 * 更新已有图表的数据
 */
export const setPptChartDataDefinition: ToolDefinition = {
  name: 'wps_ppt_set_ppt_chart_data',
  description: `更新幻灯片中已有图表的数据。

使用场景：
- "更新第2页图表的数据"
- "修改图表数据"
- "替换图表中的数值"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      chartIndex: {
        type: 'number',
        description: '图表索引（从1开始）',
      },
      data: {
        type: 'object',
        description: '新的图表数据，包含 categories（类别数组）和 series（系列数组）',
      },
    },
    required: ['slideIndex', 'chartIndex', 'data'],
  },
};

export const setPptChartDataHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, chartIndex, data } = args as {
    slideIndex: number;
    chartIndex: number;
    data: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setPptChartData',
      { slideIndex, chartIndex, data },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图表数据更新成功！\n幻灯片: 第 ${slideIndex} 页\n图表: 第 ${chartIndex} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `更新图表数据失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `更新图表数据出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 3. 设置图表样式 ====================

/**
 * 设置图表样式
 * 修改图表的颜色方案、字体、图例等样式属性
 */
export const setPptChartStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_ppt_chart_style',
  description: `设置幻灯片中图表的样式属性。

可设置的样式属性（通过 style 对象传入）：
- colorScheme: 配色方案名称
- showLegend: 是否显示图例
- legendPosition: 图例位置（top/bottom/left/right）
- showDataLabels: 是否显示数据标签
- fontSize: 字体大小
- title: 图表标题

使用场景：
- "修改图表配色为蓝色系"
- "显示图表的数据标签"
- "把图例移到底部"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      chartIndex: {
        type: 'number',
        description: '图表索引（从1开始）',
      },
      style: {
        type: 'object',
        description: '图表样式配置对象',
      },
    },
    required: ['slideIndex', 'chartIndex', 'style'],
  },
};

export const setPptChartStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, chartIndex, style } = args as {
    slideIndex: number;
    chartIndex: number;
    style: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setPptChartStyle',
      { slideIndex, chartIndex, style },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const styleItems = Object.keys(style).join('、');
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图表样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n图表: 第 ${chartIndex} 个\n已更新属性: ${styleItems}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置图表样式失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置图表样式出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 4. 创建流程图 ====================
export const chartFlowTools: RegisteredTool[] = [
  { definition: insertPptChartDefinition, handler: insertPptChartHandler },
  { definition: setPptChartDataDefinition, handler: setPptChartDataHandler },
  { definition: setPptChartStyleDefinition, handler: setPptChartStyleHandler },
];

export default chartFlowTools;

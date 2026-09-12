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
 * - wps_ppt_create_flow_chart: 创建流程图
 * - wps_ppt_create_org_chart: 创建组织架构图
 * - wps_ppt_create_timeline: 创建时间线
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
      data: {
        type: 'object',
        description: '图表数据，包含 categories（类别数组）和 series（系列数组，每个系列含 name 和 values）',
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
    required: ['slideIndex', 'chartType', 'data'],
  },
};

export const insertPptChartHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, chartType, data, left, top } = args as {
    slideIndex: number;
    chartType: string;
    data: Record<string, unknown>;
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
      { slideIndex, chartType, data, left, top },
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

/**
 * 创建流程图
 * 在幻灯片中创建流程图，自动布局节点和连接线
 */
export const createFlowChartDefinition: ToolDefinition = {
  name: 'wps_ppt_create_flow_chart',
  description: `在幻灯片中创建流程图。

nodes 数组中每个节点包含：
- id: 节点标识符
- text: 节点显示文本
- type: 节点类型（process/decision/start/end）

connections 数组中每个连接包含：
- from: 起始节点id
- to: 目标节点id
- label: 连接线标签（可选）

使用场景：
- "创建一个审批流程图"
- "画一个工作流程图"
- "制作决策流程图"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      nodes: {
        type: 'array',
        description: '流程图节点数组',
        items: {
          type: 'object',
          properties: {
            id: { type: 'string', description: '节点唯一标识' },
            text: { type: 'string', description: '节点显示文本' },
            type: { type: 'string', description: '节点类型: process/decision/start/end' },
          },
          required: ['id', 'text'],
        },
      },
      connections: {
        type: 'array',
        description: '节点之间的连接关系数组（可选）',
        items: {
          type: 'object',
          properties: {
            from: { type: 'string', description: '起始节点id' },
            to: { type: 'string', description: '目标节点id' },
            label: { type: 'string', description: '连接线标签' },
          },
          required: ['from', 'to'],
        },
      },
    },
    required: ['slideIndex', 'nodes'],
  },
};

export const createFlowChartHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, nodes, connections } = args as {
    slideIndex: number;
    nodes: Array<Record<string, unknown>>;
    connections?: Array<Record<string, unknown>>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      nodeCount: number;
      connectionCount: number;
    }>(
      'createFlowChart',
      { slideIndex, nodes, connections: connections || [] },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `流程图创建成功！\n幻灯片: 第 ${slideIndex} 页\n节点数: ${response.data.nodeCount} 个\n连接线: ${response.data.connectionCount} 条`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建流程图失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建流程图出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 5. 创建组织架构图 ====================

/**
 * 创建组织架构图
 * 根据层级数据自动生成组织架构图
 */
export const createOrgChartDefinition: ToolDefinition = {
  name: 'wps_ppt_create_org_chart',
  description: `在幻灯片中创建组织架构图。

data 对象格式（树形结构）：
- name: 节点名称（如人名/职位）
- title: 节点副标题（如部门/角色）
- children: 子节点数组（递归结构）

使用场景：
- "创建公司组织架构图"
- "画一个部门层级图"
- "生成团队架构图"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      data: {
        type: 'object',
        description: '组织架构数据（树形结构），含 name、title、children 字段',
      },
    },
    required: ['slideIndex', 'data'],
  },
};

export const createOrgChartHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, data } = args as {
    slideIndex: number;
    data: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      nodeCount: number;
      levelCount: number;
    }>(
      'createOrgChart',
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
            text: `组织架构图创建成功！\n幻灯片: 第 ${slideIndex} 页\n节点数: ${response.data.nodeCount} 个\n层级数: ${response.data.levelCount} 层`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建组织架构图失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建组织架构图出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 6. 创建时间线 ====================

/**
 * 创建时间线
 * 根据事件数据在幻灯片中生成时间线图形
 */
export const createTimelineDefinition: ToolDefinition = {
  name: 'wps_ppt_create_timeline',
  description: `在幻灯片中创建时间线。

events 数组中每个事件包含：
- date: 日期/时间标签
- title: 事件标题
- description: 事件描述（可选）

使用场景：
- "创建项目里程碑时间线"
- "画一个公司发展时间线"
- "制作历史大事记时间线"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      events: {
        type: 'array',
        description: '时间线事件数组',
        items: {
          type: 'object',
          properties: {
            date: { type: 'string', description: '日期/时间标签' },
            title: { type: 'string', description: '事件标题' },
            description: { type: 'string', description: '事件描述' },
          },
          required: ['date', 'title'],
        },
      },
    },
    required: ['slideIndex', 'events'],
  },
};

export const createTimelineHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, events } = args as {
    slideIndex: number;
    events: Array<Record<string, unknown>>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      eventCount: number;
    }>(
      'createTimeline',
      { slideIndex, events },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `时间线创建成功！\n幻灯片: 第 ${slideIndex} 页\n事件数: ${response.data.eventCount} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建时间线失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建时间线出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有图表与流程图相关的Tools
 */
export const chartFlowTools: RegisteredTool[] = [
  { definition: insertPptChartDefinition, handler: insertPptChartHandler },
  { definition: setPptChartDataDefinition, handler: setPptChartDataHandler },
  { definition: setPptChartStyleDefinition, handler: setPptChartStyleHandler },
  { definition: createFlowChartDefinition, handler: createFlowChartHandler },
  { definition: createOrgChartDefinition, handler: createOrgChartHandler },
  { definition: createTimelineDefinition, handler: createTimelineHandler },
];

export default chartFlowTools;

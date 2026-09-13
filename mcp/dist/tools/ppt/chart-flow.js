"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.chartFlowTools = exports.setPptChartStyleHandler = exports.setPptChartStyleDefinition = exports.setPptChartDataHandler = exports.setPptChartDataDefinition = exports.insertPptChartHandler = exports.insertPptChartDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
// ==================== 1. 插入图表 ====================
/**
 * 插入图表
 * 在幻灯片中插入数据图表
 */
exports.insertPptChartDefinition = {
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
    category: tools_1.ToolCategory.PRESENTATION,
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
const insertPptChartHandler = async (args) => {
    const { slideIndex, chartType, title, left, top } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertPptChart', { slideIndex, chartType, title, left, top }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const chartTypeName = {
                bar: '柱状图',
                line: '折线图',
                pie: '饼图',
                area: '面积图',
                scatter: '散点图',
                doughnut: '环形图',
                radar: '雷达图',
            };
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `图表插入成功！\n幻灯片: 第 ${slideIndex} 页\n类型: ${chartTypeName[chartType] || chartType}\n图表索引: ${response.data.chartIndex}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `插入图表失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `插入图表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertPptChartHandler = insertPptChartHandler;
// ==================== 2. 设置图表数据 ====================
/**
 * 设置图表数据
 * 更新已有图表的数据
 */
exports.setPptChartDataDefinition = {
    name: 'wps_ppt_set_ppt_chart_data',
    description: `更新幻灯片中已有图表的数据。

使用场景：
- "更新第2页图表的数据"
- "修改图表数据"
- "替换图表中的数值"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const setPptChartDataHandler = async (args) => {
    const { slideIndex, chartIndex, data } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setPptChartData', { slideIndex, chartIndex, data }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `图表数据更新成功！\n幻灯片: 第 ${slideIndex} 页\n图表: 第 ${chartIndex} 个`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `更新图表数据失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `更新图表数据出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setPptChartDataHandler = setPptChartDataHandler;
// ==================== 3. 设置图表样式 ====================
/**
 * 设置图表样式
 * 修改图表的颜色方案、字体、图例等样式属性
 */
exports.setPptChartStyleDefinition = {
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
    category: tools_1.ToolCategory.PRESENTATION,
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
const setPptChartStyleHandler = async (args) => {
    const { slideIndex, chartIndex, style } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setPptChartStyle', { slideIndex, chartIndex, style }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const styleItems = Object.keys(style).join('、');
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `图表样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n图表: 第 ${chartIndex} 个\n已更新属性: ${styleItems}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置图表样式失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置图表样式出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setPptChartStyleHandler = setPptChartStyleHandler;
// ==================== 4. 创建流程图 ====================
exports.chartFlowTools = [
    { definition: exports.insertPptChartDefinition, handler: exports.insertPptChartHandler },
    { definition: exports.setPptChartDataDefinition, handler: exports.setPptChartDataHandler },
    { definition: exports.setPptChartStyleDefinition, handler: exports.setPptChartStyleHandler },
];
exports.default = exports.chartFlowTools;
//# sourceMappingURL=chart-flow.js.map
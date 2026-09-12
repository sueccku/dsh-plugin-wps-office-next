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
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 创建进度条
 * 在幻灯片中创建可视化进度条
 */
export declare const createProgressBarDefinition: ToolDefinition;
export declare const createProgressBarHandler: ToolHandler;
/**
 * 创建仪表盘
 * 在幻灯片中创建仪表盘图表
 */
export declare const createGaugeDefinition: ToolDefinition;
export declare const createGaugeHandler: ToolHandler;
/**
 * 创建迷你图表
 * 在幻灯片中创建迷你图表（sparkline风格）
 */
export declare const createMiniChartsDefinition: ToolDefinition;
export declare const createMiniChartsHandler: ToolHandler;
/**
 * 创建环形图
 * 在幻灯片中创建环形图（Donut Chart）
 */
export declare const createDonutChartDefinition: ToolDefinition;
export declare const createDonutChartHandler: ToolHandler;
/**
 * 自动排版
 * 智能调整幻灯片中所有元素的布局
 */
export declare const autoLayoutDefinition: ToolDefinition;
export declare const autoLayoutHandler: ToolHandler;
/**
 * 智能分布
 * 将指定形状进行等距分布排列
 */
export declare const smartDistributeDefinition: ToolDefinition;
export declare const smartDistributeHandler: ToolHandler;
/**
 * 创建网格布局
 * 在幻灯片中创建网格布局结构
 */
export declare const createGridDefinition: ToolDefinition;
export declare const createGridHandler: ToolHandler;
/**
 * 导出所有数据可视化与布局Tools
 */
export declare const dataVizTools: RegisteredTool[];
export default dataVizTools;
//# sourceMappingURL=data-viz.d.ts.map
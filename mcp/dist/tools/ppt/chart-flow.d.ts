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
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 插入图表
 * 在幻灯片中插入数据图表
 */
export declare const insertPptChartDefinition: ToolDefinition;
export declare const insertPptChartHandler: ToolHandler;
/**
 * 设置图表数据
 * 更新已有图表的数据
 */
export declare const setPptChartDataDefinition: ToolDefinition;
export declare const setPptChartDataHandler: ToolHandler;
/**
 * 设置图表样式
 * 修改图表的颜色方案、字体、图例等样式属性
 */
export declare const setPptChartStyleDefinition: ToolDefinition;
export declare const setPptChartStyleHandler: ToolHandler;
/**
 * 创建流程图
 * 在幻灯片中创建流程图，自动布局节点和连接线
 */
export declare const createFlowChartDefinition: ToolDefinition;
export declare const createFlowChartHandler: ToolHandler;
/**
 * 创建组织架构图
 * 根据层级数据自动生成组织架构图
 */
export declare const createOrgChartDefinition: ToolDefinition;
export declare const createOrgChartHandler: ToolHandler;
/**
 * 创建时间线
 * 根据事件数据在幻灯片中生成时间线图形
 */
export declare const createTimelineDefinition: ToolDefinition;
export declare const createTimelineHandler: ToolHandler;
/**
 * 导出所有图表与流程图相关的Tools
 */
export declare const chartFlowTools: RegisteredTool[];
export default chartFlowTools;
//# sourceMappingURL=chart-flow.d.ts.map
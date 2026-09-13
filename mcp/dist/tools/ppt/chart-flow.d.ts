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
export declare const chartFlowTools: RegisteredTool[];
export default chartFlowTools;
//# sourceMappingURL=chart-flow.d.ts.map
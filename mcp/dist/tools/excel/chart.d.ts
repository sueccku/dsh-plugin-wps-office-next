/**
 * Input: 图表工具参数
 * Output: 图表创建/更新/导出结果
 * Pos: Excel 图表工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel图表相关Tools - 数据可视化模块
 * 负责图表的创建、属性更新以及图表/区域导出为图片操作
 *
 * 包含：
 * - wps_excel_create_chart: 创建图表（柱状图、折线图、饼图、散点图等）
 * - wps_excel_update_chart: 更新图表属性（标题、颜色、图例等）
 * - wps_excel_export_chart_as_image: 将工作表中的图表导出为位图（Chart.Export 原生 API）
 * - wps_excel_export_range_as_image: 将指定区域导出为位图（Range.CopyPicture + 临时 ChartObject 经典做法）
 *
 * @date 2026-01-24
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 支持的图表类型枚举
 * WPS Excel支持的主流图表类型
 */
export declare enum ChartType {
    /** 簇状柱形图 - 最常用的，比较数值大小 */
    COLUMN_CLUSTERED = "column_clustered",
    /** 堆积柱形图 - 看整体和部分关系 */
    COLUMN_STACKED = "column_stacked",
    /** 簇状条形图 - 横着的柱形图 */
    BAR_CLUSTERED = "bar_clustered",
    /** 折线图 - 看趋势变化 */
    LINE = "line",
    /** 带数据标记的折线图 */
    LINE_MARKERS = "line_markers",
    /** 饼图 - 看占比 */
    PIE = "pie",
    /** 环形图 - 空心饼图 */
    DOUGHNUT = "doughnut",
    /** 散点图 - 看相关性 */
    SCATTER = "scatter",
    /** 面积图 - 强调数量变化的程度 */
    AREA = "area",
    /** 雷达图 - 多维度对比 */
    RADAR = "radar"
}
/**
 * 创建图表工具
 * 根据数据范围创建各种类型的图表
 */
export declare const createChartDefinition: ToolDefinition;
export declare const createChartHandler: ToolHandler;
/**
 * 更新图表属性工具
 * 支持修改标题、颜色、图例、数据标签等属性
 */
export declare const updateChartDefinition: ToolDefinition;
export declare const updateChartHandler: ToolHandler;
/**
 * 将指定图表导出为位图图片
 *
 * 底层调用 WPS Excel 原生 Chart.Export(FileName, FilterName)：
 * - 最简单可靠的图表导出方案
 * - 不经过剪贴板，无 PDF/截屏中转造成的失真
 *
 * 与 PPT 的 wps_ppt_export_slide_as_image 形成 Issue #15 Roadmap 双子工具，
 * 解决 Claude 通过 Python 转图片造成结构失真的痛点。
 */
export declare const exportChartAsImageDefinition: ToolDefinition;
export declare const exportChartAsImageHandler: ToolHandler;
/**
 * 将指定区域导出为位图图片
 *
 * 经典做法：Range.CopyPicture + 临时 ChartObject + Chart.Paste + Chart.Export：
 * 1. range.CopyPicture(xlScreen=1, xlBitmap=2) 复制区域为位图到剪贴板
 * 2. sheet.ChartObjects.Add(0, 0, range.Width, range.Height) 创建同尺寸临时图表
 * 3. tempChart.Chart.Paste() 把剪贴板位图粘贴到图表
 * 4. tempChart.Chart.Export(outputPath, format) 导出图表为图片
 * 5. tempChart.Delete() 清理临时图表
 *
 * 注意：headless 模式与某些 WPS 版本可能因剪贴板冲突失败，handler 内部需 try/catch + 清理。
 */
export declare const exportRangeAsImageDefinition: ToolDefinition;
export declare const exportRangeAsImageHandler: ToolHandler;
/**
 * 导出所有图表相关的Tools
 */
export declare const chartTools: RegisteredTool[];
export default chartTools;
//# sourceMappingURL=chart.d.ts.map
/**
 * Input: PPT 高级美化操作参数
 * Output: 高级美化操作结果
 * Pos: PPT 高级美化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT高级美化Tools - 高级美化模块
 * 处理配色方案、自动美化、KPI卡片、装饰元素等高级美化操作
 *
 * 包含：
 * - wps_ppt_apply_color_scheme: 应用配色方案
 * - wps_ppt_auto_beautify_slide: 自动美化单页幻灯片
 * - wps_ppt_beautify_all_slides: 批量美化所有幻灯片
 * - wps_ppt_create_kpi_cards: 创建KPI指标卡片
 * - wps_ppt_create_styled_table: 创建带样式的表格
 * - wps_ppt_add_title_decoration: 添加标题装饰元素
 * - wps_ppt_add_page_indicator: 添加页码指示器
 * - wps_ppt_set_background_gradient: 设置渐变背景
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 应用配色方案
 * 为幻灯片或全文档应用统一配色方案
 */
export declare const applyColorSchemeDefinition: ToolDefinition;
export declare const applyColorSchemeHandler: ToolHandler;
/**
 * 自动美化单页幻灯片
 * 智能分析并优化单页幻灯片的排版、字体、间距和配色
 */
export declare const autoBeautifySlideDefinition: ToolDefinition;
export declare const autoBeautifySlideHandler: ToolHandler;
/**
 * 批量美化所有幻灯片
 * 对演示文稿中所有幻灯片应用统一美化风格
 */
export declare const beautifyAllSlidesDefinition: ToolDefinition;
export declare const beautifyAllSlidesHandler: ToolHandler;
/**
 * 创建KPI指标卡片
 * 在幻灯片上生成数据指标卡片，适合展示关键业绩数据
 */
export declare const createKpiCardsDefinition: ToolDefinition;
export declare const createKpiCardsHandler: ToolHandler;
/**
 * 创建带样式的表格
 * 插入预设样式的精美表格
 */
export declare const createStyledTableDefinition: ToolDefinition;
export declare const createStyledTableHandler: ToolHandler;
/**
 * 添加标题装饰元素
 * 为幻灯片标题添加装饰性元素（下划线、色块、图标等）
 */
export declare const addTitleDecorationDefinition: ToolDefinition;
export declare const addTitleDecorationHandler: ToolHandler;
/**
 * 添加页码指示器
 * 在幻灯片上添加页码/进度指示器
 */
export declare const addPageIndicatorDefinition: ToolDefinition;
export declare const addPageIndicatorHandler: ToolHandler;
/**
 * 设置渐变背景
 * 为幻灯片设置渐变色背景
 */
export declare const setBackgroundGradientDefinition: ToolDefinition;
export declare const setBackgroundGradientHandler: ToolHandler;
/**
 * 导出所有高级美化相关的Tools
 */
export declare const beautifyAdvancedTools: RegisteredTool[];
export default beautifyAdvancedTools;
//# sourceMappingURL=beautify-advanced.d.ts.map
/**
 * Input: 透视表工具参数
 * Output: 透视表创建/更新结果
 * Pos: Excel 透视表工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel透视表Tools - 数据分析模块
 * 透视表是数据汇总和分析的核心功能
 *
 * 包含：
 * - wps_excel_create_pivot_table: 创建透视表
 * - wps_excel_update_pivot_table: 更新透视表配置
 *
 * @date 2026-01-24
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 创建透视表的Tool定义
 * 这个工具是数据分析的核心入口，用户说"帮我做个数据汇总"就靠它了
 */
export declare const createPivotTableDefinition: ToolDefinition;
/**
 * 创建透视表的Handler
 */
export declare const createPivotTableHandler: ToolHandler;
/**
 * 更新透视表配置的Tool定义
 * 透视表创建之后经常需要调整，这个工具负责更新配置
 */
export declare const updatePivotTableDefinition: ToolDefinition;
/**
 * 更新透视表配置的Handler
 */
export declare const updatePivotTableHandler: ToolHandler;
/**
 * 导出所有透视表相关的Tools
 * 整整齐齐，方便注册
 */
export declare const pivotTools: RegisteredTool[];
export default pivotTools;
//# sourceMappingURL=pivot.d.ts.map
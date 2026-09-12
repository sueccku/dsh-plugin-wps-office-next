/**
 * Input: 行列操作工具参数
 * Output: 行列插入/删除/隐藏/分组结果
 * Pos: Excel 行列操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel行列操作Tools - 行列管理模块
 * 处理行列的插入、删除、隐藏、显示、分组等操作
 *
 * 包含：
 * - wps_excel_insert_rows: 插入多行
 * - wps_excel_insert_columns: 插入多列
 * - wps_excel_delete_rows: 删除多行
 * - wps_excel_delete_columns: 删除多列
 * - wps_excel_hide_rows: 隐藏行
 * - wps_excel_show_rows: 显示行
 * - wps_excel_show_columns: 显示列
 * - wps_excel_group_rows: 分组行
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 插入多行
 */
export declare const insertRowsDefinition: ToolDefinition;
export declare const insertRowsHandler: ToolHandler;
/**
 * 插入多列
 */
export declare const insertColumnsDefinition: ToolDefinition;
export declare const insertColumnsHandler: ToolHandler;
/**
 * 删除多行
 */
export declare const deleteRowsDefinition: ToolDefinition;
export declare const deleteRowsHandler: ToolHandler;
/**
 * 删除多列
 */
export declare const deleteColumnsDefinition: ToolDefinition;
export declare const deleteColumnsHandler: ToolHandler;
/**
 * 隐藏行
 */
export declare const hideRowsDefinition: ToolDefinition;
export declare const hideRowsHandler: ToolHandler;
/**
 * 显示行
 */
export declare const showRowsDefinition: ToolDefinition;
export declare const showRowsHandler: ToolHandler;
/**
 * 显示列
 */
export declare const showColumnsDefinition: ToolDefinition;
export declare const showColumnsHandler: ToolHandler;
/**
 * 分组行
 */
export declare const groupRowsDefinition: ToolDefinition;
export declare const groupRowsHandler: ToolHandler;
/**
 * 导出所有行列操作相关的Tools
 */
export declare const rowColumnTools: RegisteredTool[];
export default rowColumnTools;
//# sourceMappingURL=row-column.d.ts.map
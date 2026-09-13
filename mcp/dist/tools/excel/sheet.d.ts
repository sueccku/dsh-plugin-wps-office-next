/**
 * Input: 工作表管理工具参数
 * Output: 工作表操作结果
 * Pos: Excel 工作表管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel工作表管理Tools
 * 工作表的创建、删除、重命名、复制、移动、切换等操作
 *
 * 包含：
 * - wps_excel_create_sheet: 创建新工作表
 * - wps_excel_delete_sheet: 删除指定工作表
 * - wps_excel_rename_sheet: 重命名工作表
 * - wps_excel_copy_sheet: 复制工作表
 * - wps_excel_get_sheet_list: 获取工作表列表
 * - wps_excel_switch_sheet: 切换工作表
 * - wps_excel_move_sheet: 移动工作表
 * - wps_excel_get_selection: 获取当前选中区域
 * - wps_excel_delete_row: 删除指定行
 * - wps_excel_insert_column: 插入列
 * - wps_excel_delete_column: 删除指定列
 * - wps_excel_freeze_panes: 冻结/取消冻结窗格
 * - wps_excel_auto_fill: 自动填充单元格区域
 * - wps_excel_set_named_range: 设置命名范围
 * - wps_excel_hide_column: 隐藏/显示列
 * - wps_excel_auto_sum: 对指定列/行自动求和
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 创建新工作表
 */
export declare const createSheetDefinition: ToolDefinition;
export declare const createSheetHandler: ToolHandler;
/**
 * 删除指定工作表
 */
export declare const deleteSheetDefinition: ToolDefinition;
export declare const deleteSheetHandler: ToolHandler;
/**
 * 重命名工作表
 */
export declare const renameSheetDefinition: ToolDefinition;
export declare const renameSheetHandler: ToolHandler;
/**
 * 复制工作表
 */
export declare const copySheetDefinition: ToolDefinition;
export declare const copySheetHandler: ToolHandler;
/**
 * 获取工作表列表
 */
export declare const getSheetListDefinition: ToolDefinition;
export declare const getSheetListHandler: ToolHandler;
/**
 * 切换到指定工作表
 */
export declare const switchSheetDefinition: ToolDefinition;
export declare const switchSheetHandler: ToolHandler;
/**
 * 移动工作表到指定位置
 */
export declare const moveSheetDefinition: ToolDefinition;
export declare const moveSheetHandler: ToolHandler;
/**
 * 获取当前选中区域信息
 */
export declare const getSelectionDefinition: ToolDefinition;
export declare const getSelectionHandler: ToolHandler;
/**
 * 冻结/取消冻结窗格
 */
export declare const freezePanesDefinition: ToolDefinition;
export declare const freezePanesHandler: ToolHandler;
/**
 * 设置命名范围
 */
export declare const setNamedRangeDefinition: ToolDefinition;
export declare const setNamedRangeHandler: ToolHandler;
/**
 * 隐藏/显示列
 */
export declare const hideColumnDefinition: ToolDefinition;
export declare const hideColumnHandler: ToolHandler;
/**
 * 对指定列/行自动求和
 */
export declare const autoSumDefinition: ToolDefinition;
export declare const autoSumHandler: ToolHandler;
/**
 * 导出所有工作表管理相关的Tools
 */
export declare const sheetTools: RegisteredTool[];
export default sheetTools;
//# sourceMappingURL=sheet.d.ts.map
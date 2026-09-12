/**
 * Input: 批注/保护/公式/图片/超链接工具参数
 * Output: 批注删除/获取、取消保护、锁定单元格、数组公式、图片插入、超链接设置结果
 * Pos: Excel 批注保护及扩展工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel批注保护Tools - 批注、保护与扩展功能模块
 * 处理批注管理、工作表取消保护、单元格锁定、数组公式、图片插入、超链接等操作
 *
 * 包含：
 * - wps_excel_delete_cell_comment: 删除单元格批注
 * - wps_excel_get_cell_comments: 获取单元格批注
 * - wps_excel_unprotect_sheet: 取消保护工作表
 * - wps_excel_lock_cells: 锁定/解锁单元格
 * - wps_excel_set_array_formula: 设置数组公式
 * - wps_excel_insert_excel_image: 插入图片
 * - wps_excel_set_hyperlink: 设置超链接
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 删除单元格批注
 */
export declare const deleteCellCommentDefinition: ToolDefinition;
export declare const deleteCellCommentHandler: ToolHandler;
/**
 * 获取单元格批注
 */
export declare const getCellCommentsDefinition: ToolDefinition;
export declare const getCellCommentsHandler: ToolHandler;
/**
 * 取消保护工作表
 */
export declare const unprotectSheetDefinition: ToolDefinition;
export declare const unprotectSheetHandler: ToolHandler;
/**
 * 锁定/解锁单元格
 */
export declare const lockCellsDefinition: ToolDefinition;
export declare const lockCellsHandler: ToolHandler;
/**
 * 设置数组公式
 */
export declare const setArrayFormulaDefinition: ToolDefinition;
export declare const setArrayFormulaHandler: ToolHandler;
/**
 * 插入图片到Excel
 */
export declare const insertExcelImageDefinition: ToolDefinition;
export declare const insertExcelImageHandler: ToolHandler;
/**
 * 设置超链接
 */
export declare const setHyperlinkDefinition: ToolDefinition;
export declare const setHyperlinkHandler: ToolHandler;
/**
 * 导出所有批注保护相关的Tools
 */
export declare const commentProtectTools: RegisteredTool[];
export default commentProtectTools;
//# sourceMappingURL=comment-protect.d.ts.map
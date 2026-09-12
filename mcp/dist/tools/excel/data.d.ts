/**
 * Input: 数据处理工具参数
 * Output: 读写/清洗结果
 * Pos: Excel 数据处理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel数据处理Tools - 数据读写与清洗模块
 * 处理数据的读取、写入、去重、去空格、格式统一等操作
 *
 * 包含：
 * - wps_excel_read_range: 读取指定范围数据
 * - wps_excel_write_range: 写入数据到指定范围
 * - wps_excel_clean_data: 数据清洗（核心功能）
 * - wps_excel_remove_duplicates: 删除重复行
 * - wps_excel_sort_range: 对选定区域排序
 * - wps_excel_find_replace: 查找并替换内容
 * - wps_excel_insert_row: 插入行
 * - wps_excel_add_comment: 给单元格添加批注
 * - wps_excel_protect_sheet: 保护/取消保护工作表
 * - wps_excel_set_conditional_format: 设置条件格式
 * - wps_excel_protect_workbook: 保护/取消保护工作簿
 * - wps_excel_set_zoom: 设置工作表缩放比例
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 读取指定范围的单元格数据
 */
export declare const readRangeDefinition: ToolDefinition;
export declare const readRangeHandler: ToolHandler;
/**
 * 向指定范围写入数据
 * 批量写入数据，比一个个单元格设置快多了
 */
export declare const writeRangeDefinition: ToolDefinition;
export declare const writeRangeHandler: ToolHandler;
/**
 * 数据清洗工具
 * 一键处理脏数据，支持多种清洗操作组合
 */
export declare const cleanDataDefinition: ToolDefinition;
export declare const cleanDataHandler: ToolHandler;
/**
 * 删除重复行
 * 单独拎出来，因为这个功能用得太多了
 */
export declare const removeDuplicatesDefinition: ToolDefinition;
export declare const removeDuplicatesHandler: ToolHandler;
/**
 * 对选定区域排序
 */
export declare const sortRangeDefinition: ToolDefinition;
export declare const sortRangeHandler: ToolHandler;
/**
 * 查找并替换内容
 */
export declare const findReplaceDefinition: ToolDefinition;
export declare const findReplaceHandler: ToolHandler;
/**
 * 插入行
 */
export declare const insertRowDefinition: ToolDefinition;
export declare const insertRowHandler: ToolHandler;
/**
 * 给单元格添加批注
 */
export declare const addCommentDefinition: ToolDefinition;
export declare const addCommentHandler: ToolHandler;
/**
 * 保护/取消保护工作表
 */
export declare const protectSheetDefinition: ToolDefinition;
export declare const protectSheetHandler: ToolHandler;
/**
 * 设置条件格式
 */
export declare const setConditionalFormatDefinition: ToolDefinition;
export declare const setConditionalFormatHandler: ToolHandler;
/**
 * 保护/取消保护工作簿
 */
export declare const protectWorkbookDefinition: ToolDefinition;
export declare const protectWorkbookHandler: ToolHandler;
/**
 * 设置工作表缩放比例
 */
export declare const setZoomDefinition: ToolDefinition;
export declare const setZoomHandler: ToolHandler;
/**
 * 导出所有数据处理相关的Tools
 */
export declare const dataTools: RegisteredTool[];
export default dataTools;
//# sourceMappingURL=data.d.ts.map
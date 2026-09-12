/**
 * Input: 格式化工具参数
 * Output: 格式化操作结果
 * Pos: Excel 基础格式化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel格式化Tools - 单元格格式/边框/合并/列宽/行高
 *
 * 包含：
 * - wps_excel_set_cell_format: 设置单元格格式（字体、颜色、背景等）
 * - wps_excel_set_cell_style: 应用预定义样式到单元格
 * - wps_excel_set_border: 设置单元格边框样式
 * - wps_excel_set_number_format: 设置单元格数字格式
 * - wps_excel_merge_cells: 合并指定范围的单元格
 * - wps_excel_unmerge_cells: 拆分合并的单元格
 * - wps_excel_set_column_width: 设置列宽
 * - wps_excel_set_row_height: 设置行高
 * - wps_excel_hide_row: 隐藏/显示行
 * - wps_excel_set_data_validation: 设置数据验证规则
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
export declare const setCellFormatDefinition: ToolDefinition;
export declare const setCellFormatHandler: ToolHandler;
export declare const setCellStyleDefinition: ToolDefinition;
export declare const setCellStyleHandler: ToolHandler;
export declare const setBorderDefinition: ToolDefinition;
export declare const setBorderHandler: ToolHandler;
export declare const setNumberFormatDefinition: ToolDefinition;
export declare const setNumberFormatHandler: ToolHandler;
export declare const mergeCellsDefinition: ToolDefinition;
export declare const mergeCellsHandler: ToolHandler;
export declare const unmergeCellsDefinition: ToolDefinition;
export declare const unmergeCellsHandler: ToolHandler;
export declare const setColumnWidthDefinition: ToolDefinition;
export declare const setColumnWidthHandler: ToolHandler;
export declare const setRowHeightDefinition: ToolDefinition;
export declare const setRowHeightHandler: ToolHandler;
export declare const hideRowDefinition: ToolDefinition;
export declare const hideRowHandler: ToolHandler;
export declare const setDataValidationDefinition: ToolDefinition;
export declare const setDataValidationHandler: ToolHandler;
export declare const excelFormatTools: RegisteredTool[];
export default excelFormatTools;
//# sourceMappingURL=format.d.ts.map
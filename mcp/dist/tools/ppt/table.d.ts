/**
 * Input: PPT 表格操作参数
 * Output: 表格操作结果
 * Pos: PPT 表格工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT表格Tools - 表格管理模块
 * 处理表格的插入、单元格读写、样式设置等操作
 *
 * 包含：
 * - wps_ppt_insert_table: 在幻灯片中插入表格
 * - wps_ppt_set_table_cell: 设置表格单元格文本
 * - wps_ppt_get_table_cell: 获取表格单元格文本
 * - wps_ppt_set_table_style: 设置表格整体样式
 * - wps_ppt_set_table_cell_style: 设置表格单元格样式
 * - wps_ppt_set_table_row_style: 设置表格行样式
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 在幻灯片中插入表格
 * 支持指定行列数和位置
 */
export declare const insertPptTableDefinition: ToolDefinition;
export declare const insertPptTableHandler: ToolHandler;
/**
 * 设置表格单元格文本
 * 修改指定单元格的内容
 */
export declare const setPptTableCellDefinition: ToolDefinition;
export declare const setPptTableCellHandler: ToolHandler;
/**
 * 获取表格单元格文本
 * 读取指定单元格的内容
 */
export declare const getPptTableCellDefinition: ToolDefinition;
export declare const getPptTableCellHandler: ToolHandler;
/**
 * 导出所有表格相关的Tools
 */
/** 表格外观：位置尺寸 + 单元格/整行/整表样式（P4 由三个工具合并而来） */
export declare const setPptTableFormatDefinition: ToolDefinition;
export declare const setPptTableFormatHandler: ToolHandler;
export declare const tableTools: RegisteredTool[];
export default tableTools;
//# sourceMappingURL=table.d.ts.map
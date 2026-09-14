import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/** 书签 */
export declare const getBookmarksDefinition: ToolDefinition;
export declare const getBookmarksHandler: ToolHandler;
/** 批注（读） */
export declare const getCommentsDefinition: ToolDefinition;
export declare const getCommentsHandler: ToolHandler;
/** 文档统计 */
export declare const getDocumentStatsDefinition: ToolDefinition;
export declare const getDocumentStatsHandler: ToolHandler;
/** 超链接 */
export declare const insertHyperlinkDefinition: ToolDefinition;
export declare const insertHyperlinkHandler: ToolHandler;
/** 表格：列表 */
export declare const getTablesDefinition: ToolDefinition;
export declare const getTablesHandler: ToolHandler;
/** 表格：读数据 */
export declare const getTableDataDefinition: ToolDefinition;
export declare const getTableDataHandler: ToolHandler;
/** 表格：写单元格 */
export declare const setTableCellDefinition: ToolDefinition;
export declare const setTableCellHandler: ToolHandler;
/** 表格：增行/增列 */
export declare const addTableLinesDefinition: ToolDefinition;
export declare const addTableLinesHandler: ToolHandler;
/** 表格：删行/删列 */
export declare const deleteTableLineDefinition: ToolDefinition;
export declare const deleteTableLineHandler: ToolHandler;
/** 表格：合并单元格 */
export declare const mergeTableCellsDefinition: ToolDefinition;
export declare const mergeTableCellsHandler: ToolHandler;
/** 表格：拆分单元格 */
export declare const splitTableCellDefinition: ToolDefinition;
export declare const splitTableCellHandler: ToolHandler;
/** 表格：样式 */
export declare const setTableFormatDefinition: ToolDefinition;
export declare const setTableFormatHandler: ToolHandler;
/** 表格 → 文本 */
export declare const convertTableToTextDefinition: ToolDefinition;
export declare const convertTableToTextHandler: ToolHandler;
export declare const wordDeepTools: RegisteredTool[];
export default wordDeepTools;
//# sourceMappingURL=deep.d.ts.map
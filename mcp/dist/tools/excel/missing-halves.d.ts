import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 工作表结构与已用范围：让调用方不必猜 A1:Z200
 */
export declare const getSheetInfoDefinition: ToolDefinition;
export declare const getSheetInfoHandler: ToolHandler;
export declare const autoFitDefinition: ToolDefinition;
export declare const autoFitHandler: ToolHandler;
export declare const autoFitColumnsDefinition: ToolDefinition;
export declare const autoFitColumnsHandler: ToolHandler;
export declare const autoFitRowsDefinition: ToolDefinition;
export declare const autoFitRowsHandler: ToolHandler;
/** 自动换行 */
export declare const setWrapTextDefinition: ToolDefinition;
export declare const setWrapTextHandler: ToolHandler;
/** 查找并给位置（不改动内容）——与 find_replace 的分工是：那个替换，这个只报告在哪 */
export declare const findInSheetDefinition: ToolDefinition;
export declare const findInSheetHandler: ToolHandler;
/** 命名范围：读与删（此前只有写） */
export declare const getNamedRangesDefinition: ToolDefinition;
export declare const getNamedRangesHandler: ToolHandler;
export declare const deleteNamedRangeDefinition: ToolDefinition;
export declare const deleteNamedRangeHandler: ToolHandler;
export declare const missingHalfTools: RegisteredTool[];
export default missingHalfTools;
//# sourceMappingURL=missing-halves.d.ts.map
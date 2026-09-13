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
/** 格式刷：只搬格式，不动值与公式 */
export declare const copyFormatDefinition: ToolDefinition;
export declare const copyFormatHandler: ToolHandler;
/** 清除格式：内容留着，外观回到默认 */
export declare const clearFormatsDefinition: ToolDefinition;
export declare const clearFormatsHandler: ToolHandler;
export declare const getConditionalFormatsDefinition: ToolDefinition;
export declare const getConditionalFormatsHandler: ToolHandler;
export declare const removeConditionalFormatDefinition: ToolDefinition;
export declare const removeConditionalFormatHandler: ToolHandler;
export declare const getDataValidationsDefinition: ToolDefinition;
export declare const getDataValidationsHandler: ToolHandler;
export declare const removeDataValidationDefinition: ToolDefinition;
export declare const removeDataValidationHandler: ToolHandler;
/** 外部链接刷新 */
export declare const refreshLinksDefinition: ToolDefinition;
export declare const refreshLinksHandler: ToolHandler;
/** 合并计算：把多块来源区域汇总到目标区域 */
export declare const consolidateDefinition: ToolDefinition;
export declare const consolidateHandler: ToolHandler;
/** 强制重算 */
export declare const calculateDefinition: ToolDefinition;
export declare const calculateHandler: ToolHandler;
/** 列分组（分级显示）——行分组早有 wps_excel_group_rows */
export declare const groupColumnsDefinition: ToolDefinition;
export declare const groupColumnsHandler: ToolHandler;
export declare const missingHalfTools: RegisteredTool[];
export default missingHalfTools;
//# sourceMappingURL=missing-halves.d.ts.map
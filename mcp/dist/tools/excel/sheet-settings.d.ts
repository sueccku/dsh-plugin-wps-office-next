import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/** 读全部页面/打印/外观设置 */
export declare const getSheetSettingsDefinition: ToolDefinition;
export declare const getSheetSettingsHandler: ToolHandler;
/** 页面设置 */
export declare const setSheetPageSetupDefinition: ToolDefinition;
export declare const setSheetPageSetupHandler: ToolHandler;
/** 打印标题 */
export declare const setSheetPrintTitlesDefinition: ToolDefinition;
export declare const setSheetPrintTitlesHandler: ToolHandler;
/** 页眉页脚 */
export declare const setSheetHeaderFooterDefinition: ToolDefinition;
export declare const setSheetHeaderFooterHandler: ToolHandler;
/** 工作表外观：隐藏与标签色 */
export declare const setSheetAppearanceDefinition: ToolDefinition;
export declare const setSheetAppearanceHandler: ToolHandler;
/** 分级显示 */
export declare const setOutlineLevelsDefinition: ToolDefinition;
export declare const setOutlineLevelsHandler: ToolHandler;
/** 清除分页符 */
export declare const resetPageBreaksDefinition: ToolDefinition;
export declare const resetPageBreaksHandler: ToolHandler;
/** 公式审计：引用追踪 */
export declare const getFormulaAuditDefinition: ToolDefinition;
export declare const getFormulaAuditHandler: ToolHandler;
export declare const sheetSettingsTools: RegisteredTool[];
export default sheetSettingsTools;
//# sourceMappingURL=sheet-settings.d.ts.map
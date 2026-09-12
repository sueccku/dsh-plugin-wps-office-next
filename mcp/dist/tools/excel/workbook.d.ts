/**
 * Input: 工作簿管理工具参数
 * Output: 工作簿操作结果
 * Pos: Excel 工作簿管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 打开指定路径的Excel工作簿
 */
export declare const openWorkbookDefinition: ToolDefinition;
export declare const openWorkbookHandler: ToolHandler;
/**
 * 获取所有已打开的工作簿列表
 */
export declare const getOpenWorkbooksDefinition: ToolDefinition;
export declare const getOpenWorkbooksHandler: ToolHandler;
/**
 * 切换到指定名称的工作簿
 */
export declare const switchWorkbookDefinition: ToolDefinition;
export declare const switchWorkbookHandler: ToolHandler;
/**
 * 关闭指定工作簿
 */
export declare const closeWorkbookDefinition: ToolDefinition;
export declare const closeWorkbookHandler: ToolHandler;
/**
 * 新建空白工作簿
 */
export declare const createWorkbookDefinition: ToolDefinition;
export declare const createWorkbookHandler: ToolHandler;
/**
 * 获取指定单元格的值
 */
export declare const getCellValueDefinition: ToolDefinition;
export declare const getCellValueHandler: ToolHandler;
/**
 * 设置指定单元格的值
 */
export declare const setCellValueDefinition: ToolDefinition;
export declare const setCellValueHandler: ToolHandler;
/**
 * 获取指定单元格的公式
 */
export declare const getFormulaDefinition: ToolDefinition;
export declare const getFormulaHandler: ToolHandler;
/**
 * 获取单元格详细信息
 */
export declare const getCellInfoDefinition: ToolDefinition;
export declare const getCellInfoHandler: ToolHandler;
/**
 * 清除指定范围
 */
export declare const clearRangeDefinition: ToolDefinition;
export declare const clearRangeHandler: ToolHandler;
/**
 * 导出所有工作簿管理相关的Tools
 */
export declare const workbookTools: RegisteredTool[];
export default workbookTools;
//# sourceMappingURL=workbook.d.ts.map
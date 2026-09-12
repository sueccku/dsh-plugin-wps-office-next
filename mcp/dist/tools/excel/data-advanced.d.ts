/**
 * Input: 高级数据处理工具参数
 * Output: 数据处理结果
 * Pos: Excel 高级数据处理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 自动筛选
 */
export declare const autoFilterDefinition: ToolDefinition;
export declare const autoFilterHandler: ToolHandler;
/**
 * 复制范围
 */
export declare const copyRangeDefinition: ToolDefinition;
export declare const copyRangeHandler: ToolHandler;
/**
 * 粘贴范围
 */
export declare const pasteRangeDefinition: ToolDefinition;
export declare const pasteRangeHandler: ToolHandler;
/**
 * 填充序列
 */
export declare const fillSeriesDefinition: ToolDefinition;
export declare const fillSeriesHandler: ToolHandler;
/**
 * 转置数据
 */
export declare const transposeDefinition: ToolDefinition;
export declare const transposeHandler: ToolHandler;
/**
 * 分列
 */
export declare const textToColumnsDefinition: ToolDefinition;
export declare const textToColumnsHandler: ToolHandler;
/**
 * 分类汇总
 */
export declare const subtotalDefinition: ToolDefinition;
export declare const subtotalHandler: ToolHandler;
/**
 * 导出所有高级数据处理相关的Tools
 */
export declare const dataAdvancedTools: RegisteredTool[];
export default dataAdvancedTools;
//# sourceMappingURL=data-advanced.d.ts.map
/**
 * Input: 公式类工具参数
 * Output: 公式计算与诊断结果
 * Pos: Excel 公式工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel公式相关Tools - 公式管理模块
 * 解决用户"公式不会写"痛点的核心工具集
 *
 * 包含：
 * - wps_excel_set_formula: 设置公式到指定单元格
 * - wps_excel_generate_formula: 根据自然语言生成公式（核心功能）
 * - wps_excel_diagnose_formula: 诊断公式错误，分析原因并提供修复建议
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 设置公式到指定单元格
 * 公式功能的执行端，负责将生成的公式写入单元格
 */
export declare const setFormulaDefinition: ToolDefinition;
export declare const setFormulaHandler: ToolHandler;
/**
 * 根据自然语言生成Excel公式
 * 核心功能：获取工作表上下文，辅助AI理解表结构后生成公式
 */
export declare const generateFormulaDefinition: ToolDefinition;
export declare const generateFormulaHandler: ToolHandler;
/**
 * 诊断公式错误
 * 分析#REF!、#N/A、#VALUE!等错误类型，给出原因和修复建议
 */
export declare const diagnoseFormulaDefinition: ToolDefinition;
export declare const diagnoseFormulaHandler: ToolHandler;
/**
 * 导出所有公式相关的Tools
 */
export declare const evaluateFormulaDefinition: ToolDefinition;
export declare const evaluateFormulaHandler: (args: Record<string, unknown>) => Promise<{
    id: string;
    success: boolean;
    content: {
        type: "text";
        text: string;
    }[];
}>;
export declare const setPrintAreaDefinition: ToolDefinition;
export declare const setPrintAreaHandler: (args: Record<string, unknown>) => Promise<{
    id: string;
    success: boolean;
    content: {
        type: "text";
        text: string;
    }[];
}>;
export declare const zoomDefinition: ToolDefinition;
export declare const zoomHandler: (args: Record<string, unknown>) => Promise<{
    id: string;
    success: boolean;
    content: {
        type: "text";
        text: string;
    }[];
}>;
export declare const formulaTools: RegisteredTool[];
export default formulaTools;
//# sourceMappingURL=formula.d.ts.map
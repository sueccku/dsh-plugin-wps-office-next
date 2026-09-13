/**
 * Input: Tool 定义集合
 * Output: Tool 注册数组
 * Pos: MCP Tools 总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tools总入口 - MCP工具汇总注册模块
 * 整合Excel、Word、PPT、Common的所有Tools
 *
 * 使用方法：
 * import { allTools } from './tools';
 * toolRegistry.registerAll(allTools);
 *
 * 或者按需导入：
 * import { excelTools, wordTools, pptTools, commonTools } from './tools';
 */
import { RegisteredTool } from '../types/tools';
/**
 * 所有 MCP Tools 集合：注册 209 个 = Excel 75 + Word 32 + PPT 88 + Common 9 + 逃生舱 1 + 门面 4。
 *
 * 这里不再逐工具枚举（以前那份清单每加一个工具就会过期）：要清单看生成物 skills/<app>/reference.md，
 * 或让模型用 wps_help {app:"excel"} 查。
 * 18 个已合并的重复名不注册、只做派发期别名（见 tools/deprecated.ts）；
 * 内置工具只剩 wps_execute_method 一个逃生舱（其余 11 个重复/缓存/连接检查类已在 P0 清理中删除）。
 */
export declare const allTools: RegisteredTool[];
export { excelTools } from './excel';
export { wordTools } from './word';
export { pptTools } from './ppt';
export { commonTools } from './common';
export { formulaTools, dataTools, setFormulaDefinition, setFormulaHandler, generateFormulaDefinition, generateFormulaHandler, diagnoseFormulaDefinition, diagnoseFormulaHandler, readRangeDefinition, readRangeHandler, writeRangeDefinition, writeRangeHandler, cleanDataDefinition, cleanDataHandler, removeDuplicatesDefinition, removeDuplicatesHandler, } from './excel';
export { formatTools, contentTools, proofreadTools, applyStyleDefinition, applyStyleHandler, setFontDefinition, setFontHandler, generateTocDefinition, generateTocHandler, insertTextDefinition, insertTextHandler, findReplaceDefinition, findReplaceHandler, } from './word';
export { slideTools, addSlideDefinition, addSlideHandler, beautifyDefinition, beautifyHandler, unifyFontDefinition, unifyFontHandler, } from './ppt';
export { convertTools, convertToPdfDefinition, convertToPdfHandler, convertFormatDefinition, convertFormatHandler, getAppTypeByExtension, getFormatCode, } from './common';
/**
 * 获取所有Tool的数量
 */
export declare const getToolCount: () => number;
/**
 * 获取按应用分类的Tool数量
 */
export declare const getToolCountByApp: () => {
    excel: number;
    word: number;
    ppt: number;
    common: number;
};
export default allTools;
//# sourceMappingURL=index.d.ts.map
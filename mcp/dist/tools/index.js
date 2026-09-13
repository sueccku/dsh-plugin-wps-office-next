"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.getToolCountByApp = exports.getToolCount = exports.getFormatCode = exports.getAppTypeByExtension = exports.convertFormatHandler = exports.convertFormatDefinition = exports.convertToPdfHandler = exports.convertToPdfDefinition = exports.convertTools = exports.unifyFontHandler = exports.unifyFontDefinition = exports.beautifyHandler = exports.beautifyDefinition = exports.addSlideHandler = exports.addSlideDefinition = exports.slideTools = exports.findReplaceHandler = exports.findReplaceDefinition = exports.insertTextHandler = exports.insertTextDefinition = exports.generateTocHandler = exports.generateTocDefinition = exports.setFontHandler = exports.setFontDefinition = exports.applyStyleHandler = exports.applyStyleDefinition = exports.proofreadTools = exports.contentTools = exports.formatTools = exports.removeDuplicatesHandler = exports.removeDuplicatesDefinition = exports.cleanDataHandler = exports.cleanDataDefinition = exports.writeRangeHandler = exports.writeRangeDefinition = exports.readRangeHandler = exports.readRangeDefinition = exports.diagnoseFormulaHandler = exports.diagnoseFormulaDefinition = exports.generateFormulaHandler = exports.generateFormulaDefinition = exports.setFormulaHandler = exports.setFormulaDefinition = exports.dataTools = exports.formulaTools = exports.commonTools = exports.pptTools = exports.wordTools = exports.excelTools = exports.allTools = void 0;
const excel_1 = require("./excel");
const word_1 = require("./word");
const ppt_1 = require("./ppt");
const common_1 = require("./common");
/**
 * 所有 MCP Tools 集合：注册 209 个 = Excel 75 + Word 32 + PPT 88 + Common 9 + 逃生舱 1 + 门面 4。
 *
 * 这里不再逐工具枚举（以前那份清单每加一个工具就会过期）：要清单看生成物 skills/<app>/reference.md，
 * 或让模型用 wps_help {app:"excel"} 查。
 * 18 个已合并的重复名不注册、只做派发期别名（见 tools/deprecated.ts）；
 * 内置工具只剩 wps_execute_method 一个逃生舱（其余 11 个重复/缓存/连接检查类已在 P0 清理中删除）。
 */
exports.allTools = [
    ...excel_1.excelTools,
    ...word_1.wordTools,
    ...ppt_1.pptTools,
    ...common_1.commonTools,
];
// 按应用类型分别导出
var excel_2 = require("./excel");
Object.defineProperty(exports, "excelTools", { enumerable: true, get: function () { return excel_2.excelTools; } });
var word_2 = require("./word");
Object.defineProperty(exports, "wordTools", { enumerable: true, get: function () { return word_2.wordTools; } });
var ppt_2 = require("./ppt");
Object.defineProperty(exports, "pptTools", { enumerable: true, get: function () { return ppt_2.pptTools; } });
var common_2 = require("./common");
Object.defineProperty(exports, "commonTools", { enumerable: true, get: function () { return common_2.commonTools; } });
// Excel相关导出
var excel_3 = require("./excel");
Object.defineProperty(exports, "formulaTools", { enumerable: true, get: function () { return excel_3.formulaTools; } });
Object.defineProperty(exports, "dataTools", { enumerable: true, get: function () { return excel_3.dataTools; } });
Object.defineProperty(exports, "setFormulaDefinition", { enumerable: true, get: function () { return excel_3.setFormulaDefinition; } });
Object.defineProperty(exports, "setFormulaHandler", { enumerable: true, get: function () { return excel_3.setFormulaHandler; } });
Object.defineProperty(exports, "generateFormulaDefinition", { enumerable: true, get: function () { return excel_3.generateFormulaDefinition; } });
Object.defineProperty(exports, "generateFormulaHandler", { enumerable: true, get: function () { return excel_3.generateFormulaHandler; } });
Object.defineProperty(exports, "diagnoseFormulaDefinition", { enumerable: true, get: function () { return excel_3.diagnoseFormulaDefinition; } });
Object.defineProperty(exports, "diagnoseFormulaHandler", { enumerable: true, get: function () { return excel_3.diagnoseFormulaHandler; } });
Object.defineProperty(exports, "readRangeDefinition", { enumerable: true, get: function () { return excel_3.readRangeDefinition; } });
Object.defineProperty(exports, "readRangeHandler", { enumerable: true, get: function () { return excel_3.readRangeHandler; } });
Object.defineProperty(exports, "writeRangeDefinition", { enumerable: true, get: function () { return excel_3.writeRangeDefinition; } });
Object.defineProperty(exports, "writeRangeHandler", { enumerable: true, get: function () { return excel_3.writeRangeHandler; } });
Object.defineProperty(exports, "cleanDataDefinition", { enumerable: true, get: function () { return excel_3.cleanDataDefinition; } });
Object.defineProperty(exports, "cleanDataHandler", { enumerable: true, get: function () { return excel_3.cleanDataHandler; } });
Object.defineProperty(exports, "removeDuplicatesDefinition", { enumerable: true, get: function () { return excel_3.removeDuplicatesDefinition; } });
Object.defineProperty(exports, "removeDuplicatesHandler", { enumerable: true, get: function () { return excel_3.removeDuplicatesHandler; } });
// Word相关导出
var word_3 = require("./word");
Object.defineProperty(exports, "formatTools", { enumerable: true, get: function () { return word_3.formatTools; } });
Object.defineProperty(exports, "contentTools", { enumerable: true, get: function () { return word_3.contentTools; } });
Object.defineProperty(exports, "proofreadTools", { enumerable: true, get: function () { return word_3.proofreadTools; } });
Object.defineProperty(exports, "applyStyleDefinition", { enumerable: true, get: function () { return word_3.applyStyleDefinition; } });
Object.defineProperty(exports, "applyStyleHandler", { enumerable: true, get: function () { return word_3.applyStyleHandler; } });
Object.defineProperty(exports, "setFontDefinition", { enumerable: true, get: function () { return word_3.setFontDefinition; } });
Object.defineProperty(exports, "setFontHandler", { enumerable: true, get: function () { return word_3.setFontHandler; } });
Object.defineProperty(exports, "generateTocDefinition", { enumerable: true, get: function () { return word_3.generateTocDefinition; } });
Object.defineProperty(exports, "generateTocHandler", { enumerable: true, get: function () { return word_3.generateTocHandler; } });
Object.defineProperty(exports, "insertTextDefinition", { enumerable: true, get: function () { return word_3.insertTextDefinition; } });
Object.defineProperty(exports, "insertTextHandler", { enumerable: true, get: function () { return word_3.insertTextHandler; } });
Object.defineProperty(exports, "findReplaceDefinition", { enumerable: true, get: function () { return word_3.findReplaceDefinition; } });
Object.defineProperty(exports, "findReplaceHandler", { enumerable: true, get: function () { return word_3.findReplaceHandler; } });
// PPT相关导出
var ppt_3 = require("./ppt");
Object.defineProperty(exports, "slideTools", { enumerable: true, get: function () { return ppt_3.slideTools; } });
Object.defineProperty(exports, "addSlideDefinition", { enumerable: true, get: function () { return ppt_3.addSlideDefinition; } });
Object.defineProperty(exports, "addSlideHandler", { enumerable: true, get: function () { return ppt_3.addSlideHandler; } });
Object.defineProperty(exports, "beautifyDefinition", { enumerable: true, get: function () { return ppt_3.beautifyDefinition; } });
Object.defineProperty(exports, "beautifyHandler", { enumerable: true, get: function () { return ppt_3.beautifyHandler; } });
Object.defineProperty(exports, "unifyFontDefinition", { enumerable: true, get: function () { return ppt_3.unifyFontDefinition; } });
Object.defineProperty(exports, "unifyFontHandler", { enumerable: true, get: function () { return ppt_3.unifyFontHandler; } });
// Common相关导出
var common_3 = require("./common");
Object.defineProperty(exports, "convertTools", { enumerable: true, get: function () { return common_3.convertTools; } });
Object.defineProperty(exports, "convertToPdfDefinition", { enumerable: true, get: function () { return common_3.convertToPdfDefinition; } });
Object.defineProperty(exports, "convertToPdfHandler", { enumerable: true, get: function () { return common_3.convertToPdfHandler; } });
Object.defineProperty(exports, "convertFormatDefinition", { enumerable: true, get: function () { return common_3.convertFormatDefinition; } });
Object.defineProperty(exports, "convertFormatHandler", { enumerable: true, get: function () { return common_3.convertFormatHandler; } });
Object.defineProperty(exports, "getAppTypeByExtension", { enumerable: true, get: function () { return common_3.getAppTypeByExtension; } });
Object.defineProperty(exports, "getFormatCode", { enumerable: true, get: function () { return common_3.getFormatCode; } });
/**
 * 获取所有Tool的数量
 */
const getToolCount = () => exports.allTools.length;
exports.getToolCount = getToolCount;
/**
 * 获取按应用分类的Tool数量
 */
const getToolCountByApp = () => ({
    excel: excel_1.excelTools.length,
    word: word_1.wordTools.length,
    ppt: ppt_1.pptTools.length,
    common: common_1.commonTools.length,
});
exports.getToolCountByApp = getToolCountByApp;
exports.default = exports.allTools;
//# sourceMappingURL=index.js.map
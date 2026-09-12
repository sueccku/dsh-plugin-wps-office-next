"use strict";
/**
 * Input: 通用工具定义
 * Output: 通用工具注册数组
 * Pos: Common Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Common Tools入口 - 通用工具汇总模块
 * 整合跨应用的通用Tools
 *
 * 导出所有通用的Tools，Word/Excel/PPT均可使用
 * 主要包括文档转换、格式互转等跨应用功能
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.setSelectedTextHandler = exports.setSelectedTextDefinition = exports.getSelectedTextHandler = exports.getSelectedTextDefinition = exports.getAppInfoHandler = exports.getAppInfoDefinition = exports.wireCheckHandler = exports.wireCheckDefinition = exports.pingHandler = exports.pingDefinition = exports.saveAsHandler = exports.saveAsDefinition = exports.saveHandler = exports.saveDefinition = exports.getFormatCode = exports.getAppTypeByExtension = exports.convertFormatHandler = exports.convertFormatDefinition = exports.convertToPdfHandler = exports.convertToPdfDefinition = exports.generalTools = exports.convertTools = exports.commonTools = void 0;
const convert_1 = require("./convert");
const general_1 = require("./general");
/**
 * 所有通用的Tools
 * 包含：
 * - 转换Tools: convert_to_pdf, convert_format
 * - 通用Tools: save, save_as, ping, wire_check, get_app_info, get_selected_text, set_selected_text
 *
 * 后续可继续添加更多通用功能，如：
 * - 文档比较
 * - 批量处理
 * - 云端同步
 */
exports.commonTools = [
    ...convert_1.convertTools,
    ...general_1.generalTools,
];
// 分别导出，方便按需使用
var convert_2 = require("./convert");
Object.defineProperty(exports, "convertTools", { enumerable: true, get: function () { return convert_2.convertTools; } });
var general_2 = require("./general");
Object.defineProperty(exports, "generalTools", { enumerable: true, get: function () { return general_2.generalTools; } });
// 导出单独的定义和处理器，方便测试
var convert_3 = require("./convert");
Object.defineProperty(exports, "convertToPdfDefinition", { enumerable: true, get: function () { return convert_3.convertToPdfDefinition; } });
Object.defineProperty(exports, "convertToPdfHandler", { enumerable: true, get: function () { return convert_3.convertToPdfHandler; } });
Object.defineProperty(exports, "convertFormatDefinition", { enumerable: true, get: function () { return convert_3.convertFormatDefinition; } });
Object.defineProperty(exports, "convertFormatHandler", { enumerable: true, get: function () { return convert_3.convertFormatHandler; } });
Object.defineProperty(exports, "getAppTypeByExtension", { enumerable: true, get: function () { return convert_3.getAppTypeByExtension; } });
Object.defineProperty(exports, "getFormatCode", { enumerable: true, get: function () { return convert_3.getFormatCode; } });
// 导出通用操作工具的定义和处理器
var general_3 = require("./general");
Object.defineProperty(exports, "saveDefinition", { enumerable: true, get: function () { return general_3.saveDefinition; } });
Object.defineProperty(exports, "saveHandler", { enumerable: true, get: function () { return general_3.saveHandler; } });
Object.defineProperty(exports, "saveAsDefinition", { enumerable: true, get: function () { return general_3.saveAsDefinition; } });
Object.defineProperty(exports, "saveAsHandler", { enumerable: true, get: function () { return general_3.saveAsHandler; } });
Object.defineProperty(exports, "pingDefinition", { enumerable: true, get: function () { return general_3.pingDefinition; } });
Object.defineProperty(exports, "pingHandler", { enumerable: true, get: function () { return general_3.pingHandler; } });
Object.defineProperty(exports, "wireCheckDefinition", { enumerable: true, get: function () { return general_3.wireCheckDefinition; } });
Object.defineProperty(exports, "wireCheckHandler", { enumerable: true, get: function () { return general_3.wireCheckHandler; } });
Object.defineProperty(exports, "getAppInfoDefinition", { enumerable: true, get: function () { return general_3.getAppInfoDefinition; } });
Object.defineProperty(exports, "getAppInfoHandler", { enumerable: true, get: function () { return general_3.getAppInfoHandler; } });
Object.defineProperty(exports, "getSelectedTextDefinition", { enumerable: true, get: function () { return general_3.getSelectedTextDefinition; } });
Object.defineProperty(exports, "getSelectedTextHandler", { enumerable: true, get: function () { return general_3.getSelectedTextHandler; } });
Object.defineProperty(exports, "setSelectedTextDefinition", { enumerable: true, get: function () { return general_3.setSelectedTextDefinition; } });
Object.defineProperty(exports, "setSelectedTextHandler", { enumerable: true, get: function () { return general_3.setSelectedTextHandler; } });
exports.default = exports.commonTools;
//# sourceMappingURL=index.js.map
"use strict";
/**
 * Input: MCP Tool 类型与Schema
 * Output: Tool 类型定义
 * Pos: MCP Tool 类型定义模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tool类型定义 - 老王出品，必属精品
 * 定义MCP Tool的所有相关类型，别tm乱改
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.ToolCategory = void 0;
/**
 * Tool分类枚举 - WPS三件套对应的分类
 */
var ToolCategory;
(function (ToolCategory) {
    /** 文档操作 - Word那一套 */
    ToolCategory["DOCUMENT"] = "document";
    /** 表格操作 - Excel那一套 */
    ToolCategory["SPREADSHEET"] = "spreadsheet";
    /** 演示操作 - PPT那一套 */
    ToolCategory["PRESENTATION"] = "presentation";
    /** 通用操作 - 跨应用的通用功能 */
    ToolCategory["COMMON"] = "common";
})(ToolCategory || (exports.ToolCategory = ToolCategory = {}));
//# sourceMappingURL=tools.js.map
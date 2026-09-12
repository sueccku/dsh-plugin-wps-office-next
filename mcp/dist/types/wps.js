"use strict";
/**
 * Input: WPS 通信类型定义
 * Output: WPS 类型与枚举
 * Pos: WPS 客户端类型定义模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS相关类型定义 - 老王的WPS客户端类型
 * 跟WPS加载项通信用的类型，别tm搞混了
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.WpsAppType = void 0;
/**
 * WPS应用类型枚举
 */
var WpsAppType;
(function (WpsAppType) {
    /** WPS文字 - 对标Word */
    WpsAppType["WRITER"] = "wps";
    /** WPS表格 - 对标Excel */
    WpsAppType["SPREADSHEET"] = "et";
    /** WPS演示 - 对标PPT */
    WpsAppType["PRESENTATION"] = "wpp";
})(WpsAppType || (exports.WpsAppType = WpsAppType = {}));
//# sourceMappingURL=wps.js.map
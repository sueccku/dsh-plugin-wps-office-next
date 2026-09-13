"use strict";
/**
 * Input: 无
 * Output: 共享解析函数读取的键：函数名 -> 该函数会从 '$p' 里取哪些键
 * Pos: P1 契约真源的一部分。生成器从这里的键表推导每个 action 真正接受的键（宿主脚本据此
 *      在派发前拒绝多余键）。原先它是 build-host-actions.ps1 里的一张手写表。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P1 状态。
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.helperKeys = void 0;
/** 共享解析函数读取的键。改动这里等于改动守卫接受哪些键，务必配合测试。 */
exports.helperKeys = {
    "Resolve-Worksheet": [
        "sheet",
        "name",
        "oldName"
    ],
    "Get-TargetPres": [
        "presentationName"
    ],
    "Get-WorksheetByParam": [
        "sheet"
    ],
    "Get-RowRefList": [
        "rows",
        "row",
        "count",
        "startRow",
        "endRow"
    ],
    "Get-ColumnRefList": [
        "columns",
        "column",
        "count",
        "startColumn",
        "endColumn"
    ],
    "Resolve-PictureIndex": [
        "name",
        "shapeName",
        "shapeIndex",
        "imageIndex"
    ],
    "Resolve-TextBoxIndex": [
        "name",
        "shapeName",
        "shapeIndex",
        "textboxIndex"
    ],
    "Get-PptBackgroundSpec": [
        "background",
        "type",
        "color",
        "colors",
        "imagePath"
    ]
};
//# sourceMappingURL=bridge-helpers.js.map
"use strict";
/**
 * Input: 桥回传的破坏性操作前置统计（RangeImpact）
 * Output: 一句可复述的中文影响摘要
 * Pos: S3 数据安全审计的共享格式化层。桥里的破坏性动作（clearRange / clearFormats /
 *      deleteRows / deleteColumns / deleteSheet …）在动手前会算好「将要失去什么」，这里把它压成
 *      模型能直接对用户复述的一句话。统计是尽力而为的：字段缺失只是摘要变短，绝不影响动作结果。
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.impactText = impactText;
/** 把影响统计压成一句放在动作结果后面的中文括号。没有可用信息时返回空串。 */
function impactText(impact) {
    if (!impact)
        return '';
    const parts = [];
    if (typeof impact.nonEmpty === 'number')
        parts.push(`原有 ${impact.nonEmpty} 个非空单元格`);
    else if (typeof impact.cells === 'number')
        parts.push(`范围共 ${impact.cells} 个单元格`);
    if (Array.isArray(impact.preview) && impact.preview.length > 0) {
        parts.push(`示例内容：${impact.preview.slice(0, 5).join(' / ')}`);
    }
    return parts.length > 0 ? `（${parts.join('；')}）` : '';
}
//# sourceMappingURL=impact.js.map
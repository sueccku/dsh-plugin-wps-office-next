"use strict";
/**
 * Input: 桥回传的破坏性操作前置统计（Impact）
 * Output: 一句可复述的中文影响摘要
 * Pos: S3 数据安全审计的共享格式化层。桥里的破坏性动作（clearRange / deleteRows / deleteChart /
 *      deleteSlide …）在动手前会算好「将要失去什么」，这里把它压成模型能直接对用户复述的一句话。
 *      统计是尽力而为的：字段缺失只是摘要变短，绝不影响动作结果。
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.impactText = impactText;
const KIND_LABELS = {
    conditionalFormat: '条件格式规则',
    dataValidation: '数据验证规则',
    listRow: '表行',
    listObject: '表格对象',
    pageBreaks: '手动分页符',
    pivotTable: '透视表',
    sparkline: '迷你图',
    chart: '图表',
    namedRange: '命名范围',
    slide: '幻灯片',
    shape: '形状',
    textbox: '文本框',
    image: '图片',
    animation: '动画',
    comment: '批注',
};
/** 把影响统计压成一句放在动作结果后面的中文括号。没有可用信息时返回空串。 */
function impactText(impact) {
    if (!impact)
        return '';
    const label = impact.kind ? KIND_LABELS[impact.kind] || String(impact.kind) : '项';
    const parts = [];
    if (typeof impact.count === 'number' && impact.count > 0)
        parts.push(`共 ${impact.count} 个${label}`);
    else if (impact.name)
        parts.push(`${label}「${impact.name}」`);
    if (typeof impact.nonEmpty === 'number')
        parts.push(`原有 ${impact.nonEmpty} 个非空单元格`);
    else if (typeof impact.cells === 'number' && impact.cells > 1)
        parts.push(`范围共 ${impact.cells} 个单元格`);
    if (typeof impact.rows === 'number' && impact.rows > 0)
        parts.push(`${impact.rows} 行`);
    if (impact.detail)
        parts.push(String(impact.detail));
    else if (impact.address && impact.nonEmpty == null && impact.cells == null)
        parts.push(`原引用 ${impact.address}`);
    if (Array.isArray(impact.preview) && impact.preview.length > 0) {
        parts.push(`示例内容：${impact.preview.slice(0, 5).join(' / ')}`);
    }
    return parts.length > 0 ? `（${parts.join('；')}）` : '';
}
//# sourceMappingURL=impact.js.map
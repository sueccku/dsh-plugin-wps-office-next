/**
 * Input: 桥回传的破坏性操作前置统计（RangeImpact）
 * Output: 一句可复述的中文影响摘要
 * Pos: S3 数据安全审计的共享格式化层。桥里的破坏性动作（clearRange / clearFormats /
 *      deleteRows / deleteColumns / deleteSheet …）在动手前会算好「将要失去什么」，这里把它压成
 *      模型能直接对用户复述的一句话。统计是尽力而为的：字段缺失只是摘要变短，绝不影响动作结果。
 */
/** 桥在破坏性动作执行前回传的影响统计。 */
export interface RangeImpact {
    /** 范围地址，如 $A$1:$C$10；读不到时为 null。 */
    address?: string | null;
    /** 范围内单元格总数。 */
    cells?: number | null;
    /** 范围内的非空单元格数（真正会被抹掉的内容量）。 */
    nonEmpty?: number | null;
    /** 被影响内容的前几项示例，最多 8 项。 */
    preview?: string[] | null;
}
/** 把影响统计压成一句放在动作结果后面的中文括号。没有可用信息时返回空串。 */
export declare function impactText(impact: RangeImpact | undefined | null): string;
//# sourceMappingURL=impact.d.ts.map
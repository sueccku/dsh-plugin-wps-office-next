/**
 * Input: 桥回传的破坏性操作前置统计（Impact）
 * Output: 一句可复述的中文影响摘要
 * Pos: S3 数据安全审计的共享格式化层。桥里的破坏性动作（clearRange / deleteRows / deleteChart /
 *      deleteSlide …）在动手前会算好「将要失去什么」，这里把它压成模型能直接对用户复述的一句话。
 *      统计是尽力而为的：字段缺失只是摘要变短，绝不影响动作结果。
 */
/** 桥在破坏性动作执行前回传的影响统计。范围类与对象类字段可以混用。 */
export interface RangeImpact {
    /** 范围地址，或命名范围的 RefersTo 引用。 */
    address?: string | null;
    /** 范围内单元格总数。 */
    cells?: number | null;
    /** 范围内的非空单元格数（真正会被抹掉的内容量）。 */
    nonEmpty?: number | null;
    /** 被影响内容的前几项示例，最多 8 项。 */
    preview?: string[] | null;
    /** 对象名（表 / 图表 / 透视表 / 形状 / 命名范围…）。 */
    name?: string | null;
    /** 对象类型键，用于生成中文量词，见 KIND_LABELS。 */
    kind?: string | null;
    /** 被删除 / 覆盖的对象数量。 */
    count?: number | null;
    /** 行数。 */
    rows?: number | null;
    /** 一句补充说明（例如「清除前：横向 2 个、纵向 1 个」）。 */
    detail?: string | null;
}
/** 把影响统计压成一句放在动作结果后面的中文括号。没有可用信息时返回空串。 */
export declare function impactText(impact: RangeImpact | undefined | null): string;
//# sourceMappingURL=impact.d.ts.map
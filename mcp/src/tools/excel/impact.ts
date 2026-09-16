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

const KIND_LABELS: Record<string, string> = {
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
export function impactText(impact: RangeImpact | undefined | null): string {
  if (!impact) return '';
  const label = impact.kind ? KIND_LABELS[impact.kind] || String(impact.kind) : '项';
  const parts: string[] = [];
  if (typeof impact.count === 'number' && impact.count > 0) parts.push(`共 ${impact.count} 个${label}`);
  else if (impact.name) parts.push(`${label}「${impact.name}」`);
  if (typeof impact.nonEmpty === 'number') parts.push(`原有 ${impact.nonEmpty} 个非空单元格`);
  else if (typeof impact.cells === 'number' && impact.cells > 1) parts.push(`范围共 ${impact.cells} 个单元格`);
  if (typeof impact.rows === 'number' && impact.rows > 0) parts.push(`${impact.rows} 行`);
  if (impact.detail) parts.push(String(impact.detail));
  else if (impact.address && impact.nonEmpty == null && impact.cells == null) parts.push(`原引用 ${impact.address}`);
  if (Array.isArray(impact.preview) && impact.preview.length > 0) {
    parts.push(`示例内容：${impact.preview.slice(0, 5).join(' / ')}`);
  }
  return parts.length > 0 ? `（${parts.join('；')}）` : '';
}

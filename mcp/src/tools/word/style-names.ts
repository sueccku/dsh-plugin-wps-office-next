/**
 * Input: 模型给的样式名（中文/英文/带不带空格都可能）
 * Output: 归一化后的样式名 + 是否来自英文别名表
 * Pos: FIXES 88（L2）。WPS 12.1 中文版的样式表里**没有英文名**：`Styles.Item("Heading 1")`、
 *      `Range.Style = "Heading 1"`、`Style.NameInternational` 全部失败/为空（实测 478 个样式，
 *      NameInternational 一律空串）。而工具描述里恰恰把 "Heading 1" 当示例，模型照着调用就吃 E_FAIL。
 *      这里做力所能及的翻译 + 空白归一化；真正"样式存不存在"的裁决仍然在 WPS 那边（见桥的报错改进）。
 *      一旦我被修改，请更新我的头部注释。
 */

/** 英文别名 -> 中文 WPS 里真实存在的样式名。只收常用内置样式，不做全量翻译。 */
const ENGLISH_ALIASES: Record<string, string> = {
  normal: '正文',
  bodytext: '正文文本',
  title: '标题',
  subtitle: '副标题',
  quote: '引用',
  intensequote: '明显引用',
  emphasis: '强调',
  intenseemphasis: '明显强调',
  caption: '题注',
  listparagraph: '列表段落',
  noSpacing: '无间隔',
};

/** "heading 1" / "heading1" / "Heading  1" -> 1..9；不是标题则返回 undefined。 */
function headingLevel(key: string): number | undefined {
  const m = /^heading\s*([1-9])$/.exec(key);
  return m ? Number(m[1]) : undefined;
}

/** 去掉首尾空白、把连续空白（含全角空格）压成一个普通空格。WPS 的样式名对空格敏感：
 *  "标题  1"（两个空格）实测也会 E_FAIL。 */
export function normalizeStyleName(raw: string): string {
  return String(raw ?? '')
    .replace(/[\u3000\u00a0]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export interface ResolvedStyleName {
  /** 交给 WPS 的名字。 */
  name: string;
  /** true 表示模型给的是英文别名，已翻译。 */
  translated: boolean;
}

/**
 * 把模型给的样式名解析成 WPS 能接受的名字。
 * 1) 英文别名（Heading 1 / Normal / Title ...）-> 中文内置名；
 * 2) "标题1"（少空格）-> "标题 1"；
 * 3) 其余原样返回（含用户自定义样式名，不擅自改写）。
 */
export function resolveStyleName(raw: string): ResolvedStyleName {
  const normalized = normalizeStyleName(raw);
  if (!normalized) return { name: normalized, translated: false };

  const key = normalized.toLowerCase();
  const level = headingLevel(key);
  if (level !== undefined) return { name: `标题 ${level}`, translated: true };
  // 英文别名按「去掉空格 + 小写」查：Body Text / bodytext / BodyText 都指向同一条。
  const alias = ENGLISH_ALIASES[key] ?? ENGLISH_ALIASES[key.replace(/\s+/g, '')];
  if (alias) return { name: alias, translated: true };

  // 中文名少空格：标题1 -> 标题 1（只在"标题/副标题 + 数字"这种精确形状上补，避免误改用户样式名）
  const cjk = /^(标题|副标题|列表|索引|目录|引文目录|注释标题)([1-9])$/.exec(normalized);
  if (cjk) return { name: `${cjk[1]} ${cjk[2]}`, translated: false };

  return { name: normalized, translated: false };
}

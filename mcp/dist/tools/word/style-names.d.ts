/**
 * Input: 模型给的样式名（中文/英文/带不带空格都可能）
 * Output: 归一化后的样式名 + 是否来自英文别名表
 * Pos: FIXES 88（L2）。WPS 12.1 中文版的样式表里**没有英文名**：`Styles.Item("Heading 1")`、
 *      `Range.Style = "Heading 1"`、`Style.NameInternational` 全部失败/为空（实测 478 个样式，
 *      NameInternational 一律空串）。而工具描述里恰恰把 "Heading 1" 当示例，模型照着调用就吃 E_FAIL。
 *      这里做力所能及的翻译 + 空白归一化；真正"样式存不存在"的裁决仍然在 WPS 那边（见桥的报错改进）。
 *      一旦我被修改，请更新我的头部注释。
 */
/** 去掉首尾空白、把连续空白（含全角空格）压成一个普通空格。WPS 的样式名对空格敏感：
 *  "标题  1"（两个空格）实测也会 E_FAIL。 */
export declare function normalizeStyleName(raw: string): string;
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
export declare function resolveStyleName(raw: string): ResolvedStyleName;
//# sourceMappingURL=style-names.d.ts.map
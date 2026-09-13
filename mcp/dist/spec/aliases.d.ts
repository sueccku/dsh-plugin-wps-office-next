/**
 * Input: 无
 * Output: 桥侧的参数兼容声明：公开参数名 -> 该 action 实际读取的键；以及嵌套容器名
 * Pos: P1 契约真源的一部分。生成器从这里产出 spec/param-aliases.json 与 spec/param-containers.json，
 *      宿主脚本再据此在派发前改名与展平——取代了 build-host-actions.ps1 里那两张手写表。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P1 状态。
 */
/** 公开参数名 -> 桥读取的键。含工具 schema 没有声明、但历史上被接受的兼容拼写。 */
export declare const paramAliases: Record<string, Record<string, string>>;
/** 需要展平的嵌套容器名（其属性会被并到该 action 的扁平键集合上）。 */
export declare const paramContainers: Record<string, string[]>;
//# sourceMappingURL=aliases.d.ts.map
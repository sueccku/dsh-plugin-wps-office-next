/**
 * Input: 无
 * Output: 桥侧的参数兼容声明：公开参数名 -> 该 action 实际读取的键；以及嵌套容器名
 * Pos: P1 契约真源的一部分。生成器从这里产出 spec/param-aliases.json 与 spec/param-containers.json，
 *      宿主脚本再据此在派发前改名与展平——取代了 build-host-actions.ps1 里那两张手写表。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P1 状态。
 */
/** 公开参数名 -> 桥读取的键。含工具 schema 没有声明、但历史上被接受的兼容拼写。 */
export declare const paramAliases: Record<string, Record<string, string>>;
/**
 * 参数键无法静态提取的 action：生成器会拒绝为它们写键表，也就意味着守卫对它们不生效。
 * 这是一份**有账的例外**：只有在这里声明过的 action 才允许被跳过；出现未声明的动态 action，
 * 生成器直接失败——过去它是静默跳过的，于是「改了但没生效」和「改好了」在结果上无法区分。
 */
export declare const dynamicParamActions: Record<string, string>;
/** 需要展平的嵌套容器名（其属性会被并到该 action 的扁平键集合上）。 */
export declare const paramContainers: Record<string, string[]>;
//# sourceMappingURL=aliases.d.ts.map
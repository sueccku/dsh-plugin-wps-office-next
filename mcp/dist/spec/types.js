"use strict";
/**
 * Input: 无（纯类型与校验工具）
 * Output: 操作规格（operation spec）的类型与构造器
 * Pos: P1 契约真源。spec 是唯一真源：工具 schema、宿主键表、技能参考表、广告集都由它生成。
 *      本文件只放声明式数据形状，生成逻辑在 scripts/gen-tool-surface.mjs。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P1 状态。
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.APP_ORDER = void 0;
exports.op = op;
/** 声明式构造器：唯一作用是让 tsc 检查每个条目的字段与类型。 */
function op(spec) {
    if (!spec.tool.startsWith('wps_'))
        throw new Error('tool name must start with wps_: ' + spec.tool);
    // 只有一条硬规则：声明为 bridge 就必须有 action。其余（opaque 不许带参数、必填必须存在）
    // 是 P1-4 的收敛目标，由生成器在迁移完成后强制，草稿阶段只做记录。
    if (spec.engine !== 'opaque' && spec.engine !== 'local' && !spec.action) {
        throw new Error('a bridge operation needs an action: ' + spec.tool);
    }
    return spec;
}
/** 按 app 分组，供生成器与文档使用。 */
exports.APP_ORDER = ['excel', 'word', 'ppt', 'common'];
//# sourceMappingURL=types.js.map
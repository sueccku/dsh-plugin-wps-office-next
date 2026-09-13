/**
 * Input: 无（纯类型与校验工具）
 * Output: 操作规格（operation spec）的类型与构造器
 * Pos: P1 契约真源。spec 是唯一真源：工具 schema、宿主键表、技能参考表、广告集都由它生成。
 *      本文件只放声明式数据形状，生成逻辑在 scripts/gen-tool-surface.mjs。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P1 状态。
 */

/** 参数在桥里被读取时的类型；用于生成 JSON Schema 与文档。 */
export type ParamType = 'string' | 'number' | 'boolean' | 'object' | 'array' | 'array2d';

/** 一个参数的规格。字段名即桥读取的键（公开名 = 桥名，见 P1-4）。 */
export interface ParamSpec {
  type: ParamType;
  /** 中文说明，进 schema 的 description */
  description?: string;
  required?: boolean;
  /** 枚举取值（会生成 JSON Schema 的 enum） */
  enum?: string[];
  /** 数组元素类型（type 为 array / array2d 时） */
  items?: { type: string; description?: string };
  /** 默认值，进 schema 的 default */
  default?: unknown;
  /**
   * 原始 JSON Schema 片段。只有用上面几个友好字段表达不完整时才出现（嵌套数组、一个参数多种类型等）；
   * 生成器优先用它，保证 P1 的「逐字节复现」验收成立。它的条数是迁移进度的度量，收敛目标为 0。
   */
  schema?: Record<string, unknown>;
}

/** 操作类别：决定广告策略与安全提示。 */
export type OperationEffect = 'read' | 'write' | 'delete' | 'lifecycle' | 'export';

/** 一条操作规格。 */
export interface OperationSpec {
  /** 模型看到的工具名，形如 wps_excel_write_range */
  tool: string;
  /** 桥里的 action 名；纯 JS 工具为 null */
  action: string | null;
  app: 'excel' | 'word' | 'ppt' | 'common';
  /** 工具描述（进 schema，也是 wps_help 的文案） */
  summary: string;
  params: Record<string, ParamSpec>;
  /**
   * 原样的 required 列表：**必须逐字保留**（老 surface 并不一致——有的工具写 `required: []`，
   * 有的干脆没有这个键），否则复现验收不会通过。undefined = 原 schema 没有这个键。
   */
  required?: string[];
  effect: OperationEffect;
  /** 是否进 standard 档广告位 */
  advertised: boolean;
  /**
   * 纯 JS 工具（不经桥）或桥里读不出参数的 action，在这里写明原因；
   * 引擎为 'opaque' 的操作不允许携带 params（生成器会拒绝），杜绝"静默跳过校验"的老问题。
   */
  engine?: 'bridge' | 'local' | 'opaque';
  /** 参数名与桥键不一致时的映射；P1-4 的目标是让它归零。 */
  aliases?: Record<string, string>;
  /** 工具发送的嵌套对象名；其属性会被展平成桥读取的键（P1-4 目标：归零）。 */
  containers?: string[];
}

/** 废弃名 → 规范工具：派发期解析，不占注册位。 */
export interface DeprecatedAlias {
  tool: string;
  canonical: string;
  reason: string;
  paramMap?: Record<string, string>;
}

/** 声明式构造器：唯一作用是让 tsc 检查每个条目的字段与类型。 */
export function op(spec: OperationSpec): OperationSpec {
  if (!spec.tool.startsWith('wps_')) throw new Error('tool name must start with wps_: ' + spec.tool);
  // 只有一条硬规则：声明为 bridge 就必须有 action。其余（opaque 不许带参数、必填必须存在）
  // 是 P1-4 的收敛目标，由生成器在迁移完成后强制，草稿阶段只做记录。
  if (spec.engine !== 'opaque' && spec.engine !== 'local' && !spec.action) {
    throw new Error('a bridge operation needs an action: ' + spec.tool);
  }
  return spec;
}

/** 按 app 分组，供生成器与文档使用。 */
export const APP_ORDER: OperationSpec['app'][] = ['excel', 'word', 'ppt', 'common'];

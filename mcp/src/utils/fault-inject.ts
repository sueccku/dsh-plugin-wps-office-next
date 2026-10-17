/**
 * Input: 原始 WPS_OFFICE_FAULT_INJECT 规格（环境变量）
 * Output: 每次匹配的 tool 调用至多一个故障决定
 * Pos: **仅测试**用的受控故障注入。WPS_OFFICE_FAULT_INJECT 未设置时完全惰性：
 *      不读参数、不改任何结果、不产生任何日志，正常路径行为与没有这个文件时逐字一致。
 *
 * 为什么要这个开关：`改进.md` §10.2 要求危险/异常场景用**受控故障注入**验证，
 * 而超时与"部分未生效"在真机上无法稳定复现（真超时会让整轮所有调用一起坏，测出来的是环境故障，
 * 不是模型行为）。这里让**指定的一步**确定性地失败一次，其余调用照常。
 *
 * 规格：`<pattern>:<mode>[:<count>]`，多条用 `;` 分隔。
 *   pattern  匹配 tool 名（正则）；`wps_call` 的调用还会用其内层 tool 名一起匹配。
 *   mode     `timeout`（执行完动作后如实报「状态未知」超时，动作其实已落盘）
 *            或 `shortfall`（**不执行**该动作，直接按"这一步未生效"回报，模拟部分失败）。
 *   count    匹配到几次后失效，默认 1（一次性）。
 * 例：`wps_word_insert_text:wps_call$:timeout:1`、`wps_word_set_font:shortfall:99`。
 *
 * 只在 test/skilleval 的场景 12（写入后超时）与 14（部分未生效）里由 rounds.mjs 注入；
 * 生产/日常使用**永远不要设置**这个变量。
 * 一旦我被修改，请更新我的头部注释。
 */
import { createChildLogger } from './logger';

const logger = createChildLogger('FaultInject');

export type FaultMode = 'timeout' | 'shortfall';

export interface FaultRule {
  /** 原始片段，仅用于日志与错误提示。 */
  raw: string;
  pattern: RegExp;
  mode: FaultMode;
  remaining: number;
}

export interface ParsedFaultSpec {
  rules: FaultRule[];
  /** 语法不对的片段：不静默忽略，交给调用方打日志（typo 会让注入不按预期发生）。 */
  invalid: string[];
}

/**
 * 解析规格。任何一段写坏都不抛异常：测试注入写错时宁可少注入（并留下 invalid），
 * 也不该让 MCP server 起不来 —— 起不来会把"实验条件错了"表现成"工具全坏了"。
 */
export function parseFaultSpec(raw: string | undefined | null): ParsedFaultSpec {
  const rules: FaultRule[] = [];
  const invalid: string[] = [];
  const text = (raw || '').trim();
  if (!text) return { rules, invalid };

  for (const item of text.split(';')) {
    const piece = item.trim();
    if (!piece) continue;
    const parts = piece.split(':');
    if (parts.length < 2 || parts.length > 3) {
      invalid.push(piece);
      continue;
    }
    const [pattern, mode, countText] = parts;
    if (mode !== 'timeout' && mode !== 'shortfall') {
      invalid.push(piece);
      continue;
    }
    const count = parts.length === 3 ? Number(countText) : 1;
    if (!Number.isInteger(count) || count < 1) {
      invalid.push(piece);
      continue;
    }
    let patternRe: RegExp;
    try {
      patternRe = new RegExp(pattern);
    } catch {
      invalid.push(piece);
      continue;
    }
    rules.push({ raw: piece, pattern: patternRe, mode, remaining: count });
  }
  return { rules, invalid };
}

/** 一次调用可能匹配的候选名：tool 本身、wps_call 的内层 tool，以及内层调用里的桥动作名。
 *
 * 为什么还要看 `args.args.method/action`：模型可以绕过直达工具，用
 * `wps_call {tool:"wps_execute_method", args:{method:"setFont"}}` 走同一个桥动作；
 * 只按 tool 名匹配时这类绕行会让注入"看起来生效、实际没挡住"，实验条件就不成立了。
 */
function candidateNames(toolName: string, args?: Record<string, unknown>): string[] {
  const names = [toolName];
  const inner = args && (args as { tool?: unknown }).tool;
  if (toolName === 'wps_call' && typeof inner === 'string' && inner.length > 0) {
    names.push(inner);
    const deep = args && (args as { args?: unknown }).args;
    if (deep && typeof deep === 'object') {
      for (const key of ['method', 'action']) {
        const value = (deep as Record<string, unknown>)[key];
        if (typeof value === 'string' && value.length > 0) names.push(value);
      }
    }
  }
  return names;
}

export class FaultInjector {
  private readonly rules: FaultRule[];
  readonly invalid: string[];

  constructor(raw: string | undefined | null = process.env.WPS_OFFICE_FAULT_INJECT) {
    const parsed = parseFaultSpec(raw);
    this.rules = parsed.rules;
    this.invalid = parsed.invalid;
    if (this.invalid.length > 0) {
      logger.warn('WPS_OFFICE_FAULT_INJECT 有片段被忽略（语法不对）', { invalid: this.invalid });
    }
    if (this.rules.length > 0) {
      logger.warn('WPS_OFFICE_FAULT_INJECT 已启用（仅测试用）', { rules: this.rules.map((rule) => rule.raw) });
    }
  }

  get active(): boolean {
    return this.rules.length > 0;
  }

  /** 命中则消耗一次并返回模式；没有命中返回 null。 */
  take(toolName: string, args?: Record<string, unknown>): FaultMode | null {
    if (this.rules.length === 0) return null;
    const names = candidateNames(toolName, args);
    for (const rule of this.rules) {
      if (rule.remaining > 0 && names.some((name) => rule.pattern.test(name))) {
        rule.remaining -= 1;
        return rule.mode;
      }
    }
    return null;
  }
}

/** 注入超时时回给模型的文案：与 com-host 真实超时同一套说法（「状态未知」是判断依据）。 */
export function injectedTimeoutText(actionLabel: string): string {
  return (
    '操作超时：' + actionLabel + ' 超过 60000ms 未返回。' +
    '状态未知：WPS 可能正被一个对话框阻塞（例如文档密码框），也可能已经无响应；' +
    '插件不会自动关闭 WPS，以免丢掉你未保存的内容。' +
    '下一步：切到 WPS 窗口处理弹出的对话框，然后重试' +
    '（后续调用会先用 15000ms 的短超时快速失败，成功一次即恢复）。'
  );
}

/** 注入部分失败时回给模型的短欠项：写成"这一步没生效"，而不是含糊的"部分失败"。 */
export function injectedShortfallText(actionLabel: string): string {
  return '这一步没有生效（' + actionLabel + '）：目标对象拒绝该格式，已生效的部分保持不变';
}

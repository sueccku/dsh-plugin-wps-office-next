/**
 * Input: active toolset mode and the registered tool catalog
 * Output: the tool definitions advertised to the model
 * Pos: Single place where the model-facing tool surface is curated.
 *      Every registered tool stays callable through the wps_call dispatcher; only the
 *      advertised subset costs tokens on every request.
 *      D4（2026-10-11）：广告面按真实走查的调用数据从 86 扩到 95（依据写在 STANDARD_TOOLS 的 D4 注释块里）。
 *      一旦我被修改，请更新我的头部注释。
 */
import { ToolDefinition } from '../types/tools';
/** Advertised surface presets. */
export type ToolsetMode = 'minimal' | 'standard' | 'full';
/** Mode used when WPS_OFFICE_TOOLSET is unset or invalid. */
export declare const DEFAULT_MODE: ToolsetMode;
/** Tools that are always advertised: the dispatcher facade. */
export declare const FACADE_TOOLS: string[];
/**
 * Curated high-frequency tools for the standard mode. Everything else stays reachable
 * through wps_call; wps_help lists and describes the hidden tail on demand.
 */
export declare const STANDARD_TOOLS: string[];
/** Resolve a raw environment value into a supported mode. */
export declare function resolveMode(raw?: string): ToolsetMode;
/** Names advertised for a mode, in catalog order. */
export declare function selectToolNames(mode: ToolsetMode, allNames: string[]): string[];
/** Tool definitions advertised for a mode. */
export declare function selectTools(mode: ToolsetMode, tools: ToolDefinition[]): ToolDefinition[];
/**
 * Shorten a tool description for the advertised surface. Full text stays available
 * through wps_help, so the per-request copy only needs to route the model correctly.
 */
export declare function compactDescription(text: string, max?: number): string;
/**
 * Split a wps_help free-text query into searchable tokens.
 * Whitespace and punctuation separate tokens; CJK runs stay whole because there is
 * nothing to split on, and the bigram path below covers those.
 */
export declare function tokenizeHelpQuery(query: string): string[];
/**
 * Relevance score for a wps_help query against one tool; 0 means "not a match".
 *
 * A model asks in whatever words fit its task ("close workbook 关闭", "新建 文档 create
 * new"), so a whole-phrase substring test answered "matched: 0" for queries a human would
 * call obvious. Score per token instead, and give CJK queries a bigram path: a space-less
 * query like "关闭工作簿" is not a substring of "关闭指定的Excel工作簿" but shares three of
 * its four bigrams.
 */
export declare function scoreToolHelpMatch(name: string, description: string, query: string, tokens: string[]): number;
/**
 * User words that do not appear in any tool name/description but are the normal way a Chinese
 * office user names the task. Expanded terms are scored in addition to the original query; this
 * keeps the original bigram behaviour while letting "题注/文献/母版" reach the right family.
 */
export declare const HELP_SYNONYMS: Record<string, string[]>;
/** Expand a free-text query with the synonym groups it contains (deduplicated, original first). */
export declare function expandHelpQuery(query: string): string[];
/** Score a query with synonym expansion; best match across the original and expanded terms. */
export declare function scoreToolHelpMatchExpanded(name: string, description: string, query: string, tokens: string[]): number;
/** Scenario entry: a natural-language task -> the tools + recipe that implement it. */
export interface HelpScenarioHint {
    keywords: string[];
    tools: string[];
    recipe: string;
    note?: string;
}
export declare const HELP_SCENARIOS: HelpScenarioHint[];
/** Return scenario hints matching a user query. */
export declare function matchHelpScenarios(query: string): HelpScenarioHint[];
//# sourceMappingURL=toolset.d.ts.map
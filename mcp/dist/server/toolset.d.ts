/**
 * Input: active toolset mode and the registered tool catalog
 * Output: the tool definitions advertised to the model
 * Pos: Single place where the model-facing tool surface is curated.
 *      Every registered tool stays callable through the wps_call dispatcher; only the
 *      advertised subset costs tokens on every request.
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
//# sourceMappingURL=toolset.d.ts.map
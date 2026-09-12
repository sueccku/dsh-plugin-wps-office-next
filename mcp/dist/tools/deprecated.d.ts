/**
 * Input: none
 * Output: the deprecated-tool map applied at server startup
 * Pos: Single place where duplicated upstream tools are collapsed into one canonical name.
 *      一旦我被修改，请更新我的头部注释。
 */
/** One merged-away tool name and the canonical tool that replaces it. */
export interface DeprecatedToolSpec {
    /** Tool name that callers should use instead. */
    canonical: string;
    /** Human-readable reason, surfaced by wps_help. */
    reason: string;
    /** Rename arguments whose spelling differs between the two tools. */
    paramMap?: Record<string, string>;
}
/**
 * Tools merged into a canonical name because they drive the same COM action through the
 * same interface. Each entry keeps the old name callable as a forwarder, so existing
 * prompts keep working, but it is hidden from wps_help and no longer a separate concept.
 */
export declare const DEPRECATED_TOOLS: Record<string, DeprecatedToolSpec>;
/** Names hidden from wps_help listings. */
export declare const DEPRECATED_NAMES: Set<string>;
/** Rename the keys given by a paramMap, leaving everything else untouched. */
export declare function renameArgs(args: Record<string, unknown>, map?: Record<string, string>): Record<string, unknown>;
//# sourceMappingURL=deprecated.d.ts.map
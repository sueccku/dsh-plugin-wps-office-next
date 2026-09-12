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
export const DEPRECATED_TOOLS: Record<string, DeprecatedToolSpec> = {
  wps_excel_zoom: {
    canonical: 'wps_excel_set_zoom',
    reason: '与 wps_excel_set_zoom 参数和语义完全相同',
  },
  wps_ppt_add_speaker_notes: {
    canonical: 'wps_ppt_set_slide_notes',
    reason: '与 wps_ppt_set_slide_notes 参数和语义完全相同',
  },
  wps_word_generate_doc_toc: {
    canonical: 'wps_word_generate_toc',
    reason: '参数是 wps_word_generate_toc 的子集',
  },
};

/** Names hidden from wps_help listings. */
export const DEPRECATED_NAMES: Set<string> = new Set(Object.keys(DEPRECATED_TOOLS));

/** Rename the keys given by a paramMap, leaving everything else untouched. */
export function renameArgs(
  args: Record<string, unknown>,
  map?: Record<string, string>
): Record<string, unknown> {
  if (!map) return args;
  const out: Record<string, unknown> = { ...args };
  for (const [from, to] of Object.entries(map)) {
    if (Object.prototype.hasOwnProperty.call(out, from)) {
      if (!(to in out)) out[to] = out[from];
      delete out[from];
    }
  }
  return out;
}

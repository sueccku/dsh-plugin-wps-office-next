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
export const DEFAULT_MODE: ToolsetMode = 'standard';

/** Tools that are always advertised: the dispatcher facade. */
export const FACADE_TOOLS: string[] = [
  'wps_status',
  'wps_help',
  'wps_call',
  'wps_batch',
];

/**
 * Curated high-frequency tools for the standard mode. Everything else stays reachable
 * through wps_call; wps_help lists and describes the hidden tail on demand.
 */
export const STANDARD_TOOLS: string[] = [
  // Excel
  'wps_excel_get_sheet_list',
  'wps_excel_read_range',
  'wps_excel_write_range',
  'wps_excel_set_cell_format',
  'wps_excel_set_number_format',
  'wps_excel_set_formula',
  'wps_excel_create_chart',
  'wps_excel_create_pivot_table',
  'wps_excel_find_replace',
  'wps_excel_get_open_workbooks',
  'wps_excel_open_workbook',
  'wps_excel_get_sheet_info',
  'wps_excel_auto_fit',
  'wps_excel_find_in_sheet',
  'wps_excel_get_named_ranges',
  'wps_excel_copy_format',
  'wps_excel_clear_formats',
  'wps_excel_calculate',
  // Word
  'wps_word_get_active_document',
  'wps_word_get_document_text',
  'wps_word_get_paragraphs',
  'wps_word_insert_text',
  'wps_word_find_replace',
  'wps_word_apply_style',
  'wps_word_set_font',
  'wps_word_set_paragraph',
  'wps_word_generate_toc',
  'wps_word_smart_fill_field',
  'wps_word_open_document',
  'wps_word_create_document',
  // Presentation
  'wps_ppt_get_slide_count',
  'wps_ppt_get_slide_info',
  'wps_ppt_get_shapes',
  'wps_ppt_set_shape_fill',
  'wps_ppt_set_shape_text',
  'wps_ppt_add_slide',
  'wps_ppt_set_slide_title',
  'wps_ppt_set_slide_content',
  'wps_ppt_add_textbox',
  'wps_ppt_insert_table',
  'wps_ppt_insert_ppt_image',
  'wps_ppt_export_slide_as_image',
  'wps_ppt_open_presentation',
  // Common and conversion
  'wps_common_save',
  'wps_common_save_as',
  'wps_convert_to_pdf',
  'wps_convert_format',
];

/** Resolve a raw environment value into a supported mode. */
export function resolveMode(raw?: string): ToolsetMode {
  const value = (raw || '').trim().toLowerCase();
  if (value === 'minimal' || value === 'standard' || value === 'full') return value;
  return DEFAULT_MODE;
}

/** Names advertised for a mode, in catalog order. */
export function selectToolNames(mode: ToolsetMode, allNames: string[]): string[] {
  if (mode === 'full') return [...allNames];
  const wanted = new Set(FACADE_TOOLS);
  if (mode === 'standard') for (const name of STANDARD_TOOLS) wanted.add(name);
  return allNames.filter((name) => wanted.has(name));
}

/** Tool definitions advertised for a mode. */
export function selectTools(mode: ToolsetMode, tools: ToolDefinition[]): ToolDefinition[] {
  const names = new Set(selectToolNames(mode, tools.map((t) => t.name)));
  return tools.filter((t) => names.has(t.name));
}

/**
 * Shorten a tool description for the advertised surface. Full text stays available
 * through wps_help, so the per-request copy only needs to route the model correctly.
 */
export function compactDescription(text: string, max = 110): string {
  const flat = (text || '').replace(/\s+/g, ' ').trim();
  if (flat.length <= max) return flat;
  const stop = flat.search(/[。.!?]/);
  if (stop >= 0 && stop + 1 <= max) return flat.slice(0, stop + 1);
  return flat.slice(0, max - 1) + '…';
}

/**
 * Split a wps_help free-text query into searchable tokens.
 * Whitespace and punctuation separate tokens; CJK runs stay whole because there is
 * nothing to split on, and the bigram path below covers those.
 */
export function tokenizeHelpQuery(query: string): string[] {
  return (query || '')
    .toLowerCase()
    .split(/[\s,，、;；:：/|+()（）\[\]【】"'“”]+/)
    .map((token) => token.trim())
    .filter((token) => token.length > 0);
}

/**
 * Relevance score for a wps_help query against one tool; 0 means "not a match".
 *
 * A model asks in whatever words fit its task ("close workbook 关闭", "新建 文档 create
 * new"), so a whole-phrase substring test answered "matched: 0" for queries a human would
 * call obvious. Score per token instead, and give CJK queries a bigram path: a space-less
 * query like "关闭工作簿" is not a substring of "关闭指定的Excel工作簿" but shares three of
 * its four bigrams.
 */
export function scoreToolHelpMatch(
  name: string,
  description: string,
  query: string,
  tokens: string[]
): number {
  const toolName = (name || '').toLowerCase();
  const toolDescription = (description || '').toLowerCase();
  const phrase = (query || '').trim().toLowerCase();
  if (!phrase) return 0;

  let score = 0;
  if (toolName.includes(phrase)) score += 12;
  else if (toolDescription.includes(phrase)) score += 10;

  for (const token of tokens) {
    if (token.length < 2) continue;
    if (toolName.includes(token)) score += 4;
    else if (toolDescription.includes(token)) score += 2;
  }

  const cjk = [...phrase].filter((char) => /[\u3400-\u9fff]/.test(char));
  if (cjk.length >= 2) {
    const bigrams = new Set<string>();
    for (let i = 0; i + 1 < cjk.length; i++) bigrams.add(cjk[i] + cjk[i + 1]);
    let hits = 0;
    for (const bigram of bigrams) {
      if (toolName.includes(bigram) || toolDescription.includes(bigram)) hits++;
    }
    const coverage = hits / bigrams.size;
    if (coverage >= 0.5) score += 6 * coverage;
  }

  return score;
}
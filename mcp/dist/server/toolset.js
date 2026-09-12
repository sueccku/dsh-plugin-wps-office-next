"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.STANDARD_TOOLS = exports.FACADE_TOOLS = exports.DEFAULT_MODE = void 0;
exports.resolveMode = resolveMode;
exports.selectToolNames = selectToolNames;
exports.selectTools = selectTools;
exports.compactDescription = compactDescription;
/** Mode used when WPS_OFFICE_TOOLSET is unset or invalid. */
exports.DEFAULT_MODE = 'standard';
/** Tools that are always advertised: the dispatcher facade. */
exports.FACADE_TOOLS = [
    'wps_status',
    'wps_help',
    'wps_call',
    'wps_batch',
];
/**
 * Curated high-frequency tools for the standard mode. Everything else stays reachable
 * through wps_call; wps_help lists and describes the hidden tail on demand.
 */
exports.STANDARD_TOOLS = [
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
function resolveMode(raw) {
    const value = (raw || '').trim().toLowerCase();
    if (value === 'minimal' || value === 'standard' || value === 'full')
        return value;
    return exports.DEFAULT_MODE;
}
/** Names advertised for a mode, in catalog order. */
function selectToolNames(mode, allNames) {
    if (mode === 'full')
        return [...allNames];
    const wanted = new Set(exports.FACADE_TOOLS);
    if (mode === 'standard')
        for (const name of exports.STANDARD_TOOLS)
            wanted.add(name);
    return allNames.filter((name) => wanted.has(name));
}
/** Tool definitions advertised for a mode. */
function selectTools(mode, tools) {
    const names = new Set(selectToolNames(mode, tools.map((t) => t.name)));
    return tools.filter((t) => names.has(t.name));
}
/**
 * Shorten a tool description for the advertised surface. Full text stays available
 * through wps_help, so the per-request copy only needs to route the model correctly.
 */
function compactDescription(text, max = 110) {
    const flat = (text || '').replace(/\s+/g, ' ').trim();
    if (flat.length <= max)
        return flat;
    const stop = flat.search(/[。.!?]/);
    if (stop >= 0 && stop + 1 <= max)
        return flat.slice(0, stop + 1);
    return flat.slice(0, max - 1) + '…';
}
//# sourceMappingURL=toolset.js.map
"use strict";
/**
 * Input: none
 * Output: the deprecated-tool map applied at server startup
 * Pos: Single place where duplicated upstream tools are collapsed into one canonical name.
 *      一旦我被修改，请更新我的头部注释。
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.DEPRECATED_NAMES = exports.DEPRECATED_TOOLS = void 0;
exports.renameArgs = renameArgs;
/**
 * Tools merged into a canonical name because they drive the same COM action through the
 * same interface. Each entry keeps the old name callable as a forwarder, so existing
 * prompts keep working, but it is hidden from wps_help and no longer a separate concept.
 */
exports.DEPRECATED_TOOLS = {
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
    // 下面这批与规范名驱动同一个 COM action，只是参数拼写或覆盖范围不同。规范名都已扩成
    // "参数并集"，因此旧名转发过去不会丢能力。
    wps_ppt_insert_slide_image: {
        canonical: 'wps_ppt_insert_ppt_image',
        reason: '同一个 insertPptImage，参数名 imagePath 改名为 filePath',
        paramMap: { imagePath: 'filePath' },
    },
    wps_ppt_insert_image: {
        canonical: 'wps_ppt_insert_ppt_image',
        reason: '同一个 insertPptImage，参数名 path 改名为 filePath',
        paramMap: { path: 'filePath' },
    },
    wps_excel_auto_fill: {
        canonical: 'wps_excel_fill_series',
        reason: '同一个 fillSeries；wps_excel_fill_series 现在同时接受 sourceRange/targetRange',
    },
    wps_excel_insert_row: {
        canonical: 'wps_excel_insert_rows',
        reason: '同一个 insertRows，单行是 startRow/count 的特例',
    },
    wps_excel_hide_row: {
        canonical: 'wps_excel_hide_rows',
        reason: '同一个 hideRows；wps_excel_hide_rows 现在也接受 row/count/hide',
    },
    wps_excel_insert_column: {
        canonical: 'wps_excel_insert_columns',
        reason: '同一个 insertColumns，单列是 count 的特例',
    },
    wps_excel_delete_row: {
        canonical: 'wps_excel_delete_rows',
        reason: '同一个 deleteRows；wps_excel_delete_rows 现在也接受 row',
    },
    wps_excel_delete_column: {
        canonical: 'wps_excel_delete_columns',
        reason: '同一个 deleteColumns，单列是 count 的特例',
    },
    wps_ppt_set_animation: {
        canonical: 'wps_ppt_add_animation',
        reason: '同一个 addAnimation，参数名 animationType 改名为 effect',
        paramMap: { animationType: 'effect' },
    },
    wps_ppt_set_transition: {
        canonical: 'wps_ppt_set_slide_transition',
        reason: '同一个 setSlideTransition，参数名 transition 改名为 effect',
        paramMap: { transition: 'effect' },
    },
    wps_ppt_set_background: {
        canonical: 'wps_ppt_set_slide_background',
        reason: '同一个 setSlideBackground；规范名现在也接受平铺的 color/imagePath',
    },
    wps_ppt_add_chart: {
        canonical: 'wps_ppt_insert_ppt_chart',
        reason: '同一个 insertPptChart，参数是规范名的子集',
    },
    wps_ppt_duplicate_slide: {
        canonical: 'wps_ppt_copy_slide',
        reason: '同一个 duplicateSlide，参数是规范名的子集',
    },
    wps_ppt_align_objects: {
        canonical: 'wps_ppt_align_shapes',
        reason: '同一个 alignShapes；不给 shapeIndices 时对齐整页形状，与旧行为一致',
    },
    wps_word_set_font_style: {
        canonical: 'wps_word_set_font',
        reason: '同一个 setFont，参数名 fontName/fontSize 改名为 font_name/font_size',
        paramMap: { fontName: 'font_name', fontSize: 'font_size' },
    },
};
/** Names hidden from wps_help listings. */
exports.DEPRECATED_NAMES = new Set(Object.keys(exports.DEPRECATED_TOOLS));
/** Rename the keys given by a paramMap, leaving everything else untouched. */
function renameArgs(args, map) {
    if (!map)
        return args;
    const out = { ...args };
    for (const [from, to] of Object.entries(map)) {
        if (Object.prototype.hasOwnProperty.call(out, from)) {
            if (!(to in out))
                out[to] = out[from];
            delete out[from];
        }
    }
    return out;
}
//# sourceMappingURL=deprecated.js.map
"use strict";
/**
 * Input: Tool 定义集合
 * Output: Tool 注册数组
 * Pos: MCP Tools 总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tools总入口 - MCP工具汇总注册模块
 * 整合Excel、Word、PPT、Common的所有Tools
 *
 * 使用方法：
 * import { allTools } from './tools';
 * toolRegistry.registerAll(allTools);
 *
 * 或者按需导入：
 * import { excelTools, wordTools, pptTools, commonTools } from './tools';
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.getToolCountByApp = exports.getToolCount = exports.getFormatCode = exports.getAppTypeByExtension = exports.convertFormatHandler = exports.convertFormatDefinition = exports.convertToPdfHandler = exports.convertToPdfDefinition = exports.convertTools = exports.unifyFontHandler = exports.unifyFontDefinition = exports.beautifyHandler = exports.beautifyDefinition = exports.addSlideHandler = exports.addSlideDefinition = exports.slideTools = exports.findReplaceHandler = exports.findReplaceDefinition = exports.insertTextHandler = exports.insertTextDefinition = exports.generateTocHandler = exports.generateTocDefinition = exports.setFontHandler = exports.setFontDefinition = exports.applyStyleHandler = exports.applyStyleDefinition = exports.proofreadTools = exports.contentTools = exports.formatTools = exports.removeDuplicatesHandler = exports.removeDuplicatesDefinition = exports.cleanDataHandler = exports.cleanDataDefinition = exports.writeRangeHandler = exports.writeRangeDefinition = exports.readRangeHandler = exports.readRangeDefinition = exports.diagnoseFormulaHandler = exports.diagnoseFormulaDefinition = exports.generateFormulaHandler = exports.generateFormulaDefinition = exports.setFormulaHandler = exports.setFormulaDefinition = exports.dataTools = exports.formulaTools = exports.commonTools = exports.pptTools = exports.wordTools = exports.excelTools = exports.allTools = void 0;
const excel_1 = require("./excel");
const word_1 = require("./word");
const ppt_1 = require("./ppt");
const common_1 = require("./common");
/**
 * 所有MCP Tools集合（共235个，另有12个内置工具在mcp-server.ts中注册，全局共247个）
 *
 * Excel (82个):
 *   公式(6): set_formula, generate_formula, diagnose_formula, set_array_formula, recalculate, auto_sum
 *   数据(12): read_range, write_range, clean_data, remove_duplicates, sort_range, find_replace, insert_row, add_comment, protect_sheet, set_conditional_format, hide_column, protect_workbook
 *   图表(4): create_chart, update_chart, export_chart_as_image, export_range_as_image
 *   透视表(2): create_pivot_table, update_pivot_table
 *   工作表(16): create_sheet, delete_sheet, rename_sheet, copy_sheet, get_sheet_list, switch_sheet, move_sheet, get_selection, delete_row, insert_column, delete_column, freeze_panes, auto_fill, set_named_range, set_zoom, save_workbook
 *   格式化(10): set_cell_format, set_cell_style, set_border, set_number_format, merge_cells, unmerge_cells, set_column_width, set_row_height, auto_fit_column, auto_fit_row
 *   工作簿(10): open_workbook, get_open_workbooks, switch_workbook, close_workbook, create_workbook, get_cell_value, set_cell_value, get_formula, get_cell_info, clear_range
 *   高级数据(7): auto_filter, copy_range, paste_range, fill_series, transpose, text_to_columns, subtotal
 *
 * Word (32个):
 *   格式化(5): apply_style, set_font, generate_toc, insert_bookmark, set_page_setup
 *   内容(14): insert_text, find_replace, insert_table, set_paragraph, get_active_document, insert_image, insert_page_break, set_font_style, insert_comment, set_text_color, get_paragraphs, find_in_document, smart_fill_field, replace_bookmark_content
 *   文档管理(9): get_open_documents, switch_document, open_document, get_document_text, insert_header, insert_footer, generate_doc_toc, insert_section_break, set_line_spacing
 *   校对(4): enable_track_changes, get_track_changes_status, replace_range, proofread_basic
 *
 * PPT (112个):
 *   幻灯片(5): add_slide, beautify, unify_font, set_font_color, align_objects
 *   幻灯片操作(22): delete_slide, duplicate_slide, move_slide, get_slide_count, get_slide_info, switch_slide, set_slide_layout, get_slide_notes, set_slide_notes, add_shape, set_shape_style, add_textbox, set_slide_title, insert_image, set_shape_text, set_animation, set_background, set_slide_size, set_transition, add_chart, set_shape_fill, add_speaker_notes
 *   演示文稿管理(8): create_presentation, open_presentation, close_presentation, get_open_presentations, switch_presentation, set_slide_theme, copy_slide, insert_slide_image
 *   文本框(7): delete_textbox, get_textboxes, set_textbox_text, set_textbox_style, get_slide_title, set_slide_subtitle, set_slide_content
 *
 * Common (9个):
 *   转换(2): convert_to_pdf, convert_format
 *   （其余7个工具详见 common/ 子模块）
 *
 * 内置工具（12个，在 mcp-server.ts 中注册）:
 *   wps_check_connection, wps_get_active_document, wps_insert_text, wps_get_active_workbook,
 *   wps_get_cell_value, wps_set_cell_value, wps_get_active_presentation, wps_execute_method,
 *   wps_cache_data, wps_get_cached_data, wps_list_cache, wps_clear_cache
 */
exports.allTools = [
    ...excel_1.excelTools,
    ...word_1.wordTools,
    ...ppt_1.pptTools,
    ...common_1.commonTools,
];
// 按应用类型分别导出
var excel_2 = require("./excel");
Object.defineProperty(exports, "excelTools", { enumerable: true, get: function () { return excel_2.excelTools; } });
var word_2 = require("./word");
Object.defineProperty(exports, "wordTools", { enumerable: true, get: function () { return word_2.wordTools; } });
var ppt_2 = require("./ppt");
Object.defineProperty(exports, "pptTools", { enumerable: true, get: function () { return ppt_2.pptTools; } });
var common_2 = require("./common");
Object.defineProperty(exports, "commonTools", { enumerable: true, get: function () { return common_2.commonTools; } });
// Excel相关导出
var excel_3 = require("./excel");
Object.defineProperty(exports, "formulaTools", { enumerable: true, get: function () { return excel_3.formulaTools; } });
Object.defineProperty(exports, "dataTools", { enumerable: true, get: function () { return excel_3.dataTools; } });
Object.defineProperty(exports, "setFormulaDefinition", { enumerable: true, get: function () { return excel_3.setFormulaDefinition; } });
Object.defineProperty(exports, "setFormulaHandler", { enumerable: true, get: function () { return excel_3.setFormulaHandler; } });
Object.defineProperty(exports, "generateFormulaDefinition", { enumerable: true, get: function () { return excel_3.generateFormulaDefinition; } });
Object.defineProperty(exports, "generateFormulaHandler", { enumerable: true, get: function () { return excel_3.generateFormulaHandler; } });
Object.defineProperty(exports, "diagnoseFormulaDefinition", { enumerable: true, get: function () { return excel_3.diagnoseFormulaDefinition; } });
Object.defineProperty(exports, "diagnoseFormulaHandler", { enumerable: true, get: function () { return excel_3.diagnoseFormulaHandler; } });
Object.defineProperty(exports, "readRangeDefinition", { enumerable: true, get: function () { return excel_3.readRangeDefinition; } });
Object.defineProperty(exports, "readRangeHandler", { enumerable: true, get: function () { return excel_3.readRangeHandler; } });
Object.defineProperty(exports, "writeRangeDefinition", { enumerable: true, get: function () { return excel_3.writeRangeDefinition; } });
Object.defineProperty(exports, "writeRangeHandler", { enumerable: true, get: function () { return excel_3.writeRangeHandler; } });
Object.defineProperty(exports, "cleanDataDefinition", { enumerable: true, get: function () { return excel_3.cleanDataDefinition; } });
Object.defineProperty(exports, "cleanDataHandler", { enumerable: true, get: function () { return excel_3.cleanDataHandler; } });
Object.defineProperty(exports, "removeDuplicatesDefinition", { enumerable: true, get: function () { return excel_3.removeDuplicatesDefinition; } });
Object.defineProperty(exports, "removeDuplicatesHandler", { enumerable: true, get: function () { return excel_3.removeDuplicatesHandler; } });
// Word相关导出
var word_3 = require("./word");
Object.defineProperty(exports, "formatTools", { enumerable: true, get: function () { return word_3.formatTools; } });
Object.defineProperty(exports, "contentTools", { enumerable: true, get: function () { return word_3.contentTools; } });
Object.defineProperty(exports, "proofreadTools", { enumerable: true, get: function () { return word_3.proofreadTools; } });
Object.defineProperty(exports, "applyStyleDefinition", { enumerable: true, get: function () { return word_3.applyStyleDefinition; } });
Object.defineProperty(exports, "applyStyleHandler", { enumerable: true, get: function () { return word_3.applyStyleHandler; } });
Object.defineProperty(exports, "setFontDefinition", { enumerable: true, get: function () { return word_3.setFontDefinition; } });
Object.defineProperty(exports, "setFontHandler", { enumerable: true, get: function () { return word_3.setFontHandler; } });
Object.defineProperty(exports, "generateTocDefinition", { enumerable: true, get: function () { return word_3.generateTocDefinition; } });
Object.defineProperty(exports, "generateTocHandler", { enumerable: true, get: function () { return word_3.generateTocHandler; } });
Object.defineProperty(exports, "insertTextDefinition", { enumerable: true, get: function () { return word_3.insertTextDefinition; } });
Object.defineProperty(exports, "insertTextHandler", { enumerable: true, get: function () { return word_3.insertTextHandler; } });
Object.defineProperty(exports, "findReplaceDefinition", { enumerable: true, get: function () { return word_3.findReplaceDefinition; } });
Object.defineProperty(exports, "findReplaceHandler", { enumerable: true, get: function () { return word_3.findReplaceHandler; } });
// PPT相关导出
var ppt_3 = require("./ppt");
Object.defineProperty(exports, "slideTools", { enumerable: true, get: function () { return ppt_3.slideTools; } });
Object.defineProperty(exports, "addSlideDefinition", { enumerable: true, get: function () { return ppt_3.addSlideDefinition; } });
Object.defineProperty(exports, "addSlideHandler", { enumerable: true, get: function () { return ppt_3.addSlideHandler; } });
Object.defineProperty(exports, "beautifyDefinition", { enumerable: true, get: function () { return ppt_3.beautifyDefinition; } });
Object.defineProperty(exports, "beautifyHandler", { enumerable: true, get: function () { return ppt_3.beautifyHandler; } });
Object.defineProperty(exports, "unifyFontDefinition", { enumerable: true, get: function () { return ppt_3.unifyFontDefinition; } });
Object.defineProperty(exports, "unifyFontHandler", { enumerable: true, get: function () { return ppt_3.unifyFontHandler; } });
// Common相关导出
var common_3 = require("./common");
Object.defineProperty(exports, "convertTools", { enumerable: true, get: function () { return common_3.convertTools; } });
Object.defineProperty(exports, "convertToPdfDefinition", { enumerable: true, get: function () { return common_3.convertToPdfDefinition; } });
Object.defineProperty(exports, "convertToPdfHandler", { enumerable: true, get: function () { return common_3.convertToPdfHandler; } });
Object.defineProperty(exports, "convertFormatDefinition", { enumerable: true, get: function () { return common_3.convertFormatDefinition; } });
Object.defineProperty(exports, "convertFormatHandler", { enumerable: true, get: function () { return common_3.convertFormatHandler; } });
Object.defineProperty(exports, "getAppTypeByExtension", { enumerable: true, get: function () { return common_3.getAppTypeByExtension; } });
Object.defineProperty(exports, "getFormatCode", { enumerable: true, get: function () { return common_3.getFormatCode; } });
/**
 * 获取所有Tool的数量
 */
const getToolCount = () => exports.allTools.length;
exports.getToolCount = getToolCount;
/**
 * 获取按应用分类的Tool数量
 */
const getToolCountByApp = () => ({
    excel: excel_1.excelTools.length,
    word: word_1.wordTools.length,
    ppt: ppt_1.pptTools.length,
    common: common_1.commonTools.length,
});
exports.getToolCountByApp = getToolCountByApp;
exports.default = exports.allTools;
//# sourceMappingURL=index.js.map
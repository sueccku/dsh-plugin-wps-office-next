"use strict";
/**
 * Input: Excel 工具定义
 * Output: Excel 工具注册数组
 * Pos: Excel Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel Tools入口 - Excel工具汇总模块
 * 整合公式、数据处理、透视表、图表、工作表、格式化的所有Tools
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.updateChartDefinition = exports.createChartHandler = exports.createChartDefinition = exports.updatePivotTableHandler = exports.updatePivotTableDefinition = exports.createPivotTableHandler = exports.createPivotTableDefinition = exports.setZoomHandler = exports.setZoomDefinition = exports.protectWorkbookHandler = exports.protectWorkbookDefinition = exports.setConditionalFormatHandler = exports.setConditionalFormatDefinition = exports.protectSheetHandler = exports.protectSheetDefinition = exports.addCommentHandler = exports.addCommentDefinition = exports.findReplaceHandler = exports.findReplaceDefinition = exports.sortRangeHandler = exports.sortRangeDefinition = exports.removeDuplicatesHandler = exports.removeDuplicatesDefinition = exports.cleanDataHandler = exports.cleanDataDefinition = exports.writeRangeHandler = exports.writeRangeDefinition = exports.readRangeHandler = exports.readRangeDefinition = exports.setPrintAreaHandler = exports.setPrintAreaDefinition = exports.evaluateFormulaHandler = exports.evaluateFormulaDefinition = exports.diagnoseFormulaHandler = exports.diagnoseFormulaDefinition = exports.generateFormulaHandler = exports.generateFormulaDefinition = exports.setFormulaHandler = exports.setFormulaDefinition = exports.commentProtectTools = exports.rowColumnTools = exports.dataAdvancedTools = exports.workbookTools = exports.excelFormatTools = exports.sheetTools = exports.chartTools = exports.pivotTools = exports.dataTools = exports.formulaTools = exports.excelTools = void 0;
exports.getOpenWorkbooksDefinition = exports.openWorkbookHandler = exports.openWorkbookDefinition = exports.setDataValidationHandler = exports.setDataValidationDefinition = exports.setRowHeightHandler = exports.setRowHeightDefinition = exports.setColumnWidthHandler = exports.setColumnWidthDefinition = exports.unmergeCellsHandler = exports.unmergeCellsDefinition = exports.mergeCellsHandler = exports.mergeCellsDefinition = exports.setNumberFormatHandler = exports.setNumberFormatDefinition = exports.setBorderHandler = exports.setBorderDefinition = exports.setCellStyleHandler = exports.setCellStyleDefinition = exports.setCellFormatHandler = exports.setCellFormatDefinition = exports.hideColumnHandler = exports.hideColumnDefinition = exports.autoSumHandler = exports.autoSumDefinition = exports.setNamedRangeHandler = exports.setNamedRangeDefinition = exports.freezePanesHandler = exports.freezePanesDefinition = exports.getSelectionHandler = exports.getSelectionDefinition = exports.moveSheetHandler = exports.moveSheetDefinition = exports.switchSheetHandler = exports.switchSheetDefinition = exports.getSheetListHandler = exports.getSheetListDefinition = exports.copySheetHandler = exports.copySheetDefinition = exports.renameSheetHandler = exports.renameSheetDefinition = exports.deleteSheetHandler = exports.deleteSheetDefinition = exports.createSheetHandler = exports.createSheetDefinition = exports.exportRangeAsImageHandler = exports.exportRangeAsImageDefinition = exports.exportChartAsImageHandler = exports.exportChartAsImageDefinition = exports.updateChartHandler = void 0;
exports.getCellCommentsDefinition = exports.deleteCellCommentHandler = exports.deleteCellCommentDefinition = exports.groupRowsHandler = exports.groupRowsDefinition = exports.showColumnsHandler = exports.showColumnsDefinition = exports.showRowsHandler = exports.showRowsDefinition = exports.hideRowsHandler = exports.hideRowsDefinition = exports.deleteColumnsHandler = exports.deleteColumnsDefinition = exports.deleteRowsHandler = exports.deleteRowsDefinition = exports.insertColumnsHandler = exports.insertColumnsDefinition = exports.insertRowsHandler = exports.insertRowsDefinition = exports.subtotalHandler = exports.subtotalDefinition = exports.textToColumnsHandler = exports.textToColumnsDefinition = exports.transposeHandler = exports.transposeDefinition = exports.fillSeriesHandler = exports.fillSeriesDefinition = exports.pasteRangeHandler = exports.pasteRangeDefinition = exports.copyRangeHandler = exports.copyRangeDefinition = exports.autoFilterHandler = exports.autoFilterDefinition = exports.clearRangeHandler = exports.clearRangeDefinition = exports.getCellInfoHandler = exports.getCellInfoDefinition = exports.getFormulaHandler = exports.getFormulaDefinition = exports.setCellValueHandler = exports.setCellValueDefinition = exports.getCellValueHandler = exports.getCellValueDefinition = exports.createWorkbookHandler = exports.createWorkbookDefinition = exports.closeWorkbookHandler = exports.closeWorkbookDefinition = exports.switchWorkbookHandler = exports.switchWorkbookDefinition = exports.getOpenWorkbooksHandler = void 0;
exports.setHyperlinkHandler = exports.setHyperlinkDefinition = exports.insertExcelImageHandler = exports.insertExcelImageDefinition = exports.setArrayFormulaHandler = exports.setArrayFormulaDefinition = exports.lockCellsHandler = exports.lockCellsDefinition = exports.unprotectSheetHandler = exports.unprotectSheetDefinition = exports.getCellCommentsHandler = void 0;
const formula_1 = require("./formula");
const data_1 = require("./data");
const pivot_1 = require("./pivot");
const chart_1 = require("./chart");
const sheet_1 = require("./sheet");
const format_1 = require("./format");
const workbook_1 = require("./workbook");
const data_advanced_1 = require("./data-advanced");
const row_column_1 = require("./row-column");
const comment_protect_1 = require("./comment-protect");
const missing_halves_1 = require("./missing-halves");
const list_object_1 = require("./list-object");
const sheet_settings_1 = require("./sheet-settings");
const advanced_1 = require("./advanced");
/**
 * 所有Excel相关的Tools
 * 包含：
 * - 公式Tools: set_formula, generate_formula, diagnose_formula
 * - 数据Tools: read_range, write_range, clean_data, remove_duplicates, sort_range, find_replace, insert_row, add_comment, protect_sheet, set_conditional_format
 * - 透视表Tools: create_pivot_table, update_pivot_table
 * - 图表Tools: create_chart, update_chart, export_chart_as_image, export_range_as_image
 * - 工作表Tools: create_sheet, delete_sheet, rename_sheet, copy_sheet, get_sheet_list, switch_sheet, move_sheet, get_selection, delete_row, insert_column, delete_column, freeze_panes, auto_fill, set_named_range
 * - 格式化Tools: set_cell_format, set_cell_style, set_border, set_number_format, merge_cells, unmerge_cells, set_column_width, set_row_height
 * - 行列Tools: insert_rows, insert_columns, delete_rows, delete_columns, hide_rows, show_rows, show_columns, group_rows
 * - 批注保护Tools: delete_cell_comment, get_cell_comments, unprotect_sheet, lock_cells, set_array_formula, insert_excel_image, set_hyperlink
 * - 补全Tools（P2 第一波）: get_sheet_info, auto_fit(/columns/rows), set_wrap_text, find_in_sheet,
 *   get_named_ranges, delete_named_range
 * - 表Tools（P2-2）: create_list_object, get_list_objects, add_list_row, delete_list_row,
 *   update_list_object, set_list_object_totals, resize_list_object, unlist_list_object
 * - 页面与打印Tools（P2-3）: get_sheet_settings, set_sheet_page_setup, set_sheet_print_titles,
 *   set_sheet_header_footer, set_sheet_appearance, set_outline_levels, reset_page_breaks, get_formula_audit
 * - 高级项Tools（P2-4）: get_pivot_tables, refresh_pivot_tables, clear_pivot_table, refresh_all_data,
 *   goal_seek, add_sparkline, clear_sparkline, delete_chart, set_chart_labels
 */
exports.excelTools = [
    ...formula_1.formulaTools,
    ...data_1.dataTools,
    ...pivot_1.pivotTools,
    ...chart_1.chartTools,
    ...sheet_1.sheetTools,
    ...format_1.excelFormatTools,
    ...workbook_1.workbookTools,
    ...data_advanced_1.dataAdvancedTools,
    ...row_column_1.rowColumnTools,
    ...comment_protect_1.commentProtectTools,
    ...missing_halves_1.missingHalfTools,
    ...list_object_1.listObjectTools,
    ...sheet_settings_1.sheetSettingsTools,
    ...advanced_1.advancedTools,
];
// 分别导出，方便按需使用
var formula_2 = require("./formula");
Object.defineProperty(exports, "formulaTools", { enumerable: true, get: function () { return formula_2.formulaTools; } });
var data_2 = require("./data");
Object.defineProperty(exports, "dataTools", { enumerable: true, get: function () { return data_2.dataTools; } });
var pivot_2 = require("./pivot");
Object.defineProperty(exports, "pivotTools", { enumerable: true, get: function () { return pivot_2.pivotTools; } });
var chart_2 = require("./chart");
Object.defineProperty(exports, "chartTools", { enumerable: true, get: function () { return chart_2.chartTools; } });
var sheet_2 = require("./sheet");
Object.defineProperty(exports, "sheetTools", { enumerable: true, get: function () { return sheet_2.sheetTools; } });
var format_2 = require("./format");
Object.defineProperty(exports, "excelFormatTools", { enumerable: true, get: function () { return format_2.excelFormatTools; } });
var workbook_2 = require("./workbook");
Object.defineProperty(exports, "workbookTools", { enumerable: true, get: function () { return workbook_2.workbookTools; } });
var data_advanced_2 = require("./data-advanced");
Object.defineProperty(exports, "dataAdvancedTools", { enumerable: true, get: function () { return data_advanced_2.dataAdvancedTools; } });
var row_column_2 = require("./row-column");
Object.defineProperty(exports, "rowColumnTools", { enumerable: true, get: function () { return row_column_2.rowColumnTools; } });
var comment_protect_2 = require("./comment-protect");
Object.defineProperty(exports, "commentProtectTools", { enumerable: true, get: function () { return comment_protect_2.commentProtectTools; } });
// 导出单独的定义和处理器，方便测试
var formula_3 = require("./formula");
Object.defineProperty(exports, "setFormulaDefinition", { enumerable: true, get: function () { return formula_3.setFormulaDefinition; } });
Object.defineProperty(exports, "setFormulaHandler", { enumerable: true, get: function () { return formula_3.setFormulaHandler; } });
Object.defineProperty(exports, "generateFormulaDefinition", { enumerable: true, get: function () { return formula_3.generateFormulaDefinition; } });
Object.defineProperty(exports, "generateFormulaHandler", { enumerable: true, get: function () { return formula_3.generateFormulaHandler; } });
Object.defineProperty(exports, "diagnoseFormulaDefinition", { enumerable: true, get: function () { return formula_3.diagnoseFormulaDefinition; } });
Object.defineProperty(exports, "diagnoseFormulaHandler", { enumerable: true, get: function () { return formula_3.diagnoseFormulaHandler; } });
Object.defineProperty(exports, "evaluateFormulaDefinition", { enumerable: true, get: function () { return formula_3.evaluateFormulaDefinition; } });
Object.defineProperty(exports, "evaluateFormulaHandler", { enumerable: true, get: function () { return formula_3.evaluateFormulaHandler; } });
Object.defineProperty(exports, "setPrintAreaDefinition", { enumerable: true, get: function () { return formula_3.setPrintAreaDefinition; } });
Object.defineProperty(exports, "setPrintAreaHandler", { enumerable: true, get: function () { return formula_3.setPrintAreaHandler; } });
var data_3 = require("./data");
Object.defineProperty(exports, "readRangeDefinition", { enumerable: true, get: function () { return data_3.readRangeDefinition; } });
Object.defineProperty(exports, "readRangeHandler", { enumerable: true, get: function () { return data_3.readRangeHandler; } });
Object.defineProperty(exports, "writeRangeDefinition", { enumerable: true, get: function () { return data_3.writeRangeDefinition; } });
Object.defineProperty(exports, "writeRangeHandler", { enumerable: true, get: function () { return data_3.writeRangeHandler; } });
Object.defineProperty(exports, "cleanDataDefinition", { enumerable: true, get: function () { return data_3.cleanDataDefinition; } });
Object.defineProperty(exports, "cleanDataHandler", { enumerable: true, get: function () { return data_3.cleanDataHandler; } });
Object.defineProperty(exports, "removeDuplicatesDefinition", { enumerable: true, get: function () { return data_3.removeDuplicatesDefinition; } });
Object.defineProperty(exports, "removeDuplicatesHandler", { enumerable: true, get: function () { return data_3.removeDuplicatesHandler; } });
Object.defineProperty(exports, "sortRangeDefinition", { enumerable: true, get: function () { return data_3.sortRangeDefinition; } });
Object.defineProperty(exports, "sortRangeHandler", { enumerable: true, get: function () { return data_3.sortRangeHandler; } });
Object.defineProperty(exports, "findReplaceDefinition", { enumerable: true, get: function () { return data_3.findReplaceDefinition; } });
Object.defineProperty(exports, "findReplaceHandler", { enumerable: true, get: function () { return data_3.findReplaceHandler; } });
Object.defineProperty(exports, "addCommentDefinition", { enumerable: true, get: function () { return data_3.addCommentDefinition; } });
Object.defineProperty(exports, "addCommentHandler", { enumerable: true, get: function () { return data_3.addCommentHandler; } });
Object.defineProperty(exports, "protectSheetDefinition", { enumerable: true, get: function () { return data_3.protectSheetDefinition; } });
Object.defineProperty(exports, "protectSheetHandler", { enumerable: true, get: function () { return data_3.protectSheetHandler; } });
Object.defineProperty(exports, "setConditionalFormatDefinition", { enumerable: true, get: function () { return data_3.setConditionalFormatDefinition; } });
Object.defineProperty(exports, "setConditionalFormatHandler", { enumerable: true, get: function () { return data_3.setConditionalFormatHandler; } });
Object.defineProperty(exports, "protectWorkbookDefinition", { enumerable: true, get: function () { return data_3.protectWorkbookDefinition; } });
Object.defineProperty(exports, "protectWorkbookHandler", { enumerable: true, get: function () { return data_3.protectWorkbookHandler; } });
Object.defineProperty(exports, "setZoomDefinition", { enumerable: true, get: function () { return data_3.setZoomDefinition; } });
Object.defineProperty(exports, "setZoomHandler", { enumerable: true, get: function () { return data_3.setZoomHandler; } });
var pivot_3 = require("./pivot");
Object.defineProperty(exports, "createPivotTableDefinition", { enumerable: true, get: function () { return pivot_3.createPivotTableDefinition; } });
Object.defineProperty(exports, "createPivotTableHandler", { enumerable: true, get: function () { return pivot_3.createPivotTableHandler; } });
Object.defineProperty(exports, "updatePivotTableDefinition", { enumerable: true, get: function () { return pivot_3.updatePivotTableDefinition; } });
Object.defineProperty(exports, "updatePivotTableHandler", { enumerable: true, get: function () { return pivot_3.updatePivotTableHandler; } });
var chart_3 = require("./chart");
Object.defineProperty(exports, "createChartDefinition", { enumerable: true, get: function () { return chart_3.createChartDefinition; } });
Object.defineProperty(exports, "createChartHandler", { enumerable: true, get: function () { return chart_3.createChartHandler; } });
Object.defineProperty(exports, "updateChartDefinition", { enumerable: true, get: function () { return chart_3.updateChartDefinition; } });
Object.defineProperty(exports, "updateChartHandler", { enumerable: true, get: function () { return chart_3.updateChartHandler; } });
Object.defineProperty(exports, "exportChartAsImageDefinition", { enumerable: true, get: function () { return chart_3.exportChartAsImageDefinition; } });
Object.defineProperty(exports, "exportChartAsImageHandler", { enumerable: true, get: function () { return chart_3.exportChartAsImageHandler; } });
Object.defineProperty(exports, "exportRangeAsImageDefinition", { enumerable: true, get: function () { return chart_3.exportRangeAsImageDefinition; } });
Object.defineProperty(exports, "exportRangeAsImageHandler", { enumerable: true, get: function () { return chart_3.exportRangeAsImageHandler; } });
var sheet_3 = require("./sheet");
Object.defineProperty(exports, "createSheetDefinition", { enumerable: true, get: function () { return sheet_3.createSheetDefinition; } });
Object.defineProperty(exports, "createSheetHandler", { enumerable: true, get: function () { return sheet_3.createSheetHandler; } });
Object.defineProperty(exports, "deleteSheetDefinition", { enumerable: true, get: function () { return sheet_3.deleteSheetDefinition; } });
Object.defineProperty(exports, "deleteSheetHandler", { enumerable: true, get: function () { return sheet_3.deleteSheetHandler; } });
Object.defineProperty(exports, "renameSheetDefinition", { enumerable: true, get: function () { return sheet_3.renameSheetDefinition; } });
Object.defineProperty(exports, "renameSheetHandler", { enumerable: true, get: function () { return sheet_3.renameSheetHandler; } });
Object.defineProperty(exports, "copySheetDefinition", { enumerable: true, get: function () { return sheet_3.copySheetDefinition; } });
Object.defineProperty(exports, "copySheetHandler", { enumerable: true, get: function () { return sheet_3.copySheetHandler; } });
Object.defineProperty(exports, "getSheetListDefinition", { enumerable: true, get: function () { return sheet_3.getSheetListDefinition; } });
Object.defineProperty(exports, "getSheetListHandler", { enumerable: true, get: function () { return sheet_3.getSheetListHandler; } });
Object.defineProperty(exports, "switchSheetDefinition", { enumerable: true, get: function () { return sheet_3.switchSheetDefinition; } });
Object.defineProperty(exports, "switchSheetHandler", { enumerable: true, get: function () { return sheet_3.switchSheetHandler; } });
Object.defineProperty(exports, "moveSheetDefinition", { enumerable: true, get: function () { return sheet_3.moveSheetDefinition; } });
Object.defineProperty(exports, "moveSheetHandler", { enumerable: true, get: function () { return sheet_3.moveSheetHandler; } });
Object.defineProperty(exports, "getSelectionDefinition", { enumerable: true, get: function () { return sheet_3.getSelectionDefinition; } });
Object.defineProperty(exports, "getSelectionHandler", { enumerable: true, get: function () { return sheet_3.getSelectionHandler; } });
Object.defineProperty(exports, "freezePanesDefinition", { enumerable: true, get: function () { return sheet_3.freezePanesDefinition; } });
Object.defineProperty(exports, "freezePanesHandler", { enumerable: true, get: function () { return sheet_3.freezePanesHandler; } });
Object.defineProperty(exports, "setNamedRangeDefinition", { enumerable: true, get: function () { return sheet_3.setNamedRangeDefinition; } });
Object.defineProperty(exports, "setNamedRangeHandler", { enumerable: true, get: function () { return sheet_3.setNamedRangeHandler; } });
Object.defineProperty(exports, "autoSumDefinition", { enumerable: true, get: function () { return sheet_3.autoSumDefinition; } });
Object.defineProperty(exports, "autoSumHandler", { enumerable: true, get: function () { return sheet_3.autoSumHandler; } });
Object.defineProperty(exports, "hideColumnDefinition", { enumerable: true, get: function () { return sheet_3.hideColumnDefinition; } });
Object.defineProperty(exports, "hideColumnHandler", { enumerable: true, get: function () { return sheet_3.hideColumnHandler; } });
var format_3 = require("./format");
Object.defineProperty(exports, "setCellFormatDefinition", { enumerable: true, get: function () { return format_3.setCellFormatDefinition; } });
Object.defineProperty(exports, "setCellFormatHandler", { enumerable: true, get: function () { return format_3.setCellFormatHandler; } });
Object.defineProperty(exports, "setCellStyleDefinition", { enumerable: true, get: function () { return format_3.setCellStyleDefinition; } });
Object.defineProperty(exports, "setCellStyleHandler", { enumerable: true, get: function () { return format_3.setCellStyleHandler; } });
Object.defineProperty(exports, "setBorderDefinition", { enumerable: true, get: function () { return format_3.setBorderDefinition; } });
Object.defineProperty(exports, "setBorderHandler", { enumerable: true, get: function () { return format_3.setBorderHandler; } });
Object.defineProperty(exports, "setNumberFormatDefinition", { enumerable: true, get: function () { return format_3.setNumberFormatDefinition; } });
Object.defineProperty(exports, "setNumberFormatHandler", { enumerable: true, get: function () { return format_3.setNumberFormatHandler; } });
Object.defineProperty(exports, "mergeCellsDefinition", { enumerable: true, get: function () { return format_3.mergeCellsDefinition; } });
Object.defineProperty(exports, "mergeCellsHandler", { enumerable: true, get: function () { return format_3.mergeCellsHandler; } });
Object.defineProperty(exports, "unmergeCellsDefinition", { enumerable: true, get: function () { return format_3.unmergeCellsDefinition; } });
Object.defineProperty(exports, "unmergeCellsHandler", { enumerable: true, get: function () { return format_3.unmergeCellsHandler; } });
Object.defineProperty(exports, "setColumnWidthDefinition", { enumerable: true, get: function () { return format_3.setColumnWidthDefinition; } });
Object.defineProperty(exports, "setColumnWidthHandler", { enumerable: true, get: function () { return format_3.setColumnWidthHandler; } });
Object.defineProperty(exports, "setRowHeightDefinition", { enumerable: true, get: function () { return format_3.setRowHeightDefinition; } });
Object.defineProperty(exports, "setRowHeightHandler", { enumerable: true, get: function () { return format_3.setRowHeightHandler; } });
Object.defineProperty(exports, "setDataValidationDefinition", { enumerable: true, get: function () { return format_3.setDataValidationDefinition; } });
Object.defineProperty(exports, "setDataValidationHandler", { enumerable: true, get: function () { return format_3.setDataValidationHandler; } });
var workbook_3 = require("./workbook");
Object.defineProperty(exports, "openWorkbookDefinition", { enumerable: true, get: function () { return workbook_3.openWorkbookDefinition; } });
Object.defineProperty(exports, "openWorkbookHandler", { enumerable: true, get: function () { return workbook_3.openWorkbookHandler; } });
Object.defineProperty(exports, "getOpenWorkbooksDefinition", { enumerable: true, get: function () { return workbook_3.getOpenWorkbooksDefinition; } });
Object.defineProperty(exports, "getOpenWorkbooksHandler", { enumerable: true, get: function () { return workbook_3.getOpenWorkbooksHandler; } });
Object.defineProperty(exports, "switchWorkbookDefinition", { enumerable: true, get: function () { return workbook_3.switchWorkbookDefinition; } });
Object.defineProperty(exports, "switchWorkbookHandler", { enumerable: true, get: function () { return workbook_3.switchWorkbookHandler; } });
Object.defineProperty(exports, "closeWorkbookDefinition", { enumerable: true, get: function () { return workbook_3.closeWorkbookDefinition; } });
Object.defineProperty(exports, "closeWorkbookHandler", { enumerable: true, get: function () { return workbook_3.closeWorkbookHandler; } });
Object.defineProperty(exports, "createWorkbookDefinition", { enumerable: true, get: function () { return workbook_3.createWorkbookDefinition; } });
Object.defineProperty(exports, "createWorkbookHandler", { enumerable: true, get: function () { return workbook_3.createWorkbookHandler; } });
Object.defineProperty(exports, "getCellValueDefinition", { enumerable: true, get: function () { return workbook_3.getCellValueDefinition; } });
Object.defineProperty(exports, "getCellValueHandler", { enumerable: true, get: function () { return workbook_3.getCellValueHandler; } });
Object.defineProperty(exports, "setCellValueDefinition", { enumerable: true, get: function () { return workbook_3.setCellValueDefinition; } });
Object.defineProperty(exports, "setCellValueHandler", { enumerable: true, get: function () { return workbook_3.setCellValueHandler; } });
Object.defineProperty(exports, "getFormulaDefinition", { enumerable: true, get: function () { return workbook_3.getFormulaDefinition; } });
Object.defineProperty(exports, "getFormulaHandler", { enumerable: true, get: function () { return workbook_3.getFormulaHandler; } });
Object.defineProperty(exports, "getCellInfoDefinition", { enumerable: true, get: function () { return workbook_3.getCellInfoDefinition; } });
Object.defineProperty(exports, "getCellInfoHandler", { enumerable: true, get: function () { return workbook_3.getCellInfoHandler; } });
Object.defineProperty(exports, "clearRangeDefinition", { enumerable: true, get: function () { return workbook_3.clearRangeDefinition; } });
Object.defineProperty(exports, "clearRangeHandler", { enumerable: true, get: function () { return workbook_3.clearRangeHandler; } });
var data_advanced_3 = require("./data-advanced");
Object.defineProperty(exports, "autoFilterDefinition", { enumerable: true, get: function () { return data_advanced_3.autoFilterDefinition; } });
Object.defineProperty(exports, "autoFilterHandler", { enumerable: true, get: function () { return data_advanced_3.autoFilterHandler; } });
Object.defineProperty(exports, "copyRangeDefinition", { enumerable: true, get: function () { return data_advanced_3.copyRangeDefinition; } });
Object.defineProperty(exports, "copyRangeHandler", { enumerable: true, get: function () { return data_advanced_3.copyRangeHandler; } });
Object.defineProperty(exports, "pasteRangeDefinition", { enumerable: true, get: function () { return data_advanced_3.pasteRangeDefinition; } });
Object.defineProperty(exports, "pasteRangeHandler", { enumerable: true, get: function () { return data_advanced_3.pasteRangeHandler; } });
Object.defineProperty(exports, "fillSeriesDefinition", { enumerable: true, get: function () { return data_advanced_3.fillSeriesDefinition; } });
Object.defineProperty(exports, "fillSeriesHandler", { enumerable: true, get: function () { return data_advanced_3.fillSeriesHandler; } });
Object.defineProperty(exports, "transposeDefinition", { enumerable: true, get: function () { return data_advanced_3.transposeDefinition; } });
Object.defineProperty(exports, "transposeHandler", { enumerable: true, get: function () { return data_advanced_3.transposeHandler; } });
Object.defineProperty(exports, "textToColumnsDefinition", { enumerable: true, get: function () { return data_advanced_3.textToColumnsDefinition; } });
Object.defineProperty(exports, "textToColumnsHandler", { enumerable: true, get: function () { return data_advanced_3.textToColumnsHandler; } });
Object.defineProperty(exports, "subtotalDefinition", { enumerable: true, get: function () { return data_advanced_3.subtotalDefinition; } });
Object.defineProperty(exports, "subtotalHandler", { enumerable: true, get: function () { return data_advanced_3.subtotalHandler; } });
var row_column_3 = require("./row-column");
Object.defineProperty(exports, "insertRowsDefinition", { enumerable: true, get: function () { return row_column_3.insertRowsDefinition; } });
Object.defineProperty(exports, "insertRowsHandler", { enumerable: true, get: function () { return row_column_3.insertRowsHandler; } });
Object.defineProperty(exports, "insertColumnsDefinition", { enumerable: true, get: function () { return row_column_3.insertColumnsDefinition; } });
Object.defineProperty(exports, "insertColumnsHandler", { enumerable: true, get: function () { return row_column_3.insertColumnsHandler; } });
Object.defineProperty(exports, "deleteRowsDefinition", { enumerable: true, get: function () { return row_column_3.deleteRowsDefinition; } });
Object.defineProperty(exports, "deleteRowsHandler", { enumerable: true, get: function () { return row_column_3.deleteRowsHandler; } });
Object.defineProperty(exports, "deleteColumnsDefinition", { enumerable: true, get: function () { return row_column_3.deleteColumnsDefinition; } });
Object.defineProperty(exports, "deleteColumnsHandler", { enumerable: true, get: function () { return row_column_3.deleteColumnsHandler; } });
Object.defineProperty(exports, "hideRowsDefinition", { enumerable: true, get: function () { return row_column_3.hideRowsDefinition; } });
Object.defineProperty(exports, "hideRowsHandler", { enumerable: true, get: function () { return row_column_3.hideRowsHandler; } });
Object.defineProperty(exports, "showRowsDefinition", { enumerable: true, get: function () { return row_column_3.showRowsDefinition; } });
Object.defineProperty(exports, "showRowsHandler", { enumerable: true, get: function () { return row_column_3.showRowsHandler; } });
Object.defineProperty(exports, "showColumnsDefinition", { enumerable: true, get: function () { return row_column_3.showColumnsDefinition; } });
Object.defineProperty(exports, "showColumnsHandler", { enumerable: true, get: function () { return row_column_3.showColumnsHandler; } });
Object.defineProperty(exports, "groupRowsDefinition", { enumerable: true, get: function () { return row_column_3.groupRowsDefinition; } });
Object.defineProperty(exports, "groupRowsHandler", { enumerable: true, get: function () { return row_column_3.groupRowsHandler; } });
var comment_protect_3 = require("./comment-protect");
Object.defineProperty(exports, "deleteCellCommentDefinition", { enumerable: true, get: function () { return comment_protect_3.deleteCellCommentDefinition; } });
Object.defineProperty(exports, "deleteCellCommentHandler", { enumerable: true, get: function () { return comment_protect_3.deleteCellCommentHandler; } });
Object.defineProperty(exports, "getCellCommentsDefinition", { enumerable: true, get: function () { return comment_protect_3.getCellCommentsDefinition; } });
Object.defineProperty(exports, "getCellCommentsHandler", { enumerable: true, get: function () { return comment_protect_3.getCellCommentsHandler; } });
Object.defineProperty(exports, "unprotectSheetDefinition", { enumerable: true, get: function () { return comment_protect_3.unprotectSheetDefinition; } });
Object.defineProperty(exports, "unprotectSheetHandler", { enumerable: true, get: function () { return comment_protect_3.unprotectSheetHandler; } });
Object.defineProperty(exports, "lockCellsDefinition", { enumerable: true, get: function () { return comment_protect_3.lockCellsDefinition; } });
Object.defineProperty(exports, "lockCellsHandler", { enumerable: true, get: function () { return comment_protect_3.lockCellsHandler; } });
Object.defineProperty(exports, "setArrayFormulaDefinition", { enumerable: true, get: function () { return comment_protect_3.setArrayFormulaDefinition; } });
Object.defineProperty(exports, "setArrayFormulaHandler", { enumerable: true, get: function () { return comment_protect_3.setArrayFormulaHandler; } });
Object.defineProperty(exports, "insertExcelImageDefinition", { enumerable: true, get: function () { return comment_protect_3.insertExcelImageDefinition; } });
Object.defineProperty(exports, "insertExcelImageHandler", { enumerable: true, get: function () { return comment_protect_3.insertExcelImageHandler; } });
Object.defineProperty(exports, "setHyperlinkDefinition", { enumerable: true, get: function () { return comment_protect_3.setHyperlinkDefinition; } });
Object.defineProperty(exports, "setHyperlinkHandler", { enumerable: true, get: function () { return comment_protect_3.setHyperlinkHandler; } });
exports.default = exports.excelTools;
//# sourceMappingURL=index.js.map
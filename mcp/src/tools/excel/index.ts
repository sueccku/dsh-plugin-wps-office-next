/**
 * Input: Excel 工具定义
 * Output: Excel 工具注册数组
 * Pos: Excel Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Excel Tools入口 - Excel工具汇总模块
 * 整合公式、数据处理、透视表、图表、工作表、格式化的所有Tools
 */

import { RegisteredTool } from '../../types/tools';
import { formulaTools } from './formula';
import { dataTools } from './data';
import { pivotTools } from './pivot';
import { chartTools } from './chart';
import { sheetTools } from './sheet';
import { excelFormatTools } from './format';
import { workbookTools } from './workbook';
import { dataAdvancedTools } from './data-advanced';
import { rowColumnTools } from './row-column';
import { commentProtectTools } from './comment-protect';

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
 */
export const excelTools: RegisteredTool[] = [
  ...formulaTools,
  ...dataTools,
  ...pivotTools,
  ...chartTools,
  ...sheetTools,
  ...excelFormatTools,
  ...workbookTools,
  ...dataAdvancedTools,
  ...rowColumnTools,
  ...commentProtectTools,
];

// 分别导出，方便按需使用
export { formulaTools } from './formula';
export { dataTools } from './data';
export { pivotTools } from './pivot';
export { chartTools } from './chart';
export { sheetTools } from './sheet';
export { excelFormatTools } from './format';
export { workbookTools } from './workbook';
export { dataAdvancedTools } from './data-advanced';
export { rowColumnTools } from './row-column';
export { commentProtectTools } from './comment-protect';

// 导出单独的定义和处理器，方便测试
export {
  setFormulaDefinition,
  setFormulaHandler,
  generateFormulaDefinition,
  generateFormulaHandler,
  diagnoseFormulaDefinition,
  diagnoseFormulaHandler,
  evaluateFormulaDefinition,
  evaluateFormulaHandler,
  setPrintAreaDefinition,
  setPrintAreaHandler,
  zoomDefinition,
  zoomHandler,
} from './formula';

export {
  readRangeDefinition,
  readRangeHandler,
  writeRangeDefinition,
  writeRangeHandler,
  cleanDataDefinition,
  cleanDataHandler,
  removeDuplicatesDefinition,
  removeDuplicatesHandler,
  sortRangeDefinition,
  sortRangeHandler,
  findReplaceDefinition,
  findReplaceHandler,
  insertRowDefinition,
  insertRowHandler,
  addCommentDefinition,
  addCommentHandler,
  protectSheetDefinition,
  protectSheetHandler,
  setConditionalFormatDefinition,
  setConditionalFormatHandler,
  protectWorkbookDefinition,
  protectWorkbookHandler,
  setZoomDefinition,
  setZoomHandler,
} from './data';

export {
  createPivotTableDefinition,
  createPivotTableHandler,
  updatePivotTableDefinition,
  updatePivotTableHandler,
} from './pivot';

export {
  createChartDefinition,
  createChartHandler,
  updateChartDefinition,
  updateChartHandler,
  exportChartAsImageDefinition,
  exportChartAsImageHandler,
  exportRangeAsImageDefinition,
  exportRangeAsImageHandler,
} from './chart';

export {
  createSheetDefinition,
  createSheetHandler,
  deleteSheetDefinition,
  deleteSheetHandler,
  renameSheetDefinition,
  renameSheetHandler,
  copySheetDefinition,
  copySheetHandler,
  getSheetListDefinition,
  getSheetListHandler,
  switchSheetDefinition,
  switchSheetHandler,
  moveSheetDefinition,
  moveSheetHandler,
  getSelectionDefinition,
  getSelectionHandler,
  deleteRowDefinition,
  deleteRowHandler,
  insertColumnDefinition,
  insertColumnHandler,
  deleteColumnDefinition,
  deleteColumnHandler,
  freezePanesDefinition,
  freezePanesHandler,
  autoFillDefinition,
  autoFillHandler,
  setNamedRangeDefinition,
  setNamedRangeHandler,
  autoSumDefinition,
  autoSumHandler,
  hideColumnDefinition,
  hideColumnHandler,
} from './sheet';

export {
  setCellFormatDefinition,
  setCellFormatHandler,
  setCellStyleDefinition,
  setCellStyleHandler,
  setBorderDefinition,
  setBorderHandler,
  setNumberFormatDefinition,
  setNumberFormatHandler,
  mergeCellsDefinition,
  mergeCellsHandler,
  unmergeCellsDefinition,
  unmergeCellsHandler,
  setColumnWidthDefinition,
  setColumnWidthHandler,
  setRowHeightDefinition,
  setRowHeightHandler,
  hideRowDefinition,
  hideRowHandler,
  setDataValidationDefinition,
  setDataValidationHandler,
} from './format';

export {
  openWorkbookDefinition,
  openWorkbookHandler,
  getOpenWorkbooksDefinition,
  getOpenWorkbooksHandler,
  switchWorkbookDefinition,
  switchWorkbookHandler,
  closeWorkbookDefinition,
  closeWorkbookHandler,
  createWorkbookDefinition,
  createWorkbookHandler,
  getCellValueDefinition,
  getCellValueHandler,
  setCellValueDefinition,
  setCellValueHandler,
  getFormulaDefinition,
  getFormulaHandler,
  getCellInfoDefinition,
  getCellInfoHandler,
  clearRangeDefinition,
  clearRangeHandler,
} from './workbook';

export {
  autoFilterDefinition,
  autoFilterHandler,
  copyRangeDefinition,
  copyRangeHandler,
  pasteRangeDefinition,
  pasteRangeHandler,
  fillSeriesDefinition,
  fillSeriesHandler,
  transposeDefinition,
  transposeHandler,
  textToColumnsDefinition,
  textToColumnsHandler,
  subtotalDefinition,
  subtotalHandler,
} from './data-advanced';

export {
  insertRowsDefinition,
  insertRowsHandler,
  insertColumnsDefinition,
  insertColumnsHandler,
  deleteRowsDefinition,
  deleteRowsHandler,
  deleteColumnsDefinition,
  deleteColumnsHandler,
  hideRowsDefinition,
  hideRowsHandler,
  showRowsDefinition,
  showRowsHandler,
  showColumnsDefinition,
  showColumnsHandler,
  groupRowsDefinition,
  groupRowsHandler,
} from './row-column';

export {
  deleteCellCommentDefinition,
  deleteCellCommentHandler,
  getCellCommentsDefinition,
  getCellCommentsHandler,
  unprotectSheetDefinition,
  unprotectSheetHandler,
  lockCellsDefinition,
  lockCellsHandler,
  setArrayFormulaDefinition,
  setArrayFormulaHandler,
  insertExcelImageDefinition,
  insertExcelImageHandler,
  setHyperlinkDefinition,
  setHyperlinkHandler,
} from './comment-protect';

export default excelTools;

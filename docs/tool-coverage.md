# 工具覆盖矩阵（生成物）

> 由 `node scripts/gen-tool-coverage.mjs` 生成，**请勿手改**；CI 会重新生成并对账。
> 数据来源：`spec/tool-definitions.json`（工具目录）、`spec/advertised.json`（广告面）、
> `mcp/src/tools/**` 的静态解析（工具 → 桥 action，见 `scripts/lib/tool-action-map.mjs`）、
> `test/*.test.mjs` 与 `scripts/e2e.mjs`（工具名出现过 = 覆盖证据）。

## 汇总

| 应用 | 工具数 | 有专门测试 | 仅矩阵 ok/error | 仅矩阵 any | 广告面 |
| --- | ---: | ---: | ---: | ---: | ---: |
| Excel | 132 | 132 | 0 | 0 | 47 |
| Word | 75 | 75 | 0 | 0 | 30 |
| PPT | 89 | 82 | 7 | 0 | 15 |
| 通用 | 7 | 7 | 0 | 0 | 2 |
| 转换 | 2 | 2 | 0 | 0 | 2 |
| 逃生舱 | 1 | 1 | 0 | 0 | 0 |
| 其他 | 4 | 4 | 0 | 0 | 4 |
| **合计** | **310** | **303** | **7** | **0** | **100** |

## 矩阵

「广告」= 每次请求随 tools/list 下发的 100 个工具之一；其余经 `wps_call` / `wps_help` 触达。

| 工具 | 应用 | 桥 action | 广告 | 测试层级 | 名字出现过的文件 |
| --- | --- | --- | :---: | --- | --- |
| `wps_batch` | 其他 | — | ✅ | bespoke | arg-shape-guard.test.mjs, error-contract.test.mjs, p1-surface.test.mjs, +3 |
| `wps_call` | 其他 | — | ✅ | bespoke | advertised-surface.test.mjs, cell-format.test.mjs, close-safety.test.mjs, +35 |
| `wps_common_get_app_info` | 通用 | `getAppInfo` |  | bespoke | word-common-coverage.test.mjs |
| `wps_common_get_selected_text` | 通用 | `getSelectedText` |  | bespoke | word-common-coverage.test.mjs |
| `wps_common_ping` | 通用 | `ping` |  | bespoke | error-contract.test.mjs, open-safety.test.mjs |
| `wps_common_save` | 通用 | — | ✅ | bespoke | cell-format.test.mjs, excel-advanced.test.mjs, excel-contract-fixes.test.mjs, +22 |
| `wps_common_save_as` | 通用 | `saveAs` | ✅ | bespoke | cell-format.test.mjs, excel-advanced.test.mjs, excel-contract-fixes.test.mjs, +22 |
| `wps_common_set_selected_text` | 通用 | `setSelectedText` |  | bespoke | word-common-coverage.test.mjs |
| `wps_common_wire_check` | 通用 | `wireCheck` |  | bespoke | word-common-coverage.test.mjs |
| `wps_convert_format` | 转换 | `convertFormat` | ✅ | bespoke | word-common-coverage.test.mjs |
| `wps_convert_to_pdf` | 转换 | `convertToPDF` | ✅ | bespoke | file-ops.test.mjs, honest-reporting.test.mjs |
| `wps_excel_add_comment` | Excel | `addCellComment` |  | bespoke | destructive-guard.test.mjs, excel-contract-fixes.test.mjs |
| `wps_excel_add_icon_set` | Excel | — |  | bespoke | p1-surface.test.mjs |
| `wps_excel_add_list_row` | Excel | `addListRow` | ✅ | bespoke | excel-list-object.test.mjs |
| `wps_excel_add_pivot_calculated_field` | Excel | — |  | bespoke | p1-surface.test.mjs |
| `wps_excel_add_sparkline` | Excel | `addSparkline` |  | bespoke | destructive-guard.test.mjs, excel-advanced.test.mjs |
| `wps_excel_add_trendline` | Excel | — |  | bespoke | p1-surface.test.mjs |
| `wps_excel_advanced_filter` | Excel | — |  | bespoke | p1-surface.test.mjs |
| `wps_excel_allow_edit_range` | Excel | `allowEditRange` |  | bespoke | p1-surface.test.mjs |
| `wps_excel_auto_filter` | Excel | `autoFilter` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_auto_fit` | Excel | `autoFitAll` | ✅ | bespoke | excel-missing-halves.test.mjs, range-limits.test.mjs |
| `wps_excel_auto_fit_columns` | Excel | `autoFitColumn` |  | bespoke | excel-missing-halves.test.mjs |
| `wps_excel_auto_fit_rows` | Excel | `autoFitRow` |  | bespoke | excel-missing-halves.test.mjs |
| `wps_excel_auto_sum` | Excel | `autoSum` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_calculate` | Excel | `calculateSheet` | ✅ | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_change_link` | Excel | `changeLink` |  | bespoke | p1-surface.test.mjs |
| `wps_excel_clean_data` | Excel | `cleanData` |  | bespoke | excel-coverage.test.mjs, range-limits.test.mjs |
| `wps_excel_clear_formats` | Excel | `clearFormats` | ✅ | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_clear_pivot_table` | Excel | `clearPivotTable` | ✅ | bespoke | excel-advanced.test.mjs |
| `wps_excel_clear_range` | Excel | `clearRange` | ✅ | bespoke | advertised-surface.test.mjs, destructive-guard.test.mjs |
| `wps_excel_clear_sparkline` | Excel | `clearSparkline` |  | bespoke | destructive-guard.test.mjs, excel-advanced.test.mjs |
| `wps_excel_close_workbook` | Excel | `closeWorkbook` | ✅ | bespoke | advertised-surface.test.mjs, close-safety.test.mjs, contract-gaps.test.mjs, +12 |
| `wps_excel_consolidate` | Excel | `consolidate` |  | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_copy_format` | Excel | `copyFormat` | ✅ | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_copy_range` | Excel | `copyRange` | ✅ | bespoke | advertised-surface.test.mjs, excel-contract-fixes.test.mjs |
| `wps_excel_copy_sheet` | Excel | `copySheet` |  | bespoke | sheet-ops.test.mjs |
| `wps_excel_create_chart` | Excel | `createChart` | ✅ | bespoke | contract-gaps.test.mjs, destructive-guard.test.mjs, excel-advanced.test.mjs, +1 |
| `wps_excel_create_list_object` | Excel | `createListObject` | ✅ | bespoke | destructive-guard.test.mjs, excel-list-object.test.mjs |
| `wps_excel_create_pivot_table` | Excel | `createPivotTable` | ✅ | bespoke | excel-advanced.test.mjs, excel-coverage.test.mjs |
| `wps_excel_create_sheet` | Excel | `createSheet` | ✅ | bespoke | destructive-guard.test.mjs, excel-missing-halves-2.test.mjs, excel-page-setup.test.mjs, +2 |
| `wps_excel_create_workbook` | Excel | `createWorkbook` |  | bespoke | advertised-surface.test.mjs, close-safety.test.mjs, contract-gaps.test.mjs, +14 |
| `wps_excel_delete_cell_comment` | Excel | `deleteCellComment` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_delete_chart` | Excel | `deleteChart` |  | bespoke | destructive-guard.test.mjs, excel-advanced.test.mjs |
| `wps_excel_delete_columns` | Excel | `deleteColumns` |  | bespoke | excel-coverage.test.mjs, honest-reporting.test.mjs, merged-tools.test.mjs |
| `wps_excel_delete_list_row` | Excel | `deleteListRow` |  | bespoke | destructive-guard.test.mjs, excel-list-object.test.mjs |
| `wps_excel_delete_named_range` | Excel | `deleteNamedRange` |  | bespoke | excel-missing-halves.test.mjs |
| `wps_excel_delete_rows` | Excel | `deleteRows` | ✅ | bespoke | advertised-surface.test.mjs, destructive-guard.test.mjs, merged-tools.test.mjs |
| `wps_excel_delete_sheet` | Excel | `deleteSheet` |  | bespoke | destructive-guard.test.mjs, sheet-ops.test.mjs |
| `wps_excel_diagnose_formula` | Excel | `diagnoseFormula` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_disconnect_link` | Excel | `disconnectLink` |  | bespoke | p1-surface.test.mjs |
| `wps_excel_evaluate_formula` | Excel | — | ✅ | bespoke | advertised-surface.test.mjs, contract-gaps.test.mjs, excel-coverage.test.mjs |
| `wps_excel_export_chart_as_image` | Excel | `exportChartAsImage` | ✅ | bespoke | contract-gaps.test.mjs, excel-coverage.test.mjs |
| `wps_excel_export_range_as_image` | Excel | `exportRangeAsImage` | ✅ | bespoke | excel-coverage.test.mjs |
| `wps_excel_fill_series` | Excel | `fillSeries` |  | bespoke | excel-contract-fixes.test.mjs, merged-tools.test.mjs |
| `wps_excel_find_in_sheet` | Excel | `findInSheet` | ✅ | bespoke | excel-missing-halves.test.mjs, range-limits.test.mjs |
| `wps_excel_find_replace` | Excel | `findReplaceExcel` | ✅ | bespoke | find-replace.test.mjs |
| `wps_excel_freeze_panes` | Excel | `freezePanes` | ✅ | bespoke | advertised-surface.test.mjs, excel-contract-fixes.test.mjs |
| `wps_excel_generate_formula` | Excel | `getContext` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_get_cell_comments` | Excel | `getCellComments` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_get_cell_info` | Excel | `getCellInfo` | ✅ | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_get_cell_value` | Excel | `getCellValue` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_get_charts` | Excel | `listCharts` |  | bespoke | p1-surface.test.mjs |
| `wps_excel_get_conditional_formats` | Excel | `getConditionalFormats` | ✅ | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_get_data_validations` | Excel | `getDataValidations` |  | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_get_formula` | Excel | `getFormula` |  | bespoke | excel-list-object.test.mjs, excel-page-setup.test.mjs |
| `wps_excel_get_formula_audit` | Excel | `getFormulaAudit` |  | bespoke | excel-page-setup.test.mjs |
| `wps_excel_get_links` | Excel | `getWorkbookLinks` | ✅ | bespoke | advertised-surface.test.mjs, p1-surface.test.mjs |
| `wps_excel_get_list_objects` | Excel | `getListObjects` | ✅ | bespoke | excel-list-object.test.mjs |
| `wps_excel_get_named_ranges` | Excel | `getNamedRanges` | ✅ | bespoke | excel-missing-halves.test.mjs |
| `wps_excel_get_open_workbooks` | Excel | `getOpenWorkbooks` | ✅ | bespoke | advertised-surface.test.mjs, contract-gaps.test.mjs, excel-advanced.test.mjs, +10 |
| `wps_excel_get_pivot_tables` | Excel | `getPivotTables` |  | bespoke | excel-advanced.test.mjs |
| `wps_excel_get_selection` | Excel | `getSelection` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_get_sheet_info` | Excel | `getExcelContext` | ✅ | bespoke | advertised-surface.test.mjs, excel-advanced.test.mjs, excel-list-object.test.mjs, +5 |
| `wps_excel_get_sheet_list` | Excel | `getSheetList` | ✅ | bespoke | excel-coverage.test.mjs |
| `wps_excel_get_sheet_settings` | Excel | `getSheetSettings` | ✅ | bespoke | excel-coverage.test.mjs, excel-page-setup.test.mjs |
| `wps_excel_goal_seek` | Excel | `goalSeek` | ✅ | bespoke | excel-advanced.test.mjs |
| `wps_excel_group_columns` | Excel | `groupColumns` |  | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_group_pivot_dates` | Excel | — |  | bespoke | p1-surface.test.mjs |
| `wps_excel_group_rows` | Excel | `groupRows` |  | bespoke | excel-page-setup.test.mjs |
| `wps_excel_hide_column` | Excel | `hideColumns` |  | bespoke | excel-contract-fixes.test.mjs |
| `wps_excel_hide_rows` | Excel | `hideRows` |  | bespoke | excel-contract-fixes.test.mjs, merged-tools.test.mjs |
| `wps_excel_import_csv` | Excel | `importCsv` |  | bespoke | p1-surface.test.mjs |
| `wps_excel_insert_columns` | Excel | `insertColumns` |  | bespoke | excel-coverage.test.mjs, merged-tools.test.mjs, range-limits.test.mjs |
| `wps_excel_insert_excel_image` | Excel | `insertExcelImage` |  | bespoke | excel-contract-fixes.test.mjs, range-limits.test.mjs |
| `wps_excel_insert_page_break` | Excel | `insertExcelPageBreak` |  | bespoke | p1-surface.test.mjs |
| `wps_excel_insert_rows` | Excel | `insertRows` |  | bespoke | excel-coverage.test.mjs, merged-tools.test.mjs |
| `wps_excel_lock_cells` | Excel | `lockCells` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_merge_cells` | Excel | `mergeCells` |  | bespoke | advertised-surface.test.mjs, excel-coverage.test.mjs |
| `wps_excel_move_sheet` | Excel | `moveSheet` |  | bespoke | sheet-ops.test.mjs |
| `wps_excel_open_workbook` | Excel | `openWorkbook` | ✅ | bespoke | encrypted-preflight.test.mjs, file-ops.test.mjs, open-safety.test.mjs |
| `wps_excel_paste_range` | Excel | `pasteRange` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_protect_sheet` | Excel | `protectSheet` |  | bespoke | excel-contract-fixes.test.mjs, excel-coverage.test.mjs, honest-reporting.test.mjs |
| `wps_excel_protect_workbook` | Excel | `protectWorkbook` |  | bespoke | excel-contract-fixes.test.mjs |
| `wps_excel_read_range` | Excel | `getRangeData` | ✅ | bespoke | advertised-surface.test.mjs, arg-shape-guard.test.mjs, close-safety.test.mjs, +11 |
| `wps_excel_refresh_all_data` | Excel | `refreshAllData` |  | bespoke | excel-advanced.test.mjs |
| `wps_excel_refresh_links` | Excel | `refreshLinks` |  | bespoke | excel-missing-halves-2.test.mjs, honest-reporting.test.mjs |
| `wps_excel_refresh_pivot_tables` | Excel | `refreshPivotTables` |  | bespoke | excel-advanced.test.mjs |
| `wps_excel_remove_conditional_format` | Excel | `removeConditionalFormat` |  | bespoke | destructive-guard.test.mjs, excel-missing-halves-2.test.mjs |
| `wps_excel_remove_data_validation` | Excel | `removeDataValidation` |  | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_remove_duplicates` | Excel | `removeDuplicates` |  | bespoke | excel-coverage.test.mjs, excel-remove-duplicates.test.mjs |
| `wps_excel_rename_sheet` | Excel | `renameSheet` |  | bespoke | sheet-ops.test.mjs |
| `wps_excel_reset_page_breaks` | Excel | `resetPageBreaks` |  | bespoke | excel-page-setup.test.mjs |
| `wps_excel_resize_list_object` | Excel | `resizeListObject` |  | bespoke | excel-list-object.test.mjs |
| `wps_excel_set_array_formula` | Excel | `setArrayFormula` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_set_border` | Excel | `setBorder` |  | bespoke | excel-contract-fixes.test.mjs |
| `wps_excel_set_cell_format` | Excel | `setCellFormat` | ✅ | bespoke | cell-format.test.mjs, destructive-guard.test.mjs, excel-missing-halves-2.test.mjs |
| `wps_excel_set_cell_style` | Excel | `setCellStyle` |  | bespoke | excel-contract-fixes.test.mjs |
| `wps_excel_set_cell_value` | Excel | `setCellValue` |  | bespoke | arg-shape-guard.test.mjs, excel-coverage.test.mjs |
| `wps_excel_set_chart_axis` | Excel | `setChartAxis` |  | bespoke | p1-surface.test.mjs |
| `wps_excel_set_chart_labels` | Excel | `setChartLabels` | ✅ | bespoke | excel-advanced.test.mjs |
| `wps_excel_set_column_width` | Excel | `setColumnWidth` | ✅ | bespoke | advertised-surface.test.mjs, excel-coverage.test.mjs |
| `wps_excel_set_conditional_format` | Excel | `addConditionalFormat` | ✅ | bespoke | destructive-guard.test.mjs, excel-contract-fixes.test.mjs, excel-missing-halves-2.test.mjs |
| `wps_excel_set_data_validation` | Excel | `addDataValidation` |  | bespoke | destructive-guard.test.mjs, excel-contract-fixes.test.mjs, excel-missing-halves-2.test.mjs |
| `wps_excel_set_formula` | Excel | `setFormula` | ✅ | bespoke | excel-advanced.test.mjs, excel-page-setup.test.mjs, word-lifecycle.test.mjs |
| `wps_excel_set_hyperlink` | Excel | `setHyperlink` |  | bespoke | excel-contract-fixes.test.mjs |
| `wps_excel_set_list_object_totals` | Excel | `setListObjectTotals` |  | bespoke | excel-list-object.test.mjs |
| `wps_excel_set_named_range` | Excel | `createNamedRange` |  | bespoke | destructive-guard.test.mjs, excel-missing-halves.test.mjs |
| `wps_excel_set_number_format` | Excel | `setNumberFormat` | ✅ | bespoke | sheet-ops.test.mjs |
| `wps_excel_set_outline_levels` | Excel | `setOutlineLevels` |  | bespoke | excel-page-setup.test.mjs |
| `wps_excel_set_print_area` | Excel | — |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_set_row_height` | Excel | `setRowHeight` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_set_series_axis` | Excel | — |  | bespoke | p1-surface.test.mjs |
| `wps_excel_set_sheet_appearance` | Excel | `setSheetAppearance` | ✅ | bespoke | advertised-surface.test.mjs, excel-page-setup.test.mjs |
| `wps_excel_set_sheet_header_footer` | Excel | `setSheetHeaderFooter` | ✅ | bespoke | excel-page-setup.test.mjs |
| `wps_excel_set_sheet_page_setup` | Excel | `setSheetPageSetup` | ✅ | bespoke | excel-page-setup.test.mjs |
| `wps_excel_set_sheet_print_titles` | Excel | `setSheetPrintTitles` | ✅ | bespoke | excel-page-setup.test.mjs |
| `wps_excel_set_wrap_text` | Excel | `wrapText` |  | bespoke | excel-missing-halves.test.mjs |
| `wps_excel_set_zoom` | Excel | `setZoom` |  | bespoke | deprecated.test.mjs |
| `wps_excel_show_columns` | Excel | `showColumns` |  | bespoke | excel-contract-fixes.test.mjs |
| `wps_excel_show_rows` | Excel | `showRows` |  | bespoke | excel-contract-fixes.test.mjs |
| `wps_excel_sort_range` | Excel | `sortRange` | ✅ | bespoke | excel-contract-fixes.test.mjs, p1-surface.test.mjs |
| `wps_excel_subtotal` | Excel | `subtotal` |  | bespoke | excel-missing-halves-2.test.mjs |
| `wps_excel_switch_sheet` | Excel | `switchSheet` | ✅ | bespoke | destructive-guard.test.mjs, excel-missing-halves-2.test.mjs, sheet-ops.test.mjs |
| `wps_excel_switch_workbook` | Excel | `switchWorkbook` | ✅ | bespoke | excel-coverage.test.mjs |
| `wps_excel_text_to_columns` | Excel | — |  | bespoke | arg-shape-guard.test.mjs, excel-coverage.test.mjs |
| `wps_excel_transpose` | Excel | `transpose` |  | bespoke | excel-contract-fixes.test.mjs |
| `wps_excel_unlist_list_object` | Excel | `unlistListObject` |  | bespoke | excel-list-object.test.mjs |
| `wps_excel_unmerge_cells` | Excel | `unmergeCells` | ✅ | bespoke | advertised-surface.test.mjs, excel-coverage.test.mjs |
| `wps_excel_unprotect_sheet` | Excel | `unprotectSheet` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_update_chart` | Excel | `updateChart` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_update_list_object` | Excel | `updateListObject` |  | bespoke | excel-list-object.test.mjs |
| `wps_excel_update_pivot_table` | Excel | `updatePivotTable` |  | bespoke | excel-coverage.test.mjs |
| `wps_excel_write_range` | Excel | `setRangeData` | ✅ | bespoke | advertised-surface.test.mjs, arg-shape-guard.test.mjs, cell-format.test.mjs, +19 |
| `wps_execute_method` | 逃生舱 | — |  | bespoke | cell-format.test.mjs, close-safety.test.mjs, contract-gaps.test.mjs, +18 |
| `wps_help` | 其他 | — | ✅ | bespoke | arg-shape-guard.test.mjs, deprecated.test.mjs, merged-tools.test.mjs, +4 |
| `wps_ppt_add_animation` | PPT | `setAnimation` |  | bespoke | destructive-guard.test.mjs, merged-tools.test.mjs, ppt-contract-fixes.test.mjs, +2 |
| `wps_ppt_add_master_element` | PPT | `addMasterElement` |  | matrix ok | ppt-coverage.test.mjs |
| `wps_ppt_add_ppt_hyperlink` | PPT | `addPptHyperlink` |  | bespoke | ppt-contract-fixes.test.mjs |
| `wps_ppt_add_section` | PPT | `addSection` |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_add_shape` | PPT | `addShape` |  | bespoke | destructive-guard.test.mjs, ppt-contract-fixes.test.mjs, ppt-coverage.test.mjs, +1 |
| `wps_ppt_add_slide` | PPT | `addSlide` | ✅ | bespoke | close-safety.test.mjs, contract-gaps.test.mjs, destructive-guard.test.mjs, +8 |
| `wps_ppt_add_textbox` | PPT | `addTextBox` | ✅ | bespoke | destructive-guard.test.mjs, merged-tools.test.mjs, ppt-contract-fixes.test.mjs, +1 |
| `wps_ppt_align_shapes` | PPT | `alignShapes` |  | bespoke | merged-tools.test.mjs |
| `wps_ppt_apply_transition_to_all` | PPT | `applyTransitionToAll` |  | bespoke | ppt-contract-fixes.test.mjs |
| `wps_ppt_beautify` | PPT | `beautifySlide` |  | matrix ok | ppt-coverage.test.mjs |
| `wps_ppt_close_presentation` | PPT | `closePresentation` |  | bespoke | close-safety.test.mjs, contract-gaps.test.mjs, honest-reporting.test.mjs, +3 |
| `wps_ppt_compress_images` | PPT | — |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_copy_slide` | PPT | `duplicateSlide` |  | bespoke | merged-tools.test.mjs, ppt-contract-fixes.test.mjs |
| `wps_ppt_create_presentation` | PPT | `createPresentation` |  | bespoke | close-safety.test.mjs, contract-gaps.test.mjs, destructive-guard.test.mjs, +8 |
| `wps_ppt_delete_ppt_image` | PPT | `deletePptImage` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_delete_shape` | PPT | `deleteShape` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_delete_slide` | PPT | `deleteSlide` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_delete_textbox` | PPT | `deleteTextBox` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_distribute_shapes` | PPT | `distributeShapes` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_duplicate_master_layout` | PPT | — |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_duplicate_shape` | PPT | `duplicateShape` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_embed_fonts` | PPT | — |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_end_slide_show` | PPT | `endSlideShow` |  | bespoke | ppt-contract-fixes.test.mjs |
| `wps_ppt_export_notes` | PPT | `getSlideNotes` | ✅ | bespoke | advertised-surface.test.mjs, p1-surface.test.mjs |
| `wps_ppt_export_pdf` | PPT | — |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_export_slide_as_image` | PPT | `exportSlideAsImage` | ✅ | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_find_ppt_text` | PPT | `findPptText` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_get_animations` | PPT | `getAnimations` |  | matrix ok | ppt-coverage.test.mjs |
| `wps_ppt_get_fonts` | PPT | — |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_get_hidden_slides` | PPT | `getHiddenSlides` |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_get_open_presentations` | PPT | `getOpenPresentations` |  | bespoke | honest-reporting.test.mjs, ppt-coverage.test.mjs, ppt-slimming.test.mjs, +1 |
| `wps_ppt_get_shapes` | PPT | `getShapes` | ✅ | bespoke | contract-gaps.test.mjs, ppt-coverage.test.mjs, ppt-slide-index.test.mjs |
| `wps_ppt_get_slide_count` | PPT | `getSlideCount` | ✅ | bespoke | ppt-coverage.test.mjs, ppt-slide-index.test.mjs |
| `wps_ppt_get_slide_info` | PPT | `getSlideInfo` | ✅ | bespoke | ppt-coverage.test.mjs, ppt-slide-index.test.mjs, target-identity.test.mjs |
| `wps_ppt_get_slide_master` | PPT | `getSlideMaster` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_get_slide_notes` | PPT | — |  | bespoke | p1-surface.test.mjs, ppt-coverage.test.mjs |
| `wps_ppt_get_slide_title` | PPT | `getSlideTitle` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_get_table_cell` | PPT | `getPptTableCell` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_get_textboxes` | PPT | `getTextBoxes` |  | matrix ok | ppt-coverage.test.mjs |
| `wps_ppt_group_shapes` | PPT | `groupShapes` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_insert_ppt_chart` | PPT | `insertPptChart` |  | bespoke | merged-tools.test.mjs, ppt-coverage.test.mjs |
| `wps_ppt_insert_ppt_image` | PPT | `insertPptImage` | ✅ | bespoke | contract-gaps.test.mjs, destructive-guard.test.mjs, merged-tools.test.mjs, +3 |
| `wps_ppt_insert_slides_from_file` | PPT | `insertSlidesFromFile` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_insert_table` | PPT | `insertPptTable` | ✅ | bespoke | p1-surface.test.mjs, ppt-coverage.test.mjs, ppt-shape-position.test.mjs, +1 |
| `wps_ppt_move_slide` | PPT | `moveSlide` |  | bespoke | ppt-coverage.test.mjs, ppt-slide-index.test.mjs |
| `wps_ppt_open_presentation` | PPT | `openPresentation` | ✅ | bespoke | encrypted-preflight.test.mjs, ppt-coverage.test.mjs |
| `wps_ppt_remove_animation` | PPT | `removeAnimation` |  | bespoke | ppt-contract-fixes.test.mjs |
| `wps_ppt_remove_ppt_hyperlink` | PPT | `removePptHyperlink` |  | matrix ok | ppt-coverage.test.mjs |
| `wps_ppt_remove_slide_transition` | PPT | `removeSlideTransition` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_replace_ppt_image` | PPT | `replacePptImage` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_replace_ppt_text` | PPT | `replacePptText` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_active_target` | PPT | `getOpenPresentations` |  | bespoke | ppt-coverage.test.mjs, word-range-format.test.mjs |
| `wps_ppt_set_animation_order` | PPT | `setAnimationOrder` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_background_color` | PPT | `setBackgroundColor` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_background_gradient` | PPT | `setBackgroundGradient` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_background_image` | PPT | `setBackgroundImage` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_font_color` | PPT | `setFontColor` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_image_style` | PPT | `setImageStyle` |  | bespoke | ppt-contract-fixes.test.mjs, ppt-slide-index.test.mjs |
| `wps_ppt_set_master_background` | PPT | `setMasterBackground` |  | bespoke | ppt-contract-fixes.test.mjs |
| `wps_ppt_set_placeholder_lock` | PPT | — |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_set_ppt_chart_data` | PPT | `setPptChartData` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_ppt_chart_style` | PPT | `setPptChartStyle` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_shape_effect` | PPT | `setShapeEffect` |  | bespoke | ppt-contract-fixes.test.mjs, ppt-slimming.test.mjs |
| `wps_ppt_set_shape_fill` | PPT | `setShapeFill` | ✅ | bespoke | ppt-contract-fixes.test.mjs |
| `wps_ppt_set_shape_position` | PPT | `setShapePosition` |  | bespoke | ppt-coverage.test.mjs, ppt-shape-position.test.mjs, ppt-slide-index.test.mjs |
| `wps_ppt_set_shape_style` | PPT | `setShapeStyle` |  | bespoke | ppt-contract-fixes.test.mjs, ppt-coverage.test.mjs |
| `wps_ppt_set_shape_text` | PPT | `setShapeText` | ✅ | bespoke | ppt-contract-fixes.test.mjs |
| `wps_ppt_set_shape_z_order` | PPT | `setShapeZOrder` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_background` | PPT | `setSlideBackground` |  | bespoke | arg-shape-guard.test.mjs, merged-tools.test.mjs, ppt-contract-fixes.test.mjs |
| `wps_ppt_set_slide_content` | PPT | `setSlideContent` | ✅ | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_footer` | PPT | `setSlideFooter` |  | bespoke | ppt-contract-fixes.test.mjs, ppt-slimming.test.mjs |
| `wps_ppt_set_slide_hidden` | PPT | `setSlideHidden` |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_set_slide_layout` | PPT | `setSlideLayout` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_notes` | PPT | `setSlideNotes` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_size` | PPT | `setSlideSize` |  | bespoke | p1-surface.test.mjs, ppt-coverage.test.mjs |
| `wps_ppt_set_slide_subtitle` | PPT | `setSlideSubtitle` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_theme` | PPT | `setSlideTheme` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_title` | PPT | `setSlideTitle` | ✅ | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_transition` | PPT | `setSlideTransition` |  | bespoke | merged-tools.test.mjs, ppt-contract-fixes.test.mjs |
| `wps_ppt_set_table_cell` | PPT | `setPptTableCell` |  | bespoke | ppt-coverage.test.mjs, ppt-slimming.test.mjs |
| `wps_ppt_set_table_format` | PPT | `setPptTableFormat` | ✅ | bespoke | ppt-shape-position.test.mjs, ppt-slimming.test.mjs |
| `wps_ppt_set_textbox_style` | PPT | `setTextBoxStyle` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_set_textbox_text` | PPT | `setTextBoxText` |  | matrix ok | ppt-coverage.test.mjs |
| `wps_ppt_set_theme_colors` | PPT | — |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_set_theme_fonts` | PPT | `setThemeFonts` |  | bespoke | p1-surface.test.mjs |
| `wps_ppt_start_slide_show` | PPT | `startSlideShow` |  | bespoke | ppt-contract-fixes.test.mjs |
| `wps_ppt_switch_presentation` | PPT | `switchPresentation` |  | bespoke | ppt-coverage.test.mjs |
| `wps_ppt_switch_slide` | PPT | `switchSlide` |  | matrix ok | ppt-coverage.test.mjs |
| `wps_ppt_unify_font` | PPT | `unifyFont` |  | bespoke | ppt-contract-fixes.test.mjs |
| `wps_status` | 其他 | — | ✅ | bespoke | advertised-surface.test.mjs, deprecated.test.mjs, skill-quality.test.mjs, +2 |
| `wps_word_accept_revisions` | Word | `acceptRevisions` |  | bespoke | word-produce.test.mjs |
| `wps_word_add_content_control` | Word | `addContentControl` |  | bespoke | word-longtail.test.mjs |
| `wps_word_add_endnote` | Word | `addEndnote` |  | bespoke | word-longtail.test.mjs |
| `wps_word_add_footnote` | Word | `addFootnote` |  | bespoke | word-longtail.test.mjs |
| `wps_word_add_table_lines` | Word | `addTableLines` |  | bespoke | word-deep.test.mjs |
| `wps_word_apply_style` | Word | `applyStyle` | ✅ | bespoke | word-common-coverage.test.mjs, word-page-setup-gaps.test.mjs, word-range-format.test.mjs |
| `wps_word_close_document` | Word | `closeDocument` | ✅ | bespoke | advertised-surface.test.mjs, close-safety.test.mjs, contract-gaps.test.mjs, +9 |
| `wps_word_compare_documents` | Word | `compareDocuments` | ✅ | bespoke | advertised-surface.test.mjs, p1-surface.test.mjs |
| `wps_word_convert_table_to_text` | Word | `convertTableToText` |  | bespoke | word-deep.test.mjs |
| `wps_word_create_document` | Word | `createDocument` | ✅ | bespoke | advertised-surface.test.mjs, close-safety.test.mjs, contract-gaps.test.mjs, +14 |
| `wps_word_create_envelope` | Word | — |  | bespoke | p1-surface.test.mjs |
| `wps_word_create_labels` | Word | — |  | bespoke | p1-surface.test.mjs |
| `wps_word_delete_comment` | Word | `deleteComment` |  | bespoke | word-produce.test.mjs |
| `wps_word_delete_table_line` | Word | `deleteTableLine` |  | bespoke | destructive-guard.test.mjs, word-deep.test.mjs |
| `wps_word_enable_track_changes` | Word | `enableTrackChanges` |  | bespoke | word-produce.test.mjs |
| `wps_word_find_in_document` | Word | `findInDocument` |  | bespoke | word-common-coverage.test.mjs |
| `wps_word_find_replace` | Word | `findReplace` | ✅ | bespoke | find-replace.test.mjs, word-range-format.test.mjs |
| `wps_word_generate_toc` | Word | `generateTOC` | ✅ | bespoke | word-common-coverage.test.mjs, word-page-setup-gaps.test.mjs |
| `wps_word_get_active_document` | Word | `getActiveDocument` | ✅ | bespoke | advertised-surface.test.mjs, target-identity.test.mjs, word-lifecycle.test.mjs |
| `wps_word_get_bookmarks` | Word | `getBookmarks` |  | bespoke | word-deep.test.mjs |
| `wps_word_get_comments` | Word | `getComments` | ✅ | bespoke | word-deep.test.mjs, word-produce.test.mjs |
| `wps_word_get_content_controls` | Word | `getContentControls` |  | bespoke | word-longtail.test.mjs |
| `wps_word_get_document_stats` | Word | `getDocumentStats` | ✅ | bespoke | word-deep.test.mjs |
| `wps_word_get_document_text` | Word | `getDocumentText` | ✅ | bespoke | advertised-surface.test.mjs, close-safety.test.mjs, excel-contract-fixes.test.mjs, +5 |
| `wps_word_get_fields` | Word | `getDocumentFields` | ✅ | bespoke | advertised-surface.test.mjs, p1-surface.test.mjs |
| `wps_word_get_header_footer` | Word | `getHeaderFooter` |  | bespoke | word-page-setup-gaps.test.mjs |
| `wps_word_get_notes` | Word | `getNotes` | ✅ | bespoke | word-longtail.test.mjs |
| `wps_word_get_open_documents` | Word | `getOpenDocuments` | ✅ | bespoke | advertised-surface.test.mjs, warning-channel.test.mjs, word-deep.test.mjs, +5 |
| `wps_word_get_page_setup` | Word | `getPageSetup` | ✅ | bespoke | word-page-setup-gaps.test.mjs, word-section-layout.test.mjs |
| `wps_word_get_paragraphs` | Word | `getDocumentParagraphs` | ✅ | bespoke | contract-gaps.test.mjs, fault-injection.test.mjs, word-common-coverage.test.mjs, +2 |
| `wps_word_get_protection` | Word | `getDocumentProtection` |  | bespoke | p1-surface.test.mjs |
| `wps_word_get_revisions` | Word | `getRevisions` | ✅ | bespoke | word-produce.test.mjs |
| `wps_word_get_table_data` | Word | `getTableData` | ✅ | bespoke | word-deep.test.mjs |
| `wps_word_get_tables` | Word | `getDocumentTables` | ✅ | bespoke | word-deep.test.mjs |
| `wps_word_get_track_changes_status` | Word | `getTrackChangesStatus` |  | bespoke | word-common-coverage.test.mjs |
| `wps_word_insert_bookmark` | Word | `insertBookmark` |  | bespoke | word-common-coverage.test.mjs, word-deep.test.mjs, word-longtail.test.mjs |
| `wps_word_insert_caption` | Word | `insertCaption` |  | bespoke | p1-surface.test.mjs |
| `wps_word_insert_comment` | Word | `addComment` |  | bespoke | contract-gaps.test.mjs, word-deep.test.mjs, word-produce.test.mjs |
| `wps_word_insert_cross_reference` | Word | `insertCrossReference` |  | bespoke | word-longtail.test.mjs |
| `wps_word_insert_footer` | Word | `insertFooter` |  | bespoke | file-ops.test.mjs |
| `wps_word_insert_header` | Word | `insertHeader` |  | bespoke | file-ops.test.mjs, p1-surface.test.mjs |
| `wps_word_insert_hyperlink` | Word | `insertHyperlink` |  | bespoke | word-deep.test.mjs |
| `wps_word_insert_image` | Word | `insertImage` |  | bespoke | word-common-coverage.test.mjs |
| `wps_word_insert_index` | Word | `insertIndex` |  | bespoke | word-longtail.test.mjs |
| `wps_word_insert_page_break` | Word | `insertPageBreak` |  | bespoke | word-common-coverage.test.mjs, word-page-setup-gaps.test.mjs |
| `wps_word_insert_page_numbers` | Word | `insertPageNumbers` | ✅ | bespoke | word-page-setup-gaps.test.mjs, word-produce.test.mjs, word-section-layout.test.mjs |
| `wps_word_insert_section_break` | Word | `insertSectionBreak` |  | bespoke | word-common-coverage.test.mjs, word-section-layout.test.mjs |
| `wps_word_insert_table` | Word | `insertTable` |  | bespoke | destructive-guard.test.mjs, p1-surface.test.mjs, word-deep.test.mjs, +1 |
| `wps_word_insert_table_of_figures` | Word | `insertTableOfFigures` |  | bespoke | p1-surface.test.mjs |
| `wps_word_insert_text` | Word | `insertText` | ✅ | bespoke | advertised-surface.test.mjs, close-safety.test.mjs, contract-gaps.test.mjs, +15 |
| `wps_word_mail_merge` | Word | `mailMerge` | ✅ | bespoke | word-longtail.test.mjs |
| `wps_word_merge_table_cells` | Word | `mergeTableCells` |  | bespoke | word-deep.test.mjs |
| `wps_word_open_document` | Word | `openDocument` | ✅ | bespoke | encrypted-preflight.test.mjs, open-safety.test.mjs, word-longtail.test.mjs |
| `wps_word_proofread_basic` | Word | — |  | bespoke | word-common-coverage.test.mjs |
| `wps_word_protect` | Word | `protectDocument` | ✅ | bespoke | advertised-surface.test.mjs, p1-surface.test.mjs |
| `wps_word_reject_revisions` | Word | `rejectRevisions` |  | bespoke | word-produce.test.mjs |
| `wps_word_replace_bookmark_content` | Word | `replaceBookmarkContent` |  | bespoke | word-common-coverage.test.mjs |
| `wps_word_replace_range` | Word | `replaceRange` | ✅ | bespoke | advertised-surface.test.mjs, word-common-coverage.test.mjs, word-range-format.test.mjs |
| `wps_word_resolve_revisions` | Word | `resolveRevisions` |  | bespoke | p1-surface.test.mjs |
| `wps_word_select_range` | Word | `selectDocumentRange` |  | bespoke | p1-surface.test.mjs |
| `wps_word_set_columns` | Word | `setColumns` |  | bespoke | word-produce.test.mjs |
| `wps_word_set_content_control` | Word | `setContentControl` |  | bespoke | p1-surface.test.mjs |
| `wps_word_set_font` | Word | `setFont` | ✅ | bespoke | arg-shape-guard.test.mjs, deprecated.test.mjs, fault-injection.test.mjs, +4 |
| `wps_word_set_line_spacing` | Word | `setLineSpacing` |  | bespoke | word-common-coverage.test.mjs |
| `wps_word_set_page_setup` | Word | `setPageSetup` | ✅ | bespoke | word-common-coverage.test.mjs, word-page-setup-gaps.test.mjs, word-section-layout.test.mjs |
| `wps_word_set_paragraph` | Word | — | ✅ | bespoke | contract-gaps.test.mjs, fault-injection.test.mjs, word-common-coverage.test.mjs, +1 |
| `wps_word_set_table_cell` | Word | `setTableCell` | ✅ | bespoke | word-deep.test.mjs |
| `wps_word_set_table_format` | Word | `setTableFormat` |  | bespoke | word-deep.test.mjs, word-section-layout.test.mjs |
| `wps_word_set_table_formula` | Word | `setTableFormula` |  | bespoke | p1-surface.test.mjs |
| `wps_word_set_text_color` | Word | `setTextColor` |  | bespoke | word-common-coverage.test.mjs |
| `wps_word_smart_fill_field` | Word | `smartFillField` | ✅ | bespoke | word-common-coverage.test.mjs |
| `wps_word_split_table_cell` | Word | `splitTableCell` |  | bespoke | word-deep.test.mjs |
| `wps_word_switch_document` | Word | `switchDocument` | ✅ | bespoke | word-common-coverage.test.mjs |
| `wps_word_update_fields` | Word | — |  | bespoke | p1-surface.test.mjs |
| `wps_word_update_style` | Word | — |  | bespoke | p1-surface.test.mjs |

## 没有任何测试点到（未驱动，0）

无——310 个工具都至少被一个测试或 e2e 在代码里点到（注释不算）。

## 解析不出桥 action 的工具（29）

- `wps_batch`
- `wps_call`
- `wps_common_save`
- `wps_excel_add_icon_set`
- `wps_excel_add_pivot_calculated_field`
- `wps_excel_add_trendline`
- `wps_excel_advanced_filter`
- `wps_excel_evaluate_formula`
- `wps_excel_group_pivot_dates`
- `wps_excel_set_print_area`
- `wps_excel_set_series_axis`
- `wps_excel_text_to_columns`
- `wps_execute_method`
- `wps_help`
- `wps_ppt_compress_images`
- `wps_ppt_duplicate_master_layout`
- `wps_ppt_embed_fonts`
- `wps_ppt_export_pdf`
- `wps_ppt_get_fonts`
- `wps_ppt_get_slide_notes`
- `wps_ppt_set_placeholder_lock`
- `wps_ppt_set_theme_colors`
- `wps_status`
- `wps_word_create_envelope`
- `wps_word_create_labels`
- `wps_word_proofread_basic`
- `wps_word_set_paragraph`
- `wps_word_update_fields`
- `wps_word_update_style`

## 已知 WPS 不支持（实测，不再尝试）

以下不是「没做」，是 COM 层面不成立（证据见 `docs/FIXES.md` 与探针脚本），推广材料应写「不支持」：

- Word 水印：页眉 `Shapes` 拒绝一切添加，`Count` 恒为 0。
- 文档属性：`BuiltInDocumentProperties` / `CustomDocumentProperties` 是坏壳。
- Excel 切片器：`SlicerCaches.Add2` 可调用，但 `Slicers.Count` 恒为 0。
- 方案管理器：`Worksheet.Scenarios` 在 COM 里是方法，语义读不干净。

## 说明

- 测试层级由 `scripts/lib/coverage-tiers.mjs` 给出（唯一口径）：**bespoke** = 有非矩阵条目的代码点到它；**matrix ok/error** = 只在矩阵里、至少断言了成功或明确失败；**matrix any** = 只在矩阵里、只断言「没挂住」（最弱，棘轮只许降）。
- 「名字出现过的文件」只是定位用的索引（按字符串出现统计，含注释），不代表行为被断言。
- 棘轮断言在 `test/spec-reproduction.test.mjs`；矩阵与实时冒烟也可用 `node scripts/smoke-tools.mjs` 查看。

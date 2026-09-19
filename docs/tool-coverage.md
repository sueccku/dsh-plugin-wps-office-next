# 工具覆盖矩阵（生成物）

> 由 `node scripts/gen-tool-coverage.mjs` 生成，**请勿手改**；CI 会重新生成并对账。
> 数据来源：`spec/tool-definitions.json`（工具目录）、`spec/advertised.json`（广告面）、
> `mcp/src/tools/**` 的静态解析（工具 → 桥 action，见 `scripts/lib/tool-action-map.mjs`）、
> `test/*.test.mjs` 与 `scripts/e2e.mjs`（工具名出现过 = 覆盖证据）。

## 汇总

| 应用 | 工具数 | 被测试点名 | 广告面 |
| --- | ---: | ---: | ---: |
| Excel | 118 | 118 | 26 |
| Word | 59 | 59 | 21 |
| PPT | 76 | 76 | 14 |
| 通用 | 7 | 7 | 2 |
| 转换 | 2 | 2 | 2 |
| 其他 | 4 | 4 | 4 |
| **合计** | **267** | **267** | **69** |

## 矩阵

「广告」= 每次请求随 tools/list 下发的 69 个工具之一；其余经 `wps_call` / `wps_help` 触达。

| 工具 | 应用 | 桥 action | 广告 | 被测试点名（证据） |
| --- | --- | --- | :---: | --- |
| `wps_batch` | 其他 | — | ✅ | arg-shape-guard.test.mjs, error-contract.test.mjs, e2e |
| `wps_call` | 其他 | — | ✅ | cell-format.test.mjs, close-safety.test.mjs, deprecated.test.mjs, +24 |
| `wps_common_get_app_info` | 通用 | `getAppInfo` |  | word-common-coverage.test.mjs |
| `wps_common_get_selected_text` | 通用 | `getSelectedText` |  | word-common-coverage.test.mjs |
| `wps_common_ping` | 通用 | `ping` |  | error-contract.test.mjs, open-safety.test.mjs |
| `wps_common_save` | 通用 | `save` | ✅ | file-ops.test.mjs, word-common-coverage.test.mjs |
| `wps_common_save_as` | 通用 | `saveAs` | ✅ | file-ops.test.mjs, word-common-coverage.test.mjs |
| `wps_common_set_selected_text` | 通用 | `setSelectedText` |  | word-common-coverage.test.mjs |
| `wps_common_wire_check` | 通用 | `wireCheck` |  | word-common-coverage.test.mjs |
| `wps_convert_format` | 转换 | `convertFormat` | ✅ | word-common-coverage.test.mjs |
| `wps_convert_to_pdf` | 转换 | `convertToPDF` | ✅ | file-ops.test.mjs |
| `wps_excel_add_comment` | Excel | `addCellComment` |  | destructive-guard.test.mjs, excel-contract-fixes.test.mjs |
| `wps_excel_add_list_row` | Excel | `addListRow` | ✅ | excel-list-object.test.mjs |
| `wps_excel_add_sparkline` | Excel | `addSparkline` |  | destructive-guard.test.mjs, excel-advanced.test.mjs |
| `wps_excel_auto_filter` | Excel | `autoFilter` |  | excel-coverage.test.mjs |
| `wps_excel_auto_fit` | Excel | `autoFitAll` | ✅ | excel-missing-halves.test.mjs |
| `wps_excel_auto_fit_columns` | Excel | `autoFitColumn` |  | excel-missing-halves.test.mjs |
| `wps_excel_auto_fit_rows` | Excel | `autoFitRow` |  | excel-missing-halves.test.mjs |
| `wps_excel_auto_sum` | Excel | `autoSum` |  | excel-coverage.test.mjs |
| `wps_excel_calculate` | Excel | `calculateSheet` | ✅ | excel-missing-halves-2.test.mjs |
| `wps_excel_clean_data` | Excel | `cleanData` |  | excel-coverage.test.mjs |
| `wps_excel_clear_formats` | Excel | `clearFormats` | ✅ | excel-missing-halves-2.test.mjs |
| `wps_excel_clear_pivot_table` | Excel | `clearPivotTable` | ✅ | excel-advanced.test.mjs |
| `wps_excel_clear_range` | Excel | `clearRange` |  | destructive-guard.test.mjs |
| `wps_excel_clear_sparkline` | Excel | `clearSparkline` |  | destructive-guard.test.mjs, excel-advanced.test.mjs |
| `wps_excel_close_workbook` | Excel | `closeWorkbook` |  | close-safety.test.mjs, deprecated.test.mjs, excel-advanced.test.mjs, +5 |
| `wps_excel_consolidate` | Excel | `consolidate` |  | excel-missing-halves-2.test.mjs |
| `wps_excel_copy_format` | Excel | `copyFormat` | ✅ | excel-missing-halves-2.test.mjs |
| `wps_excel_copy_range` | Excel | `copyRange` |  | excel-contract-fixes.test.mjs |
| `wps_excel_copy_sheet` | Excel | `copySheet` |  | sheet-ops.test.mjs |
| `wps_excel_create_chart` | Excel | `createChart` | ✅ | destructive-guard.test.mjs, excel-advanced.test.mjs, excel-coverage.test.mjs |
| `wps_excel_create_list_object` | Excel | `createListObject` | ✅ | destructive-guard.test.mjs, excel-list-object.test.mjs |
| `wps_excel_create_pivot_table` | Excel | `createPivotTable` | ✅ | excel-advanced.test.mjs |
| `wps_excel_create_sheet` | Excel | `createSheet` |  | destructive-guard.test.mjs, excel-missing-halves-2.test.mjs, excel-page-setup.test.mjs, +1 |
| `wps_excel_create_workbook` | Excel | `createWorkbook` |  | close-safety.test.mjs, deprecated.test.mjs, excel-advanced.test.mjs, +6 |
| `wps_excel_delete_cell_comment` | Excel | `deleteCellComment` |  | excel-coverage.test.mjs |
| `wps_excel_delete_chart` | Excel | `deleteChart` |  | destructive-guard.test.mjs, excel-advanced.test.mjs |
| `wps_excel_delete_columns` | Excel | `deleteColumns` |  | excel-coverage.test.mjs |
| `wps_excel_delete_list_row` | Excel | `deleteListRow` |  | destructive-guard.test.mjs, excel-list-object.test.mjs |
| `wps_excel_delete_named_range` | Excel | `deleteNamedRange` |  | excel-missing-halves.test.mjs |
| `wps_excel_delete_rows` | Excel | `deleteRows` |  | destructive-guard.test.mjs |
| `wps_excel_delete_sheet` | Excel | `deleteSheet` |  | destructive-guard.test.mjs, sheet-ops.test.mjs |
| `wps_excel_diagnose_formula` | Excel | `diagnoseFormula` |  | excel-coverage.test.mjs |
| `wps_excel_evaluate_formula` | Excel | — |  | excel-coverage.test.mjs |
| `wps_excel_export_chart_as_image` | Excel | `exportChartAsImage` |  | excel-coverage.test.mjs |
| `wps_excel_export_range_as_image` | Excel | `exportRangeAsImage` |  | excel-coverage.test.mjs |
| `wps_excel_fill_series` | Excel | `fillSeries` |  | excel-contract-fixes.test.mjs, merged-tools.test.mjs |
| `wps_excel_find_in_sheet` | Excel | `findInSheet` | ✅ | excel-missing-halves.test.mjs |
| `wps_excel_find_replace` | Excel | `findReplaceExcel` | ✅ | find-replace.test.mjs |
| `wps_excel_freeze_panes` | Excel | `freezePanes` |  | excel-contract-fixes.test.mjs |
| `wps_excel_generate_formula` | Excel | `getContext` |  | excel-coverage.test.mjs |
| `wps_excel_get_cell_comments` | Excel | `getCellComments` |  | excel-coverage.test.mjs |
| `wps_excel_get_cell_info` | Excel | `getCellInfo` |  | excel-missing-halves-2.test.mjs |
| `wps_excel_get_cell_value` | Excel | `getCellValue` |  | excel-coverage.test.mjs |
| `wps_excel_get_conditional_formats` | Excel | `getConditionalFormats` |  | excel-missing-halves-2.test.mjs |
| `wps_excel_get_data_validations` | Excel | `getDataValidations` |  | excel-missing-halves-2.test.mjs |
| `wps_excel_get_formula` | Excel | `getFormula` |  | excel-list-object.test.mjs, excel-page-setup.test.mjs |
| `wps_excel_get_formula_audit` | Excel | `getFormulaAudit` |  | excel-page-setup.test.mjs |
| `wps_excel_get_list_objects` | Excel | `getListObjects` | ✅ | excel-list-object.test.mjs |
| `wps_excel_get_named_ranges` | Excel | `getNamedRanges` | ✅ | excel-missing-halves.test.mjs |
| `wps_excel_get_open_workbooks` | Excel | `getOpenWorkbooks` | ✅ | excel-advanced.test.mjs, excel-list-object.test.mjs, excel-missing-halves-2.test.mjs, +4 |
| `wps_excel_get_pivot_tables` | Excel | `getPivotTables` |  | excel-advanced.test.mjs |
| `wps_excel_get_selection` | Excel | `getSelection` |  | excel-coverage.test.mjs |
| `wps_excel_get_sheet_info` | Excel | `getExcelContext` | ✅ | excel-advanced.test.mjs, excel-list-object.test.mjs, excel-missing-halves-2.test.mjs, +2 |
| `wps_excel_get_sheet_list` | Excel | `getSheetList` | ✅ | excel-coverage.test.mjs |
| `wps_excel_get_sheet_settings` | Excel | `getSheetSettings` | ✅ | excel-page-setup.test.mjs |
| `wps_excel_goal_seek` | Excel | `goalSeek` | ✅ | excel-advanced.test.mjs |
| `wps_excel_group_columns` | Excel | `groupColumns` |  | excel-missing-halves-2.test.mjs |
| `wps_excel_group_rows` | Excel | `groupRows` |  | excel-page-setup.test.mjs |
| `wps_excel_hide_column` | Excel | `hideColumns` |  | excel-contract-fixes.test.mjs |
| `wps_excel_hide_rows` | Excel | `hideRows` |  | excel-contract-fixes.test.mjs, merged-tools.test.mjs |
| `wps_excel_insert_columns` | Excel | `insertColumns` |  | excel-coverage.test.mjs |
| `wps_excel_insert_excel_image` | Excel | `insertExcelImage` |  | excel-contract-fixes.test.mjs |
| `wps_excel_insert_rows` | Excel | `insertRows` |  | excel-coverage.test.mjs |
| `wps_excel_lock_cells` | Excel | `lockCells` |  | excel-coverage.test.mjs |
| `wps_excel_merge_cells` | Excel | `mergeCells` |  | excel-coverage.test.mjs |
| `wps_excel_move_sheet` | Excel | `moveSheet` |  | sheet-ops.test.mjs |
| `wps_excel_open_workbook` | Excel | `openWorkbook` | ✅ | file-ops.test.mjs, open-safety.test.mjs |
| `wps_excel_paste_range` | Excel | `pasteRange` |  | excel-coverage.test.mjs |
| `wps_excel_protect_sheet` | Excel | `protectSheet` |  | excel-contract-fixes.test.mjs |
| `wps_excel_protect_workbook` | Excel | `protectWorkbook` |  | excel-contract-fixes.test.mjs |
| `wps_excel_read_range` | Excel | `getRangeData` | ✅ | arg-shape-guard.test.mjs, excel-contract-fixes.test.mjs, excel-list-object.test.mjs, +3 |
| `wps_excel_refresh_all_data` | Excel | `refreshAllData` |  | excel-advanced.test.mjs |
| `wps_excel_refresh_links` | Excel | `refreshLinks` |  | excel-missing-halves-2.test.mjs |
| `wps_excel_refresh_pivot_tables` | Excel | `refreshPivotTables` |  | excel-advanced.test.mjs |
| `wps_excel_remove_conditional_format` | Excel | `removeConditionalFormat` |  | destructive-guard.test.mjs, excel-missing-halves-2.test.mjs |
| `wps_excel_remove_data_validation` | Excel | `removeDataValidation` |  | excel-missing-halves-2.test.mjs |
| `wps_excel_remove_duplicates` | Excel | `removeDuplicates` |  | excel-coverage.test.mjs |
| `wps_excel_rename_sheet` | Excel | `renameSheet` |  | sheet-ops.test.mjs |
| `wps_excel_reset_page_breaks` | Excel | `resetPageBreaks` |  | excel-page-setup.test.mjs |
| `wps_excel_resize_list_object` | Excel | `resizeListObject` |  | excel-list-object.test.mjs |
| `wps_excel_set_array_formula` | Excel | `setArrayFormula` |  | excel-coverage.test.mjs |
| `wps_excel_set_border` | Excel | `setBorder` |  | excel-contract-fixes.test.mjs |
| `wps_excel_set_cell_format` | Excel | `setCellFormat` | ✅ | cell-format.test.mjs, destructive-guard.test.mjs, excel-missing-halves-2.test.mjs |
| `wps_excel_set_cell_style` | Excel | `setCellStyle` |  | excel-contract-fixes.test.mjs |
| `wps_excel_set_cell_value` | Excel | `setCellValue` |  | arg-shape-guard.test.mjs, excel-coverage.test.mjs |
| `wps_excel_set_chart_labels` | Excel | `setChartLabels` | ✅ | excel-advanced.test.mjs |
| `wps_excel_set_column_width` | Excel | `setColumnWidth` |  | excel-coverage.test.mjs |
| `wps_excel_set_conditional_format` | Excel | `addConditionalFormat` |  | destructive-guard.test.mjs, excel-contract-fixes.test.mjs, excel-missing-halves-2.test.mjs |
| `wps_excel_set_data_validation` | Excel | `addDataValidation` |  | destructive-guard.test.mjs, excel-contract-fixes.test.mjs, excel-missing-halves-2.test.mjs |
| `wps_excel_set_formula` | Excel | `setFormula` | ✅ | excel-advanced.test.mjs, excel-page-setup.test.mjs, word-lifecycle.test.mjs |
| `wps_excel_set_hyperlink` | Excel | `setHyperlink` |  | excel-contract-fixes.test.mjs |
| `wps_excel_set_list_object_totals` | Excel | `setListObjectTotals` |  | excel-list-object.test.mjs |
| `wps_excel_set_named_range` | Excel | `createNamedRange` |  | destructive-guard.test.mjs, excel-missing-halves.test.mjs |
| `wps_excel_set_number_format` | Excel | `setNumberFormat` | ✅ | sheet-ops.test.mjs |
| `wps_excel_set_outline_levels` | Excel | `setOutlineLevels` |  | excel-page-setup.test.mjs |
| `wps_excel_set_print_area` | Excel | — |  | excel-coverage.test.mjs |
| `wps_excel_set_row_height` | Excel | `setRowHeight` |  | excel-coverage.test.mjs |
| `wps_excel_set_sheet_appearance` | Excel | `setSheetAppearance` |  | excel-page-setup.test.mjs |
| `wps_excel_set_sheet_header_footer` | Excel | `setSheetHeaderFooter` |  | excel-page-setup.test.mjs |
| `wps_excel_set_sheet_page_setup` | Excel | `setSheetPageSetup` | ✅ | excel-page-setup.test.mjs |
| `wps_excel_set_sheet_print_titles` | Excel | `setSheetPrintTitles` |  | excel-page-setup.test.mjs |
| `wps_excel_set_wrap_text` | Excel | `wrapText` |  | excel-missing-halves.test.mjs |
| `wps_excel_set_zoom` | Excel | `setZoom` |  | deprecated.test.mjs |
| `wps_excel_show_columns` | Excel | `showColumns` |  | excel-contract-fixes.test.mjs |
| `wps_excel_show_rows` | Excel | `showRows` |  | excel-contract-fixes.test.mjs |
| `wps_excel_sort_range` | Excel | `sortRange` |  | excel-contract-fixes.test.mjs |
| `wps_excel_subtotal` | Excel | `subtotal` |  | excel-missing-halves-2.test.mjs |
| `wps_excel_switch_sheet` | Excel | `switchSheet` |  | destructive-guard.test.mjs, excel-missing-halves-2.test.mjs, sheet-ops.test.mjs |
| `wps_excel_switch_workbook` | Excel | `switchWorkbook` |  | excel-coverage.test.mjs |
| `wps_excel_text_to_columns` | Excel | — |  | excel-coverage.test.mjs |
| `wps_excel_transpose` | Excel | `transpose` |  | excel-contract-fixes.test.mjs |
| `wps_excel_unlist_list_object` | Excel | `unlistListObject` |  | excel-list-object.test.mjs |
| `wps_excel_unmerge_cells` | Excel | `unmergeCells` |  | excel-coverage.test.mjs |
| `wps_excel_unprotect_sheet` | Excel | `unprotectSheet` |  | excel-coverage.test.mjs |
| `wps_excel_update_chart` | Excel | `updateChart` |  | excel-coverage.test.mjs |
| `wps_excel_update_list_object` | Excel | `updateListObject` |  | excel-list-object.test.mjs |
| `wps_excel_update_pivot_table` | Excel | `updatePivotTable` |  | excel-coverage.test.mjs |
| `wps_excel_write_range` | Excel | `setRangeData` | ✅ | arg-shape-guard.test.mjs, cell-format.test.mjs, destructive-guard.test.mjs, +11 |
| `wps_execute_method` | execute | — |  | cell-format.test.mjs, close-safety.test.mjs, destructive-guard.test.mjs, +14 |
| `wps_help` | 其他 | — | ✅ | deprecated.test.mjs, merged-tools.test.mjs, word-lifecycle.test.mjs, +1 |
| `wps_ppt_add_animation` | PPT | `setAnimation` |  | destructive-guard.test.mjs, ppt-contract-fixes.test.mjs, ppt-coverage.test.mjs, +1 |
| `wps_ppt_add_master_element` | PPT | `addMasterElement` |  | ppt-coverage.test.mjs |
| `wps_ppt_add_ppt_hyperlink` | PPT | `addPptHyperlink` |  | ppt-contract-fixes.test.mjs |
| `wps_ppt_add_shape` | PPT | `addShape` |  | destructive-guard.test.mjs, ppt-contract-fixes.test.mjs, ppt-coverage.test.mjs, +1 |
| `wps_ppt_add_slide` | PPT | `addSlide` | ✅ | destructive-guard.test.mjs, merged-tools.test.mjs, ppt-contract-fixes.test.mjs, +2 |
| `wps_ppt_add_textbox` | PPT | `addTextBox` | ✅ | destructive-guard.test.mjs, merged-tools.test.mjs, ppt-contract-fixes.test.mjs, +1 |
| `wps_ppt_align_shapes` | PPT | `alignShapes` |  | merged-tools.test.mjs |
| `wps_ppt_apply_transition_to_all` | PPT | `applyTransitionToAll` |  | ppt-contract-fixes.test.mjs |
| `wps_ppt_beautify` | PPT | `beautifySlide` |  | ppt-coverage.test.mjs |
| `wps_ppt_close_presentation` | PPT | `closePresentation` |  | close-safety.test.mjs, ppt-slimming.test.mjs |
| `wps_ppt_copy_slide` | PPT | `duplicateSlide` |  | ppt-contract-fixes.test.mjs |
| `wps_ppt_create_presentation` | PPT | `createPresentation` |  | close-safety.test.mjs, destructive-guard.test.mjs, merged-tools.test.mjs, +3 |
| `wps_ppt_delete_ppt_image` | PPT | `deletePptImage` |  | ppt-coverage.test.mjs |
| `wps_ppt_delete_shape` | PPT | `deleteShape` |  | ppt-coverage.test.mjs |
| `wps_ppt_delete_slide` | PPT | `deleteSlide` |  | ppt-coverage.test.mjs |
| `wps_ppt_delete_textbox` | PPT | `deleteTextBox` |  | ppt-coverage.test.mjs |
| `wps_ppt_distribute_shapes` | PPT | `distributeShapes` |  | ppt-coverage.test.mjs |
| `wps_ppt_duplicate_shape` | PPT | `duplicateShape` |  | ppt-coverage.test.mjs |
| `wps_ppt_export_slide_as_image` | PPT | `exportSlideAsImage` | ✅ | ppt-coverage.test.mjs |
| `wps_ppt_find_ppt_text` | PPT | `findPptText` |  | ppt-coverage.test.mjs |
| `wps_ppt_get_animations` | PPT | `getAnimations` |  | ppt-coverage.test.mjs |
| `wps_ppt_get_open_presentations` | PPT | `getOpenPresentations` |  | ppt-slimming.test.mjs, e2e |
| `wps_ppt_get_shapes` | PPT | `getShapes` | ✅ | ppt-coverage.test.mjs |
| `wps_ppt_get_slide_count` | PPT | `getSlideCount` | ✅ | ppt-coverage.test.mjs |
| `wps_ppt_get_slide_info` | PPT | `getSlideInfo` | ✅ | ppt-coverage.test.mjs |
| `wps_ppt_get_slide_master` | PPT | `getSlideMaster` |  | ppt-coverage.test.mjs |
| `wps_ppt_get_slide_notes` | PPT | `getSlideNotes` |  | ppt-coverage.test.mjs |
| `wps_ppt_get_slide_title` | PPT | `getSlideTitle` |  | ppt-coverage.test.mjs |
| `wps_ppt_get_table_cell` | PPT | `getPptTableCell` |  | ppt-coverage.test.mjs |
| `wps_ppt_get_textboxes` | PPT | `getTextBoxes` |  | ppt-coverage.test.mjs |
| `wps_ppt_group_shapes` | PPT | `groupShapes` |  | ppt-coverage.test.mjs |
| `wps_ppt_insert_ppt_chart` | PPT | `insertPptChart` |  | ppt-coverage.test.mjs |
| `wps_ppt_insert_ppt_image` | PPT | `insertPptImage` | ✅ | destructive-guard.test.mjs, merged-tools.test.mjs, ppt-coverage.test.mjs |
| `wps_ppt_insert_slides_from_file` | PPT | `insertSlidesFromFile` |  | ppt-coverage.test.mjs |
| `wps_ppt_insert_table` | PPT | `insertPptTable` | ✅ | ppt-coverage.test.mjs, ppt-slimming.test.mjs |
| `wps_ppt_move_slide` | PPT | `moveSlide` |  | ppt-coverage.test.mjs |
| `wps_ppt_open_presentation` | PPT | `openPresentation` | ✅ | ppt-coverage.test.mjs |
| `wps_ppt_remove_animation` | PPT | `removeAnimation` |  | ppt-contract-fixes.test.mjs |
| `wps_ppt_remove_ppt_hyperlink` | PPT | `removePptHyperlink` |  | ppt-coverage.test.mjs |
| `wps_ppt_remove_slide_transition` | PPT | `removeSlideTransition` |  | ppt-coverage.test.mjs |
| `wps_ppt_replace_ppt_image` | PPT | `replacePptImage` |  | ppt-coverage.test.mjs |
| `wps_ppt_replace_ppt_text` | PPT | `replacePptText` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_active_target` | PPT | `getOpenPresentations` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_animation_order` | PPT | `setAnimationOrder` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_background_color` | PPT | `setBackgroundColor` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_background_gradient` | PPT | `setBackgroundGradient` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_background_image` | PPT | `setBackgroundImage` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_font_color` | PPT | `setFontColor` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_image_style` | PPT | `setImageStyle` |  | ppt-contract-fixes.test.mjs |
| `wps_ppt_set_master_background` | PPT | `setMasterBackground` |  | ppt-contract-fixes.test.mjs |
| `wps_ppt_set_ppt_chart_data` | PPT | `setPptChartData` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_ppt_chart_style` | PPT | `setPptChartStyle` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_shape_effect` | PPT | `setShapeEffect` |  | ppt-contract-fixes.test.mjs, ppt-slimming.test.mjs |
| `wps_ppt_set_shape_fill` | PPT | `setShapeFill` | ✅ | ppt-contract-fixes.test.mjs |
| `wps_ppt_set_shape_position` | PPT | `setShapePosition` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_shape_style` | PPT | `setShapeStyle` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_shape_text` | PPT | `setShapeText` | ✅ | ppt-contract-fixes.test.mjs |
| `wps_ppt_set_shape_z_order` | PPT | `setShapeZOrder` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_background` | PPT | `setSlideBackground` |  | arg-shape-guard.test.mjs, merged-tools.test.mjs, ppt-contract-fixes.test.mjs |
| `wps_ppt_set_slide_content` | PPT | `setSlideContent` | ✅ | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_footer` | PPT | `setSlideFooter` |  | ppt-contract-fixes.test.mjs, ppt-slimming.test.mjs |
| `wps_ppt_set_slide_layout` | PPT | `setSlideLayout` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_notes` | PPT | `setSlideNotes` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_size` | PPT | `setSlideSize` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_subtitle` | PPT | `setSlideSubtitle` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_theme` | PPT | `setSlideTheme` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_title` | PPT | `setSlideTitle` | ✅ | ppt-coverage.test.mjs |
| `wps_ppt_set_slide_transition` | PPT | `setSlideTransition` |  | ppt-contract-fixes.test.mjs |
| `wps_ppt_set_table_cell` | PPT | `setPptTableCell` |  | ppt-coverage.test.mjs, ppt-slimming.test.mjs |
| `wps_ppt_set_table_format` | PPT | `setPptTableFormat` | ✅ | ppt-slimming.test.mjs |
| `wps_ppt_set_textbox_style` | PPT | `setTextBoxStyle` |  | ppt-coverage.test.mjs |
| `wps_ppt_set_textbox_text` | PPT | `setTextBoxText` |  | ppt-coverage.test.mjs |
| `wps_ppt_start_slide_show` | PPT | `startSlideShow` |  | ppt-contract-fixes.test.mjs |
| `wps_ppt_switch_presentation` | PPT | `switchPresentation` |  | ppt-coverage.test.mjs |
| `wps_ppt_switch_slide` | PPT | `switchSlide` |  | ppt-coverage.test.mjs |
| `wps_ppt_unify_font` | PPT | `unifyFont` |  | ppt-contract-fixes.test.mjs |
| `wps_status` | 其他 | — | ✅ | deprecated.test.mjs, e2e |
| `wps_word_accept_revisions` | Word | `acceptRevisions` |  | word-produce.test.mjs |
| `wps_word_add_content_control` | Word | `addContentControl` |  | word-longtail.test.mjs |
| `wps_word_add_endnote` | Word | `addEndnote` |  | word-longtail.test.mjs |
| `wps_word_add_footnote` | Word | `addFootnote` |  | word-longtail.test.mjs |
| `wps_word_add_table_lines` | Word | `addTableLines` |  | word-deep.test.mjs |
| `wps_word_apply_style` | Word | `applyStyle` | ✅ | word-common-coverage.test.mjs |
| `wps_word_close_document` | Word | `closeDocument` |  | word-deep.test.mjs, word-lifecycle.test.mjs, word-longtail.test.mjs, +1 |
| `wps_word_convert_table_to_text` | Word | `convertTableToText` |  | word-deep.test.mjs |
| `wps_word_create_document` | Word | `createDocument` | ✅ | destructive-guard.test.mjs, orphan-reclaim.test.mjs, word-common-coverage.test.mjs, +5 |
| `wps_word_delete_comment` | Word | `deleteComment` |  | word-produce.test.mjs |
| `wps_word_delete_table_line` | Word | `deleteTableLine` |  | destructive-guard.test.mjs, word-deep.test.mjs |
| `wps_word_enable_track_changes` | Word | `enableTrackChanges` |  | word-produce.test.mjs |
| `wps_word_find_in_document` | Word | `findInDocument` |  | word-common-coverage.test.mjs |
| `wps_word_find_replace` | Word | `findReplace` | ✅ | find-replace.test.mjs |
| `wps_word_generate_toc` | Word | `generateTOC` | ✅ | word-common-coverage.test.mjs |
| `wps_word_get_active_document` | Word | `getActiveDocument` | ✅ | word-lifecycle.test.mjs |
| `wps_word_get_bookmarks` | Word | `getBookmarks` |  | word-deep.test.mjs |
| `wps_word_get_comments` | Word | `getComments` | ✅ | word-deep.test.mjs, word-produce.test.mjs |
| `wps_word_get_content_controls` | Word | `getContentControls` |  | word-longtail.test.mjs |
| `wps_word_get_document_stats` | Word | `getDocumentStats` | ✅ | word-deep.test.mjs |
| `wps_word_get_document_text` | Word | `getDocumentText` | ✅ | excel-contract-fixes.test.mjs, find-replace.test.mjs |
| `wps_word_get_notes` | Word | `getNotes` | ✅ | word-longtail.test.mjs |
| `wps_word_get_open_documents` | Word | `getOpenDocuments` |  | word-deep.test.mjs, word-lifecycle.test.mjs, word-longtail.test.mjs, +2 |
| `wps_word_get_paragraphs` | Word | `getDocumentParagraphs` | ✅ | word-common-coverage.test.mjs |
| `wps_word_get_revisions` | Word | `getRevisions` | ✅ | word-produce.test.mjs |
| `wps_word_get_table_data` | Word | `getTableData` | ✅ | word-deep.test.mjs |
| `wps_word_get_tables` | Word | `getDocumentTables` | ✅ | word-deep.test.mjs |
| `wps_word_get_track_changes_status` | Word | `getTrackChangesStatus` |  | word-common-coverage.test.mjs |
| `wps_word_insert_bookmark` | Word | `insertBookmark` |  | word-common-coverage.test.mjs, word-deep.test.mjs, word-longtail.test.mjs |
| `wps_word_insert_comment` | Word | `addComment` |  | word-deep.test.mjs, word-produce.test.mjs |
| `wps_word_insert_cross_reference` | Word | `insertCrossReference` |  | word-longtail.test.mjs |
| `wps_word_insert_footer` | Word | `insertFooter` |  | file-ops.test.mjs |
| `wps_word_insert_header` | Word | `insertHeader` |  | file-ops.test.mjs |
| `wps_word_insert_hyperlink` | Word | `insertHyperlink` |  | word-deep.test.mjs |
| `wps_word_insert_image` | Word | `insertImage` |  | word-common-coverage.test.mjs |
| `wps_word_insert_index` | Word | `insertIndex` |  | word-longtail.test.mjs |
| `wps_word_insert_page_break` | Word | `insertPageBreak` |  | word-common-coverage.test.mjs |
| `wps_word_insert_page_numbers` | Word | `insertPageNumbers` | ✅ | word-produce.test.mjs |
| `wps_word_insert_section_break` | Word | `insertSectionBreak` |  | word-common-coverage.test.mjs |
| `wps_word_insert_table` | Word | `insertTable` |  | destructive-guard.test.mjs, word-deep.test.mjs |
| `wps_word_insert_text` | Word | `insertText` | ✅ | excel-contract-fixes.test.mjs, file-ops.test.mjs, find-replace.test.mjs, +5 |
| `wps_word_mail_merge` | Word | `mailMerge` | ✅ | word-longtail.test.mjs |
| `wps_word_merge_table_cells` | Word | `mergeTableCells` |  | word-deep.test.mjs |
| `wps_word_open_document` | Word | `openDocument` | ✅ | open-safety.test.mjs |
| `wps_word_proofread_basic` | Word | — |  | word-common-coverage.test.mjs |
| `wps_word_reject_revisions` | Word | `rejectRevisions` |  | word-produce.test.mjs |
| `wps_word_replace_bookmark_content` | Word | `replaceBookmarkContent` |  | word-common-coverage.test.mjs |
| `wps_word_replace_range` | Word | `replaceRange` |  | word-common-coverage.test.mjs |
| `wps_word_set_columns` | Word | `setColumns` |  | word-produce.test.mjs |
| `wps_word_set_font` | Word | `setFont` | ✅ | merged-tools.test.mjs, word-common-coverage.test.mjs |
| `wps_word_set_line_spacing` | Word | `setLineSpacing` |  | word-common-coverage.test.mjs |
| `wps_word_set_page_setup` | Word | `setPageSetup` |  | word-common-coverage.test.mjs |
| `wps_word_set_paragraph` | Word | — | ✅ | word-common-coverage.test.mjs |
| `wps_word_set_table_cell` | Word | `setTableCell` | ✅ | word-deep.test.mjs |
| `wps_word_set_table_format` | Word | `setTableFormat` |  | word-deep.test.mjs |
| `wps_word_set_text_color` | Word | `setTextColor` |  | word-common-coverage.test.mjs |
| `wps_word_smart_fill_field` | Word | `smartFillField` | ✅ | word-common-coverage.test.mjs |
| `wps_word_split_table_cell` | Word | `splitTableCell` |  | word-deep.test.mjs |
| `wps_word_switch_document` | Word | `switchDocument` |  | word-common-coverage.test.mjs |

## 未被测试点名的工具（0）

无——267 个工具全部至少被一个测试或 e2e 点名。

## 解析不出桥 action 的工具（10）

- `wps_batch`
- `wps_call`
- `wps_excel_evaluate_formula`
- `wps_excel_set_print_area`
- `wps_excel_text_to_columns`
- `wps_execute_method`
- `wps_help`
- `wps_status`
- `wps_word_proofread_basic`
- `wps_word_set_paragraph`

## 已知 WPS 不支持（实测，不再尝试）

以下不是「没做」，是 COM 层面不成立（证据见 `docs/FIXES.md` 与探针脚本），推广材料应写「不支持」：

- Word 水印：页眉 `Shapes` 拒绝一切添加，`Count` 恒为 0。
- 文档属性：`BuiltInDocumentProperties` / `CustomDocumentProperties` 是坏壳。
- Excel 切片器：`SlicerCaches.Add2` 可调用，但 `Slicers.Count` 恒为 0。
- 方案管理器：`Worksheet.Scenarios` 在 COM 里是方法，语义读不干净。

## 说明

- 「被测试点名」只统计工具名在测试 / e2e 源码里出现过，是覆盖的**必要条件**，不是充分条件；真正验证行为的是各场景测试本身。
- 覆盖率的 ratchet 断言在 `test/spec-reproduction.test.mjs`；未覆盖清单也可用 `node scripts/smoke-tools.mjs` 查看。

---
name: wps-excel
description: 用 WPS 表格工具读写区域、设置格式与公式、建图表与透视表、导出与转换。
whenToUse: 任务针对表格、工作簿、单元格区域、公式、图表、透视表或数据清洗时。
---

# WPS 表格

## 直接广告的工具

wps_excel_get_sheet_list、wps_excel_read_range、wps_excel_write_range、wps_excel_set_cell_format、wps_excel_set_number_format、wps_excel_set_formula、wps_excel_create_chart、wps_excel_create_pivot_table、wps_excel_find_replace、wps_excel_get_open_workbooks、wps_excel_open_workbook，以及通用的 wps_common_save、wps_common_save_as、wps_convert_to_pdf、wps_convert_format。

其余 80 多个表格工具（排序、筛选、条件格式、批注、冻结窗格、命名范围、数据验证、转置、分列、分类汇总、导出图片等）通过 wps_call 使用，清单见同目录 reference.md。

## 参数约定

- 单元格定位用 {sheet, row, col}，行列都从 1 开始。
- 区域用 A1 记法字符串，如 A1:C10；**sheet 参数在几乎所有工具上都已生效**，不传时才用当前活动工作表。要操作非活动工作表就显式传 sheet。
- read_range 的 include_header 为 true 时把首行当表头返回。
- 批量写入用 wps_excel_write_range，数据是二维数组，从起始单元格向右下填充。
- 常用工具的参数名（写错会被参数校验拒绝，不再静默忽略）：

| 工具 | 参数 |
|---|---|
| sort_range | range、column（从 1 开始的列号）、ascending |
| auto_filter | range、column（筛选列）、criteria |
| copy_range | source、destination（不传 destination 才只复制到剪贴板） |
| transpose | source、destination |
| subtotal | range、groupBy、function、columns |
| fill_series | range、type、step、startValue、direction |
| auto_fill | sourceRange、targetRange |
| set_border | range、borderStyle、position、color |
| set_data_validation | range、type、formula（list 类型可直接给 "a,b,c"） |
| set_conditional_format | range、condition（如 ">100"、"between(1,10)"）、format（red_fill 等或 #RRGGBB） |
| set_cell_style | range、style（具名样式，名字不存在会报错） |
| insert_excel_image | filePath、cell、width、height |
| get_cell_comments | range（可选，缩小范围） |
| protect_sheet / protect_workbook | protect（false 为取消保护） |
| freeze_panes | row、column（"冻结到第几行/列"），freeze=false 取消冻结 |
| hide_rows / show_rows | startRow、endRow（或 row + count） |
| hide_column / show_column | column、count、hide |

## 工作表管理

- create_sheet / copy_sheet / move_sheet 的 **position 是 0 基**（0 = 最前）。create/copy 不传 position 时追加到末尾。
- delete_sheet **必须显式给 name**，不会退回到活动工作表；工作簿只剩一张表时会拒绝删除。
- rename_sheet 用 {oldName, newName}；switch_sheet 用 {name}。
- 这些工具的返回值里同时给出 0 基的 position、1 基的 index 与工作表总数。

## 常用流程

1. wps_excel_get_open_workbooks 或 wps_excel_get_sheet_list 确认目标工作簿与工作表。
2. 写入前先 read_range 看现有内容，避免覆盖用户数据。
3. 写入用 write_range 一次提交整块数据。
4. 公式用 set_formula，单元格地址必须是有效的 A1 记法。**给区域设公式会把同一个公式字符串写进每个单元格**（与 Excel 的 Formula 赋值一致，不会按行/列调整相对引用）——需要逐行递增的引用就逐格调用，或直接用 write_range 写入数值。
5. 图表与透视表的源区域必须包含表头行、数据连续无空行。
6. 写回后用 read_range 复验，再 wps_common_save。

## 已知坑

- 查找替换的搜索引擎对大小写与全半角敏感；find_replace 返回的是**真实替换计数**。
- 条件格式、排序、筛选、转置等此前会静默忽略关键参数，现在参数名写错会报错，但**值仍不合法时也可能报错**——例如给不存在的具名样式、未知的条件格式名。报错信息会说明可用取值。
- Hyperlinks 不会覆盖已有单元格的值：给非空单元格设超链接时，显示文本不会替换原内容。
- 插入图片的路径可以是相对路径（按当前工作目录解析），文件不存在会明确报 image file not found。
- create_chart 没有 has_header 参数（此前宣称但无法实现，已移除）；图表类型用 chart_type。
- 全列引用如 A:A 在公式里代价高，尽量写成 A1:A1000 这类有界区域。
- set_formula 写区域时**不做相对引用调整**：对 B2:B5 一次写入 `=SUMIF(明细!$A$2:$A$13,A2,明细!$C$2:$C$13)`，B2:B5 会全部引用 A2，看起来像公式算错。结果文本里会带这条提示；要逐行公式请逐格调用 set_formula。
- generate_formula 需要 description 参数，不传会校验失败。
- 导出图片依赖剪贴板，并发或无人值守时可能失败，失败信息要如实上报。

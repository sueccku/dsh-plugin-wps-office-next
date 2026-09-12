# Excel 参数契约测试结果

方法：scripts/contract-test.mjs 为每个工具按 schema 合成参数，经 MCP 调用，
并开启 WPS_OFFICE_TRACE（mcp/src/client/wps-client.ts 的参数追踪）记录工具**实际发出**的参数键，
再与 wps-com.ps1 里各 action **实际读取**的 `$p.*` 键做差集。
「工具发了、桥从不读」的参数就是被静默忽略的参数。

## 可信结论：53 个工具存在被忽略的参数

这一类结论不依赖文档状态，来自两次独立运行的一致结果。

### A. 正在广告的工具里就有（对用户影响最大）

| 工具 | 被忽略的参数 | 后果 |
|---|---|---|
| wps_excel_find_replace | find, replace | 桥读 findText/replaceText，所以查找替换静默无效；**且 findReplace 的实现是 Word 专用（Get-WpsWord）**，Excel 调用实际作用在 Word 文档上 |
| wps_excel_set_cell_format | format, sheet | 桥只读 numberFormat，整个 format 对象（粗体/字号/背景色等）被丢弃 |
| wps_excel_set_number_format | sheet | 只在活动工作表生效 |
| wps_excel_create_chart | sheet, hasHeader | 忽略工作表与表头设置 |
| wps_excel_open_workbook | filePath | 桥读 path，按文档参数调用会失败 |

### B. 整组工作表操作被忽略

delete_sheet(name)、rename_sheet(oldName)、switch_sheet(name)、copy_sheet(name/newName/position)、
move_sheet(name/position)、create_sheet(position) —— 桥读的是 sheet/before/after。

### C. 其他参数名不匹配

sort_range(column/ascending)、set_conditional_format(condition/format)、set_data_validation(type/formula)、
copy_range 与 transpose(source/destination)、set_hyperlink(url/text)、insert_excel_image(filePath/imagePath/cell)、
hide_rows 与 show_rows(startRow/endRow)、show_columns(startColumn/endColumn)、hide_column(count/hide)、
freeze_panes(freeze)、protect_sheet 与 protect_workbook(protect)、close_workbook(save)、add_comment(cell)、
set_border(borderStyle)、set_cell_style(style)、auto_filter(column)、subtotal(columns)、fill_series(direction)。

### D. 纯粹的 sheet 参数被忽略（约 20 个工具）

merge/unmerge_cells、set_column_width、set_row_height、clear_range、paste_range、text_to_columns、
insert/delete_rows、insert/delete_columns、group_rows、get_formula、get_cell_info、lock_cells、
set_array_formula、delete_cell_comment、get_cell_comments、remove_duplicates 等。
这些工具承诺可指定工作表，实际总在活动工作表上操作。

## 不可信的结论：失败清单

首次运行泄漏了 85 个工作簿（每轮 create+close 的清理大半失败，最后整个 COM 层被模态对话框阻塞），
因此失败清单被工作簿状态污染，**不能作为缺陷依据**。需要的是一次干净状态下的重跑。

已确认非缺陷（测试参数问题，非产品问题）：update_pivot_table / update_chart 需要额外的目标参数；
open_workbook 用了 png 路径。

## 下一步

1. WPS 重启后在干净状态下重跑，得到可信的失败清单。
2. 按用户优先级修复：先 Excel 高频工具（find_replace、set_cell_format、set_number_format、工作表操作组），
   再 Word，最后 PPT。
3. 每修一项补一条回归用例，纳入 scripts/contract-test.mjs 的可断言子集。

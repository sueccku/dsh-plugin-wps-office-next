# 工具总清单（v0.9.0 快照：270 个；当前 310）

> **2026-10-21 P1 注**：本文是按 v0.9.0 快照整理的旧清单；P1 累计新增 40 个隐藏工具
> （Word 14、PPT 12、Excel 14），当前注册 **310 / 广告 95**。
> 最新清单以 `docs/tool-coverage.md`、`skills/*/reference.md`（自动生成）与 `spec/tool-definitions.json` 为准。
>
> 数据来源：`spec/tool-definitions.json`（工具目录）与 `spec/advertised.json`（广告面），
> 按 v0.9.0 检出快照整理。工具增删后请以 `docs/tool-coverage.md` 与 `node scripts/snapshot-tools.mjs` 为准。

本项目的 MCP 工具共 **270 个**：其中 **95 个**随每次请求直接广告给模型（★），
其余 **175 个**不占每次请求的上下文，模型通过 `wps_help` 查目录、再用 `wps_call` 或 `wps_batch` 触达。

| 应用 | 工具数 | 广告面 |
| --- | ---: | ---: |
| 表格（WPS 表格 / Excel） | 118 | 46 |
| 文字（WPS 文字 / Word） | 61 | 27 |
| 演示（WPS 演示 / PPT） | 77 | 14 |
| 通用（保存 / 应用信息 / 选区） | 7 | 2 |
| 格式转换 | 2 | 2 |
| 逃生舱 | 1 | 0 |
| 门面与元工具 | 4 | 4 |
| **合计** | **270** | **95** |

> 270 = 应用模块里注册的 265 个（Excel 118 + Word 61 + PPT 77 + 通用 7 + 转换 2）
> ＋ 5 个门面/逃生舱工具（`wps_batch`、`wps_call`、`wps_help`、`wps_status`、`wps_execute_method`）。

## 表格（WPS 表格 / Excel）（118）

| 工具 | 简介 |
| --- | --- |
| `wps_excel_add_comment` | 给单元格添加批注。 |
| ★ `wps_excel_add_list_row` | 给表末尾追加一行，可同时写入这一行的值（按列顺序）。 |
| `wps_excel_add_sparkline` | 在单元格区域里加迷你图（单元格内的微型图表）：range 是数据，location 是放图的位置，两者形状要一致（如 B2:B5 → C2:C5）。 |
| `wps_excel_auto_filter` | 对Excel指定范围应用自动筛选。 |
| ★ `wps_excel_auto_fit` | 按内容自动调整列宽与行高。 |
| `wps_excel_auto_fit_columns` | 按内容自动调整列宽（不动行高）。 |
| `wps_excel_auto_fit_rows` | 按内容自动调整行高（不动列宽）。 |
| `wps_excel_auto_sum` | 对指定范围的列或行自动求和，并将结果写入目标单元格。 |
| ★ `wps_excel_calculate` | 强制重算公式。 |
| `wps_excel_clean_data` | 数据清洗工具，支持多种清洗操作的组合。 |
| ★ `wps_excel_clear_formats` | 清除区域的格式（字体、颜色、边框、数字格式），单元格内容保留。 |
| ★ `wps_excel_clear_pivot_table` | 清除透视表在表上的报表区域（数据源不动）。 |
| ★ `wps_excel_clear_range` | 清除指定范围的内容、格式或全部。 |
| `wps_excel_clear_sparkline` | 清除指定区域上的迷你图（数据不动）。 |
| ★ `wps_excel_close_workbook` | 关闭指定的Excel工作簿，可选是否保存。 |
| `wps_excel_consolidate` | 把多块来源区域按指定函数汇总写入目标区域（Excel 的合并计算）。 |
| ★ `wps_excel_copy_format` | 把一块区域的格式复制到另一块区域（只复制格式，不改数值与公式）。 |
| ★ `wps_excel_copy_range` | 复制Excel指定范围到目标位置。 |
| `wps_excel_copy_sheet` | 复制当前工作簿中的指定工作表。 |
| ★ `wps_excel_create_chart` | 在Excel中创建图表。 |
| ★ `wps_excel_create_list_object` | 把一块区域变成「表」（ListObject）：自动带表头与筛选按钮，并可用结构化引用（表名[列名]）写公式，之后能按表加行/删行/加总计行。 |
| ★ `wps_excel_create_pivot_table` | 创建Excel透视表，用于数据汇总和分析。 |
| ★ `wps_excel_create_sheet` | 在当前工作簿中创建新的工作表。 |
| `wps_excel_create_workbook` | 新建一个空白Excel工作簿。 |
| `wps_excel_delete_cell_comment` | 删除Excel单元格上的批注。 |
| `wps_excel_delete_chart` | 删除工作表上的图表（不删它引用的数据）。 |
| `wps_excel_delete_columns` | 删除Excel中指定位置的一列或多列。 |
| `wps_excel_delete_list_row` | 删除表里的第几行（表体行，从 1 开始，不含表头）。 |
| `wps_excel_delete_named_range` | 删除指定的命名范围（只删名字，不动单元格内容）。 |
| ★ `wps_excel_delete_rows` | 删除Excel中指定位置的一行或多行。 |
| `wps_excel_delete_sheet` | 删除当前工作簿中的指定工作表。 |
| `wps_excel_diagnose_formula` | 诊断公式错误，分析原因并提供修复建议。 |
| ★ `wps_excel_evaluate_formula` | 计算并返回公式结果（由 Excel 求值，不写入任何单元格）。 |
| ★ `wps_excel_export_chart_as_image` | 将工作表中指定的图表导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。 |
| ★ `wps_excel_export_range_as_image` | 将工作表中指定区域导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。 |
| `wps_excel_fill_series` | 自动填充序列数据。 |
| ★ `wps_excel_find_in_sheet` | 在工作表里查找文本并返回每一个命中的单元格地址（不改动任何内容）。 |
| ★ `wps_excel_find_replace` | 在Excel中查找并替换内容。 |
| ★ `wps_excel_freeze_panes` | 冻结/取消冻结窗格。 |
| `wps_excel_generate_formula` | 根据自然语言描述生成Excel公式。 |
| `wps_excel_get_cell_comments` | 获取Excel指定范围内的所有批注。 |
| ★ `wps_excel_get_cell_info` | 获取单元格的详细信息（值、公式、格式等）。 |
| `wps_excel_get_cell_value` | 获取Excel指定单元格的值。 |
| ★ `wps_excel_get_conditional_formats` | 列出区域上生效的条件格式规则（序号 + 类型），用于先看清楚再改。 |
| `wps_excel_get_data_validations` | 读取区域上的数据验证规则（类型、来源公式、提示语）。 |
| `wps_excel_get_formula` | 获取Excel指定单元格的公式。 |
| `wps_excel_get_formula_audit` | 审计一个单元格的公式依赖：它引用了谁（precedents）、谁引用了它（dependents）、直接引用几处，并可选在界面上画出追踪箭头。 |
| ★ `wps_excel_get_list_objects` | 列出工作簿（或指定工作表）里的全部「表」及其结构：名字、范围、行列数、列名、表格样式、总计行开关，以及每列可直接写进公式的结构化引用。 |
| ★ `wps_excel_get_named_ranges` | 列出工作簿里的全部命名范围及其引用位置。 |
| ★ `wps_excel_get_open_workbooks` | 获取当前所有已打开的Excel工作簿列表。 |
| `wps_excel_get_pivot_tables` | 列出工作簿（或指定工作表）上的透视表：名字、所在区域，以及（能读到时）行字段与数据字段。 |
| `wps_excel_get_selection` | 获取当前Excel中选中区域的信息，包括范围地址、行列数等。 |
| ★ `wps_excel_get_sheet_info` | 获取工作表的结构信息：工作簿与工作表名、已用范围地址、表头、当前单元格。 |
| ★ `wps_excel_get_sheet_list` | 获取当前工作簿的所有工作表列表，包含名称、索引和是否为活动工作表。 |
| ★ `wps_excel_get_sheet_settings` | 读一张工作表的页面设置、打印设置、页眉页脚与外观：方向、纸张、页边距（磅）、缩放或按页适配、是否居中、打印区域与打印标题、页眉页脚、可见性、标签色、手动分页符数量。 |
| ★ `wps_excel_goal_seek` | 单变量求解：反复调整 changingCell，直到 cell 的公式结果等于 goal。 |
| `wps_excel_group_columns` | 把一段列折叠分组（分级显示）。 |
| `wps_excel_group_rows` | 对Excel中指定范围的行进行分组，便于折叠/展开管理。 |
| `wps_excel_hide_column` | 隐藏或显示指定列。 |
| `wps_excel_hide_rows` | 隐藏Excel中指定范围的行。 |
| `wps_excel_insert_columns` | 在Excel中指定位置插入一列或多列。 |
| `wps_excel_insert_excel_image` | 在Excel中插入图片到指定位置。 |
| `wps_excel_insert_rows` | 在Excel中指定位置插入一行或多行。 |
| `wps_excel_lock_cells` | 锁定或解锁Excel指定范围的单元格（需配合工作表保护使用）。 |
| `wps_excel_merge_cells` | 合并Excel指定范围的单元格。 |
| `wps_excel_move_sheet` | 移动指定工作表到新的位置。 |
| ★ `wps_excel_open_workbook` | 打开指定路径的Excel工作簿文件。 |
| `wps_excel_paste_range` | 粘贴已复制的内容到指定位置。 |
| `wps_excel_protect_sheet` | 保护或取消保护工作表。 |
| `wps_excel_protect_workbook` | 保护或取消保护工作簿，防止结构被修改（如添加/删除工作表）。 |
| ★ `wps_excel_read_range` | 读取Excel指定范围的单元格数据，返回二维数组格式的数据。 |
| `wps_excel_refresh_all_data` | 刷新整个工作簿的外部数据连接与透视表（相当于 Excel 的「全部刷新」）。 |
| `wps_excel_refresh_links` | 刷新工作簿引用的全部外部链接并报告条数。 |
| `wps_excel_refresh_pivot_tables` | 刷新透视表：给 pivotTable 只刷新那一张，不给就刷新目标工作表上的全部。 |
| `wps_excel_remove_conditional_format` | 删除区域上的条件格式规则。 |
| `wps_excel_remove_data_validation` | 删除区域上的数据验证规则（下拉框、输入限制）。 |
| `wps_excel_remove_duplicates` | 删除指定范围内的重复行。 |
| `wps_excel_rename_sheet` | 重命名当前工作簿中的指定工作表。 |
| `wps_excel_reset_page_breaks` | 清除工作表上的手动分页符，恢复按内容自动分页。 |
| `wps_excel_resize_list_object` | 调整表覆盖的区域（长短变化），表名、样式与结构化引用都保留。 |
| `wps_excel_set_array_formula` | 为Excel指定范围设置数组公式（CSE数组公式）。 |
| `wps_excel_set_border` | 设置Excel单元格边框样式，支持不同粗细、位置和颜色。 |
| ★ `wps_excel_set_cell_format` | 设置Excel单元格格式，包括字体、颜色、背景色、粗体、斜体、字号等。 |
| `wps_excel_set_cell_style` | 应用预定义样式到Excel单元格，如标题、强调、输入、输出等内置样式。 |
| `wps_excel_set_cell_value` | 设置Excel指定单元格的值。 |
| ★ `wps_excel_set_chart_labels` | 给图表加标题与坐标轴标题（分类轴 = 横轴，数值轴 = 纵轴）。 |
| ★ `wps_excel_set_column_width` | 设置Excel指定列的列宽。 |
| ★ `wps_excel_set_conditional_format` | 设置条件格式。 |
| `wps_excel_set_data_validation` | 设置Excel单元格的数据验证规则，如下拉列表、数值范围、日期范围等。 |
| ★ `wps_excel_set_formula` | 在指定单元格设置Excel公式。 |
| `wps_excel_set_hyperlink` | 为Excel单元格设置超链接。 |
| `wps_excel_set_list_object_totals` | 开/关总计行，并可指定某一列的汇总方式。 |
| `wps_excel_set_named_range` | 设置命名范围。 |
| ★ `wps_excel_set_number_format` | 设置Excel单元格的数字格式。 |
| `wps_excel_set_outline_levels` | 控制分级显示的展开层级与汇总位置：rowLevels/columnLevels 指定行/列显示到第几级（1 表示全部折叠），summaryRow/summaryColumn 指定汇总行在上还是下、汇总列在左还是右。 |
| `wps_excel_set_print_area` | 设置打印区域（range 省略时清除）。 |
| `wps_excel_set_row_height` | 设置Excel指定行的行高。 |
| ★ `wps_excel_set_sheet_appearance` | 设置工作表的可见性与标签色：可见 / 隐藏 / 深度隐藏（veryHidden，用户界面上无法取消隐藏），以及标签颜色（十六进制如 #FF9900）。 |
| ★ `wps_excel_set_sheet_header_footer` | 设置打印页眉页脚。 |
| ★ `wps_excel_set_sheet_page_setup` | 设置工作表的页面：方向、纸张、页边距、缩放、是否居中、是否打印网格线与行列标题。 |
| ★ `wps_excel_set_sheet_print_titles` | 设置打印时每页重复的行/列（打印标题）：如行 $1:$1 让表头每页都出现，列 $A:$A 让第一列每页都出现。 |
| `wps_excel_set_wrap_text` | 设置或取消单元格的自动换行。 |
| `wps_excel_set_zoom` | 设置当前工作表的缩放比例（10-400%）。 |
| `wps_excel_show_columns` | 显示Excel中已隐藏的列。 |
| `wps_excel_show_rows` | 显示Excel中已隐藏的行。 |
| ★ `wps_excel_sort_range` | 对Excel选定区域按指定列排序。 |
| `wps_excel_subtotal` | 创建分类汇总。 |
| ★ `wps_excel_switch_sheet` | 切换到指定的工作表，使其成为活动工作表。 |
| ★ `wps_excel_switch_workbook` | 切换到指定名称的Excel工作簿。 |
| `wps_excel_text_to_columns` | 将文本按分隔符拆分到多列。 |
| `wps_excel_transpose` | 转置数据（行列互换）。 |
| `wps_excel_unlist_list_object` | 把表转回普通区域（数据与格式保留）：结构化引用、筛选按钮与表对象都会消失。 |
| ★ `wps_excel_unmerge_cells` | 拆分Excel中已合并的单元格，恢复为独立的单元格。 |
| `wps_excel_unprotect_sheet` | 取消保护当前工作表。 |
| `wps_excel_update_chart` | 更新Excel图表的属性，包括标题、颜色、图例、数据标签等。 |
| `wps_excel_update_list_object` | 改表本身的设置：改名、换表格样式、显示或隐藏表头行、显示或隐藏筛选按钮。 |
| `wps_excel_update_pivot_table` | 更新已有透视表的配置，包括添加/移除字段、修改聚合方式等。 |
| ★ `wps_excel_write_range` | 向Excel指定范围写入数据。 |

## 文字（WPS 文字 / Word）（61）

| 工具 | 简介 |
| --- | --- |
| `wps_word_accept_revisions` | 接受修订：给 index 只接受那一处，不填则接受全部。 |
| `wps_word_add_content_control` | 在光标处插入一个内容控件（可填写的结构化区域），可指定类型、标题、标签与初始文本。 |
| `wps_word_add_endnote` | 在光标处插入尾注（文档末尾的注释）。 |
| `wps_word_add_footnote` | 在光标处插入脚注（页面底部的注释）。 |
| `wps_word_add_table_lines` | 给表格增加行或列。 |
| ★ `wps_word_apply_style` | 应用Word样式到当前选中区域或指定范围。 |
| ★ `wps_word_close_document` | 关闭 Word 文档，可选是否保存。 |
| `wps_word_convert_table_to_text` | 把表格转成普通文本，列之间用一个分隔符（tab/comma/paragraph，或直接给一个字符）。 |
| ★ `wps_word_create_document` | 新建一个空白 Word 文档（不是打开已有文件）。 |
| `wps_word_delete_comment` | 删除批注：给 index 只删那一条（序号见 wps_word_get_comments），不填则全部删除。 |
| `wps_word_delete_table_line` | 删除表格的第几行或第几列（从 1 开始）。 |
| `wps_word_enable_track_changes` | 开启或关闭Word文档的修订模式（Track Changes）。 |
| `wps_word_find_in_document` | 在Word文档中查找文本并返回位置信息，不执行替换操作。 |
| ★ `wps_word_find_replace` | 在Word文档中查找并替换文本。 |
| ★ `wps_word_generate_toc` | 根据文档中的标题样式自动生成目录。 |
| ★ `wps_word_get_active_document` | 获取当前WPS Writer活动文档的基本信息 |
| `wps_word_get_bookmarks` | 列出文档里的全部书签（名字与位置）。 |
| ★ `wps_word_get_comments` | 列出文档里的全部批注：序号、正文、作者、时间。 |
| `wps_word_get_content_controls` | 列出文档里的内容控件（可填写的结构化区域）：序号、类型、标题、标签、当前文本。 |
| ★ `wps_word_get_document_stats` | 文档统计：页数、字数、字符数、段落数、行数。 |
| ★ `wps_word_get_document_text` | 获取当前Word文档的文本内容。 |
| `wps_word_get_header_footer` | 读取某一节的页眉与页脚：文本内容与其中的域（页码、总页数等）。 |
| ★ `wps_word_get_notes` | 列出文档里的脚注与尾注及其内容。 |
| ★ `wps_word_get_open_documents` | 获取当前WPS Writer中所有已打开的文档列表。 |
| ★ `wps_word_get_page_setup` | 读取某一节的页面设置：纸张方向、纸张大小、页边距（磅）与页面宽高。 |
| ★ `wps_word_get_paragraphs` | 获取Word文档的段落结构信息，返回每段的文本、样式和字符位置。 |
| ★ `wps_word_get_revisions` | 列出文档里的修订（插入/删除/替换等），并报告「修订跟踪」当前是否打开。 |
| ★ `wps_word_get_table_data` | 按行列读出某张表的全部单元格文本（合并单元格的非起点格子会是空字符串）。 |
| ★ `wps_word_get_tables` | 列出文档里的全部表格：序号、行列数、样式、文本预览。 |
| `wps_word_get_track_changes_status` | 获取当前文档的修订模式状态。 |
| `wps_word_insert_bookmark` | 在当前光标位置或选中区域插入书签。 |
| `wps_word_insert_comment` | 在Word文档选中内容处插入批注 |
| `wps_word_insert_cross_reference` | 在光标处插入交叉引用（引用标题/书签/脚注等）。 |
| `wps_word_insert_footer` | 设置页脚内容。 |
| `wps_word_insert_header` | 设置页眉内容。 |
| `wps_word_insert_hyperlink` | 在光标处插入超链接。 |
| `wps_word_insert_image` | 在Word文档中插入图片。 |
| `wps_word_insert_index` | 在文档末尾插入索引（按索引项自动生成）。 |
| `wps_word_insert_page_break` | 在文档光标位置插入分页符 |
| ★ `wps_word_insert_page_numbers` | 给某一节的页眉或页脚插入页码。 |
| `wps_word_insert_section_break` | 插入分节符，用于将文档分为不同的节，以便对各节应用不同的页面设置。 |
| `wps_word_insert_table` | 在Word文档光标位置插入表格 |
| ★ `wps_word_insert_text` | 在Word文档中插入文本。 |
| ★ `wps_word_mail_merge` | 邮件合并：以当前文档为母版，接一个数据文件（CSV），把表格里的值填到母版的占位符处；默认合并成一份新文档，给了 splitDir 就每个数据行单独存一份（母版不动）。 |
| `wps_word_merge_table_cells` | 把一块矩形区域合并成一个单元格：从（startRow, startColumn）到（endRow, endColumn）。 |
| ★ `wps_word_open_document` | 打开指定路径的Word文档。 |
| `wps_word_proofread_basic` | 对中文文本进行基础校对，检测常见问题。 |
| `wps_word_reject_revisions` | 拒绝修订：给 index 只拒绝那一处，不填则拒绝全部（回到改之前的原文）。 |
| `wps_word_replace_bookmark_content` | 替换Word文档中书签的内容。 |
| ★ `wps_word_replace_range` | 按字符范围精确替换Word文档中的文本。 |
| `wps_word_set_columns` | 设置分栏：栏数、栏间距、是否加分隔线。 |
| ★ `wps_word_set_font` | 设置字体格式，包括字体名称、字号、加粗、斜体、颜色等。 |
| `wps_word_set_line_spacing` | 设置段落行距。 |
| ★ `wps_word_set_page_setup` | 设置某一节的页面布局：页面方向、纸张大小、页边距。 |
| ★ `wps_word_set_paragraph` | 设置段落格式：对齐方式、行间距、段前/段后间距、缩进。 |
| ★ `wps_word_set_table_cell` | 写入表格的某个单元格（行列都从 1 开始）。 |
| `wps_word_set_table_format` | 设置表格外观：表格样式名、是否显示边框、按内容或按窗口自动调整宽度、给表头行加底纹色、表头行跨页重复。 |
| `wps_word_set_text_color` | 设置Word文档中选中文字的颜色 |
| ★ `wps_word_smart_fill_field` | 智能填写Word模板中的字段。 |
| `wps_word_split_table_cell` | 把一个单元格拆成 rows x columns 个小格（默认 1 x 2）。 |
| ★ `wps_word_switch_document` | 切换到指定名称的文档。 |

## 演示（WPS 演示 / PPT）（77）

| 工具 | 简介 |
| --- | --- |
| `wps_ppt_add_animation` | 给形状加动画，三种用法：给 preset 就按预设给整页（或指定形状）加依次出场的动画；给 effect 且 effectKind=emphasis 加强调动画；只给 effect 则加入场/退场动画。 |
| `wps_ppt_add_master_element` | 向母版中添加新元素。 |
| `wps_ppt_add_ppt_hyperlink` | 为幻灯片中的形状添加超链接。 |
| `wps_ppt_add_shape` | 在幻灯片中添加形状。 |
| ★ `wps_ppt_add_slide` | 添加新幻灯片到演示文稿。 |
| ★ `wps_ppt_add_textbox` | 在幻灯片中添加文本框。 |
| `wps_ppt_align_shapes` | 对齐幻灯片中的多个形状。 |
| `wps_ppt_apply_transition_to_all` | 为所有幻灯片应用统一的切换效果。 |
| `wps_ppt_beautify` | 一键美化幻灯片，优化排版、配色、字体和间距。 |
| `wps_ppt_close_presentation` | 关闭演示文稿。 |
| `wps_ppt_copy_slide` | 复制幻灯片到指定位置。 |
| `wps_ppt_create_presentation` | 新建空白演示文稿。 |
| `wps_ppt_delete_ppt_image` | 删除幻灯片中指定的图片。 |
| `wps_ppt_delete_shape` | 删除幻灯片中指定的形状。 |
| `wps_ppt_delete_slide` | 删除指定的幻灯片。 |
| `wps_ppt_delete_textbox` | 删除幻灯片上指定的文本框。 |
| `wps_ppt_distribute_shapes` | 等距分布幻灯片中的多个形状。 |
| `wps_ppt_duplicate_shape` | 复制幻灯片中的指定形状。 |
| `wps_ppt_end_slide_show` | 结束正在进行的幻灯片放映。 |
| ★ `wps_ppt_export_slide_as_image` | 将指定幻灯片导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。 |
| `wps_ppt_find_ppt_text` | 在演示文稿中搜索指定文本。 |
| `wps_ppt_get_animations` | 获取幻灯片上所有动画效果的列表。 |
| `wps_ppt_get_open_presentations` | 获取当前所有已打开的演示文稿列表。 |
| ★ `wps_ppt_get_shapes` | 获取幻灯片中所有形状的列表信息。 |
| ★ `wps_ppt_get_slide_count` | 获取演示文稿中的幻灯片总数。 |
| ★ `wps_ppt_get_slide_info` | 获取指定幻灯片的详细信息，包括布局、元素列表等。 |
| `wps_ppt_get_slide_master` | 获取当前演示文稿的母版信息。 |
| `wps_ppt_get_slide_notes` | 获取指定幻灯片的备注内容。 |
| `wps_ppt_get_slide_title` | 获取指定幻灯片的标题。 |
| `wps_ppt_get_table_cell` | 获取PPT表格中指定单元格的文本内容。 |
| `wps_ppt_get_textboxes` | 获取幻灯片上所有文本框的列表。 |
| `wps_ppt_group_shapes` | 将幻灯片中的多个形状组合为一个组。 |
| `wps_ppt_insert_ppt_chart` | 在幻灯片中插入数据图表。 |
| ★ `wps_ppt_insert_ppt_image` | 插入图片到幻灯片中。 |
| `wps_ppt_insert_slides_from_file` | 从另一个 PPT 文件把整页幻灯片插入到【当前活动演示文稿】，并保留来源幻灯片的原始格式（字体/配色/版式/图片）。 |
| ★ `wps_ppt_insert_table` | 在幻灯片中插入表格。 |
| `wps_ppt_move_slide` | 移动幻灯片到指定位置。 |
| ★ `wps_ppt_open_presentation` | 打开指定路径的演示文稿文件。 |
| `wps_ppt_remove_animation` | 移除幻灯片中指定的动画效果。 |
| `wps_ppt_remove_ppt_hyperlink` | 移除幻灯片中形状的超链接。 |
| `wps_ppt_remove_slide_transition` | 移除幻灯片的切换效果。 |
| `wps_ppt_replace_ppt_image` | 原位替换幻灯片中的某张图片：保留原图的位置、尺寸、旋转角度，删除旧图后在同一矩形内插入新图。 |
| `wps_ppt_replace_ppt_text` | 在演示文稿中查找并替换文本。 |
| `wps_ppt_set_active_target` | 锁定后续所有 PPT 操作的【目标演示文稿】（按文件名）。 |
| `wps_ppt_set_animation_order` | 调整动画在时间线上的播放顺序。 |
| `wps_ppt_set_background_color` | 设置幻灯片背景为指定颜色。 |
| `wps_ppt_set_background_gradient` | 为幻灯片设置渐变色背景。 |
| `wps_ppt_set_background_image` | 设置幻灯片背景为指定图片。 |
| `wps_ppt_set_font_color` | 设置幻灯片中指定形状的文字颜色。 |
| `wps_ppt_set_image_style` | 设置幻灯片中指定图片的样式。 |
| `wps_ppt_set_master_background` | 设置母版背景样式。 |
| `wps_ppt_set_ppt_chart_data` | 更新幻灯片中已有图表的数据。 |
| `wps_ppt_set_ppt_chart_style` | 设置幻灯片中图表的样式属性。 |
| `wps_ppt_set_shape_effect` | 给形状设置视觉效果：阴影、边框、渐变填充、填充透明度；只改给出来的项，其余保持原样。 |
| ★ `wps_ppt_set_shape_fill` | 设置幻灯片中指定形状的填充颜色。 |
| `wps_ppt_set_shape_position` | 设置幻灯片中指定形状的位置和大小。 |
| `wps_ppt_set_shape_style` | 设置幻灯片中形状的样式，包括填充颜色、边框颜色、边框粗细和圆角半径。 |
| ★ `wps_ppt_set_shape_text` | 设置幻灯片中指定形状的文字内容。 |
| `wps_ppt_set_shape_z_order` | 设置形状在幻灯片中的层级顺序（Z轴排列）。 |
| `wps_ppt_set_slide_background` | 设置幻灯片背景，支持多种背景类型。 |
| ★ `wps_ppt_set_slide_content` | 设置幻灯片的正文内容区域。 |
| `wps_ppt_set_slide_footer` | 一次设置演示文稿的页脚三件套：幻灯片编号、页脚文字、日期时间，以及日期的显示格式。 |
| `wps_ppt_set_slide_layout` | 设置幻灯片的版式布局。 |
| `wps_ppt_set_slide_notes` | 设置幻灯片的备注内容，用于演讲提示。 |
| `wps_ppt_set_slide_size` | 设置演示文稿的幻灯片尺寸。 |
| `wps_ppt_set_slide_subtitle` | 设置幻灯片的副标题。 |
| `wps_ppt_set_slide_theme` | 把演示文稿套用为指定的主题模板（.thmx / .potx / .pptx）。 |
| ★ `wps_ppt_set_slide_title` | 设置幻灯片的标题文本。 |
| `wps_ppt_set_slide_transition` | 设置幻灯片的页面切换效果。 |
| `wps_ppt_set_table_cell` | 设置PPT表格中指定单元格的文本内容。 |
| ★ `wps_ppt_set_table_format` | 设置表格的外观与样式。 |
| `wps_ppt_set_textbox_style` | 设置文本框样式（字体大小、颜色、粗体等）。 |
| `wps_ppt_set_textbox_text` | 设置指定文本框的文本内容。 |
| `wps_ppt_start_slide_show` | 开始幻灯片放映。 |
| `wps_ppt_switch_presentation` | 切换到指定名称的演示文稿。 |
| `wps_ppt_switch_slide` | 切换到指定的幻灯片页面。 |
| `wps_ppt_unify_font` | 统一演示文稿中所有幻灯片的字体。 |

## 通用（保存 / 应用信息 / 选区）（7）

| 工具 | 简介 |
| --- | --- |
| `wps_common_get_app_info` | 获取当前 WPS 应用的基本信息。 |
| `wps_common_get_selected_text` | 获取当前文档中选中的文本内容。 |
| `wps_common_ping` | 检测WPS应用连接状态。 |
| ★ `wps_common_save` | 保存当前文档。 |
| ★ `wps_common_save_as` | 将当前文档另存为指定路径和格式。 |
| `wps_common_set_selected_text` | 替换当前文档中选中的文本内容。 |
| `wps_common_wire_check` | 检查与WPS加载项之间的通信线路状态。 |

## 格式转换（2）

| 工具 | 简介 |
| --- | --- |
| ★ `wps_convert_format` | 将当前文档转换为其他格式。 |
| ★ `wps_convert_to_pdf` | 将当前文档转换为PDF格式。 |

## 逃生舱（1）

| 工具 | 简介 |
| --- | --- |
| `wps_execute_method` | 最后手段：直接调用原始 WPS COM 方法。 |

## 门面与元工具（4）

| 工具 | 简介 |
| --- | --- |
| ★ `wps_batch` | 按顺序批量执行多个 WPS 工具调用，用于跨应用的连续操作；单次最多 50 项。 |
| ★ `wps_call` | 执行任意已注册但未直接广告的 WPS 工具；先用 wps_help 取到工具名与参数 |
| ★ `wps_help` | 查询未直接广告的工具：无参看分组概览，传 app 或 query 查目录，传 tool 取完整参数 schema |
| ★ `wps_status` | 查看 WPS 连接状态、当前活动应用与所选工具面；编辑前先调用它 |

## 说明

- ★ = 当前档位（standard）下直接广告的工具；未标 ★ 的仍已注册，可经 `wps_help` + `wps_call` 调用。
- `minimal` 档只广告 4 个工具，`full` 档广告全部 270 个；档位由 `wps_status` 汇报。
- `wps_batch` 单次最多 50 次调用；`wps_execute_method` 是不校验参数的原始 COM 逃生舱，文档中不推荐优先使用。
- 名称前缀与分类对应：`wps_excel_*` / `wps_word_*` / `wps_ppt_*` / `wps_common_*` / `wps_convert_*`。

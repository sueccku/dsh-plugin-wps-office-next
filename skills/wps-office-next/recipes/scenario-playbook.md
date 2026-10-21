# 80 场景任务配方（P0，2026-10-19；P2 补充批量）

> 用法：先按"用户要的产出"找组，再按"起手体检 → 主链路 → 验收 → 回退"执行。
> 每条都假设：已 `wps_status` 确认连接；多文件时先 `get_open_*` 防误改；改原件前先 `save_as` 副本。
> 标记 `[边界]` 的场景不要硬做，按 `references/scope-and-boundaries.md` 说明 + 给替代路径。
> **文件级批量**（XA9/XA10/XA11）的失败清单/目录编排见同目录 `batch-files.md`，不要把通用文件系统工具塞进插件。

## Excel（E1–E8、XL1–XL20）

| 场景 | 起手体检 | 主链路 | 验收/回退 |
| --- | --- | --- | --- |
| E1 多簿汇总 | `get_open_workbooks`/`get_sheet_list`/`get_sheet_info` | `open_workbook readOnly:true` → `read_range` → `save_as` 总表 → `write_range` → `remove_duplicates` → 汇总公式/透视 | 回读行数/总额；源文件只读；"自动刷新"无工具时如实说明 |
| E2 条件格式整行 | `get_conditional_formats`/`get_sheet_info` | `set_conditional_format`（cellValue/colorScale/dataBar）；公式规则当前不支持 | 回读规则条数；公式规则/整行做不到时说明 |
| E3 VLOOKUP | `get_sheet_list`/`read_range` | 单格 `set_formula` + `copy_range`（多格 `set_formula` 是广播同一公式） | 抽查公式文本/结果；未匹配清单 |
| E4 拆列/去重/空行 | `read_range`/`get_sheet_info` | `text_to_columns`（分隔符）或读入后 `write_range`；`remove_duplicates`；`delete_rows` | 行数/重复键回读；"拆到行"无工具时用公式/文本处理 |
| E5 透视 | `get_sheet_list` | `create_pivot_table` → `update_pivot_table` → `refresh_pivot_tables` | 行/列/值字段回读；切片器 WPS 不支持 |
| E6 打印就绪 | `get_sheet_settings` | `set_sheet_page_setup` → `set_sheet_print_titles` → `set_sheet_header_footer` → `convert_to_pdf` | 回读页面设置；起始页码/打印页数无工具时说明 |
| E7 图表诊断 | `read_range`/`get_cell_info` | 数据源修正；`update_chart` 改标题/类型/图例；日期轴/空值当前无参数 | 导图目检；不能改轴时如实说明 |
| E8 保护/协作 `[边界]` | `get_sheet_settings` | `protect_sheet`/`lock_cells`（允许编辑区域无工具） | 保护状态回读；云协作说明边界 |
| XL1 SUMIFS 全 0 | `get_cell_info` 查文本型数字/合并格 | 先转数值/取消合并/统一表头 → `set_formula` 单格+填充 | 抽查关键合计；计算模式/卡顿说明 |
| XL2 二维取值 | `get_sheet_list`/`read_range` | INDEX/MATCH 或 `set_formula`；确认 WPS 函数版本 | 未匹配清单；XLOOKUP 不可用时给替代 |
| XL3 全表错误值 | `read_range`（分页） | 先诊断错误类型，再决定是否 IFERROR；不要整列覆盖硬编码 | 错误地址清单；覆盖前备份 |
| XL4 地址拆分 | `read_range` | 文本函数/分列；直辖市规则单独处理 | 抽样核对省/市/区 |
| XL5 图标集 | `get_conditional_formats` | 当前无 icon set；用色阶/数据条或手工 | 明确说明缺口，不假装设置成功 |
| XL6 固定宽度/智能填充 | `read_range` | `text_to_columns` 只有分隔符；固定宽度/Ctrl+E 无工具 → 读入后模型重写 | 回读新列；说明替代 |
| XL7 CSV 导入 | `open_workbook`（只读） | 编码/列格式不可控；必要时在 WPS 里导入或先转 xlsx | 前导零/日期抽样；做不到时说明 |
| XL8 去重留最近 | `read_range` | 多键排序后 `remove_duplicates`，或用公式标记 | 去重键/保留规则回读 |
| XL9 透视计算字段/切片器 | `create_pivot_table` | `update_pivot_table` 能改字段；计算字段/日期分组/切片器受限 | 回读字段；切片器说明 WPS 限制 |
| XL10 组合图/次轴/趋势线 | `create_chart`/`update_chart` | 当前只有基础图表类型；组合/次轴/趋势线无参数 | 导图目检；说明缺口 |
| XL11 冻结/分级/隐藏 | `get_sheet_settings` | `freeze_panes`/`group_rows`/`hide_column`/`show_columns` | 回读设置；无回读项说明 |
| XL12 允许编辑区域 | `get_sheet_settings` | `lock_cells`+`protect_sheet`；AllowEditRanges/FormulaHidden 无工具 | 保护状态；缺口说明 |
| XL13 忘记密码 `[边界]` | — | 不可解密 | 明确拒绝，建议从备份恢复 |
| XL14 大表卡顿 | `get_sheet_settings`/`read_range` | 分页读；公式/对象/计算模式只能抽样；无诊断工具 | 抽样结论 + 建议另存/拆分 |
| XL15 外链管理 | `refresh_links` | 只能刷新；列出/改源/断开无工具 | 说明缺口；建议在 WPS 里编辑链接 |
| XL16 多簿静态合并 | `get_open_workbooks` | 逐个只读读取 → 按列名对齐 → `write_range` 总表 | 行数/合计；表头不在第 1 行时先定位 |
| XL17 多键排序/高级筛选 | `sort_range`/`auto_filter` | 单列排序/单条件筛选可做；多键/高级筛选需多遍或公式 | 回读顺序；不稳定时如实说明 |
| XL18 批注/超链接/名称清理 | `get_cell_comments`/`get_named_ranges` | 批注可读删；超链接只能写不能读；名称可列出/删除 | 清理数量回读；超链接缺口说明 |
| XL19 目标倒推/敏感性 | `goal_seek` | 单变量求解可做；模拟运算表/方案管理器 WPS 限制 | 回读结果；说明边界 |
| XL20 打印区域/分页符 | `get_sheet_settings` | `set_print_area`/`reset_page_breaks`/页脚；插入分页符无工具 | 回读区域/页脚；说明缺口 |

## Word（W1–W6、WD1–WD15）

| 场景 | 起手体检 | 主链路 | 验收/回退 |
| --- | --- | --- | --- |
| W1 长文排版 | `get_document_stats`/`get_paragraphs`（分页） | `set_page_setup` → `set_font`（range 精确）→ `apply_style` → `insert_page_numbers` | 样式/边距回读；题注/文献/多级编号缺工具时说明 |
| W2 目录页码 | `get_page_setup`/`get_header_footer` | `generate_toc`（insert/update）→ `insert_page_numbers` | 目录条目/节页码回读；起始值限制说明 |
| W3 邮件合并 | `get_document_text` 看占位符 | 名单→CSV → `mail_merge`（splitDir/fileNameField） | 产物数量/字段抽查；一对多/图片域说明 |
| W4 修订批注 | `get_revisions`/`get_comments` | `accept_revisions`/`reject_revisions`/`delete_comment`/`enable_track_changes` | 回读修订/批注/修订模式 |
| W5 表格跨页 | `get_tables`/`get_table_data` | `set_table_format`（headerRepeat 单行）/`set_paragraph` | 回读表头重复；多行/允许跨页断行说明 |
| W6 格式统一 | `get_paragraphs`/`find_in_document` | `find_replace`/`apply_style`/`set_font`；无"选择格式相似" | 抽样回读；逐段写 |
| WD1 水印 `[边界]` | — | WPS COM 不支持 | 说明限制，建议 WPS 界面 |
| WD2 分节横纵 | `get_page_setup` | `insert_section_break` → `set_page_setup(section)` | 逐节回读方向；打印/PDF 需人工确认 |
| WD3 奇偶页眉 | `get_header_footer` | 当前只有整段 text + section；firstPage/oddEven 无工具 | 说明缺口；逐节写 |
| WD4 页眉横线 | `get_header_footer` | `insert_header text:""` 清空；边框/横线无工具 | 回读文本；说明限制 |
| WD5 题注/编号 | `get_paragraphs` | 无 SEQ/题注/更新域工具 | 说明缺口；建议手工/域 |
| WD6 交叉引用 | `find_in_document`/`get_bookmarks` | `insert_cross_reference` 可插；更新域/目标选取无工具 | 回读引用文本；说明限制 |
| WD7 图表目录 | `generate_toc` | 无 TOC \c/图表目录工具 | 说明缺口；可用 TOC 域兜底 |
| WD8 文档比较 | `get_revisions` | 无 CompareDocuments 工具 | 说明缺口；建议 WPS 审阅→比较 |
| WD9 标签/信封 | `get_document_text` | `mail_merge` 只做一人一份；一页多记录/套打无参数 | 说明缺口；建议 WPS 邮件合并 |
| WD10 图片域证书 | `get_document_text` | 纯文本合并可做；图片合并域无工具 | 说明缺口；逐份插图 |
| WD11 内容控件/表单 | `get_content_controls` | `add_content_control`（type/title/tag/text）；下拉/日期/保护无参数 | 回读控件；说明参数边界 |
| WD12 按作者修订 | `get_revisions`/`get_comments` | 只有整体/index 接受；按作者筛选无工具 | 逐条 index 处理；说明缺口 |
| WD13 表格公式/排序 | `get_tables`/`get_table_data` | `set_table_cell` 写文本；公式/排序/斜线无工具 | 说明缺口；公式可写文本值 |
| WD14 文档属性 `[边界]` | — | WPS COM 坏壳 | 说明限制，建议 WPS 界面/365 |
| WD15 限制编辑 | `get_document_stats` | 无 protect/restrict 工具 | 说明缺口；交付 PDF/只读文件 |

## PPT（P1–P4、PT1–PT10）

| 场景 | 起手体检 | 主链路 | 验收/回退 |
| --- | --- | --- | --- |
| P1 套模板 | `get_open_presentations`/`set_active_target`/`get_slide_count` | `set_slide_theme` → `unify_font` → `set_slide_footer`；逐页 `get_shapes` 体检 | 逐页抽查；导图目检 |
| P2 对齐分布 | `get_shapes` | `align_shapes`/`distribute_shapes`/`group_shapes`/`set_shape_position` | 回读几何；相对页面居中要手算 |
| P3 动画切换 | `get_slide_count`/`get_shapes` | `add_animation`（逐页/逐形状）/`apply_transition_to_all`；平滑映射缺 | 回读动画；播放不可自动验证 |
| P4 只改字色 | `get_shapes`/`get_textboxes` | `set_font_color` 逐形状；无对象类型/页码过滤/字体回读 | 导图目检；说明验证缺口 |
| PT1 母版版式 | `get_slide_master` | 现有 `add_master_element` 只支持 textbox/shape；新版式/占位符锁定无工具 | 说明缺口；导图目检 |
| PT2 主题色/字体 | `get_slide_master` | `set_slide_theme`（文件模板）可套；主题色/字体编辑无工具 | 说明缺口；用模板文件替代 |
| PT3 4:3→16:9 | `get_slide_count`/`get_slide_info` | `set_slide_size`；无 scaleContent/受影响页报告 | 先 save_as 副本；逐页导图检查 |
| PT4 只换字体 | `get_shapes` | `unify_font` 可统一字体；无 sourceFont/preserveSize/回读 | 导图目检；说明可能改字号 |
| PT5 字体嵌入 | `get_slide_count` | 无嵌入/体积回读工具 | 说明缺口；建议 WPS 字体嵌入 |
| PT6 图片裁剪/压缩/替换 | `get_shapes` | `set_image_style` 支持 crop*；`replace_ppt_image`；压缩/体积回读缺 | 回读尺寸；导图目检 |
| PT7 SmartArt `[边界]` | — | D3 已删 | 说明边界 |
| PT8 音视频 `[边界]` | — | D3 已删 | 说明边界 |
| PT9 节/隐藏页 | `get_slide_count` | 无节/隐藏页工具；导出选项只有整份 | 说明缺口；用多份文件替代 |
| PT10 备注/提词 | `get_slide_notes` | `set_slide_notes` 逐页写；`get_slide_notes` 读；导出备注 PDF 无选项 | 回读备注；导出受限说明 |

## 跨应用/文件（X1、F1、XA1–XA15）

| 场景 | 起手体检 | 主链路 | 验收/回退 |
| --- | --- | --- | --- |
| X1 Excel→PPT 活链接 `[边界]` | `read_range`/`get_slide_count` | 一次性读数据 → `insert_ppt_chart`/`set_ppt_chart_data` 或表格；无活链接 | 明确"不可刷新"；静态交付 |
| F1 批量打印/导 PDF `[边界]` | 文件清单（Harness） | 逐个 `open`→`convert_to_pdf`→`close save:false`；打印不做 | 产物清单；appType 必传；失败清单 |
| XA1 表格→Word | `read_range`/`get_tables` | `insert_table` 带 data（P0 起）→ 逐格修正 | 表格行列/关键值回读 |
| XA2 表格→PPT | `read_range` | `insert_table` + `set_table_cell`（PPT 无 data，逐格） | 行列/关键值回读；成本高时说明 |
| XA3 Word 大纲→PPT | `get_paragraphs` | 逐段读 → `add_slide` 带 title/content；无大纲导入 | 页序/标题回读 |
| XA4 PPT 备注→Word | `get_slide_notes` | 逐页读 → `insert_text`；可加门面（P1） | 页数/文本抽查 |
| XA5–XA8 PDF `[边界]` | — | 无 PDF 工具 | 明确边界 + 外部工具 |
| XA9 批量转换 | 文件清单（Harness） | 逐个 `open`→`convert_to_pdf`/`convert_format`→`close save:false` | 失败清单；会员/风控说明 |
| XA10 批量改名 | 文件清单（Harness） | 插件无文件系统工具，由 Harness 完成 | 改名清单；不要用 COM |
| XA11 批量套模板 | 文件清单（Harness） | 逐文件 `open readOnly:false`→改→`save_as`→`close save:false` | 抽样核对；无 O(N) 批处理门面 |
| XA12 OLE/链接 `[边界]` | — | 无 OLE 工具 | 说明边界；一次性数据 |
| XA13 手机端 `[边界]` | — | 桌面 COM 不覆盖 | 说明边界 |
| XA14 WPS AI `[边界]` | — | 不在工具面 | 说明边界 |
| XA15 大文件/长文卡顿 | `get_sheet_info`/`get_document_stats` | 分页读、抽样；无诊断工具 | 结论注明抽样 |

## 通用收尾（每个场景都做）

1. 用户要原件不动 → 先 `save_as` 副本，再改副本；改完回读目标。
2. 用户要保存 → 明确 `save:true`；否则默认不保存关闭。
3. 报告分三栏：已执行 / 已核对 / 未验证（视觉、PDF 页数、打印效果）。
4. 遇到不支持：先报边界与原因，再给替代路径；不要反复 `execute_method`。

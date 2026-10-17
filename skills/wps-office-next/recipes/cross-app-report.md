# 配方：跨应用报告（表格数据 → 文字报告 / 演示汇报 → 导出）

适用：用户要"把这个工作簿的数据做成一份 Word 报告 / PPT 汇报"，或"数据在表格里，结论写进文档"。
跨应用的关键不再是工具，而是**源与目的的对应关系**和**两次目标切换之间的竞争窗口**。

## 0. 先确定的事

1. 源：哪一本工作簿、哪一张表、哪一块**有界**区域（表头在第几行、数据到第几行）。
2. 目的：新建一份，还是写进已有文档 / 演示？输出路径是什么？
3. 落哪些字段：不是整表照搬，而是"哪几列进哪一个位置"。
4. 口径：汇总行怎么算、空值/重复怎么处理（用户给了就直接用）。

## 1. 读源（表格侧）

```
wps_excel_get_open_workbooks            # 有几本打开
wps_excel_switch_workbook {name}        # 点名源工作簿（多本时必做）
wps_excel_get_sheet_info {sheet}        # usedRange / 表头
wps_excel_read_range {range:"A1:F200", sheet}   # 一次读成二维数组
```

- 一次读**一整块**，不要逐格读。
- 读回来的第一行是表头时，先确认列序与列名，再按列名取值——不要按位置硬编码。
- 需要汇总数时在表格侧算好（或用 `wps_excel_evaluate_formula` 取公式结果），把**结果值**搬到目的侧；不要把公式连到另一个应用里去。

## 2. 写目的（文字侧示例）

```
wps_word_create_document                # 或 wps_word_open_document {path}
wps_word_get_active_document            # 核对：name / path 是不是我要写的那份
wps_word_insert_text {text, position:"end"}
wps_word_apply_style {styleName:"标题 1", range:{start,end}}
wps_word_insert_table                   # 表格数据用表格，不要用正文空格对齐
wps_word_set_table_cell {table, row, column, text}
```

**每次切换应用后重新读一次目标身份**（`wps_word_get_active_document` / `wps_excel_get_sheet_info` / `wps_ppt_get_slide_info`）。
两个应用交替写入时，焦点切换是真实风险：身份读回不一致就停下来重新解析目标，不要继续写。

## 3. 写目的（演示侧示例）

```
wps_ppt_open_presentation {path}        # 或 wps_ppt_create_presentation
wps_ppt_get_slide_info {slideIndex:1}   # 核对 presentationName / presentationPath + 拿画布尺寸
wps_ppt_add_slide {layout:"titleContent"}   # 加页时就定版式
wps_ppt_set_slide_title {slideIndex, title}
wps_ppt_set_slide_content {slideIndex, content}
```

坐标是**磅**、画布尺寸是**像素**（见 `wps-ppt` 技能 §2 与它的 `references/layout-and-shapes.md`）。放进表格的数字要短，长表格拆成多页或改成要点。

## 4. 核对（这一步不能省）

| 核对项 | 怎么做 |
|---|---|
| 源数据没被改 | 源侧重新 `read_range` 一个关键区域，与写前读到的一致 |
| 目的内容对得上 | 目的侧回读关键段落/单元格，与源数据**逐项对应**（不是"看起来像"） |
| 目标身份没漂 | 两个应用的 `get_*` 身份读回与开始时一致 |
| 数字口径 | 总额/计数与源侧独立算出来的值一致 |
| 落盘 | `wps_common_save_as {outputPath}`，再确认输出文件真的存在 |

## 5. 导出

```
wps_common_save
wps_convert_to_pdf {appType:"word", outputPath:"...\\报告.pdf"}
```

- **必须显式传 `appType`**：不传时按 Excel → Word → PPT 取第一个**正在运行**的文档，而跨应用任务里三个应用往往同时开着，会导错对象。
- 导出只证明"文件生成了"，**不证明排版正确**。能读图/PDF 渲染结果就核对；不能就明确说"视觉效果未验证"。

## 6. 验收标准

- [ ] 源工作簿的特征区域与任务开始时逐项一致（没有被顺手改掉）。
- [ ] 目的文件里每个关键数字都能追到源区域的具体单元格。
- [ ] 非目标文件（任务涉及的**另一个**已打开文档）没有被修改。
- [ ] 输出路径存在且可独立读取；未保存的内容没有被丢弃。
- [ ] 报告里区分了「已核对」与「未验证（视觉/渲染）」。

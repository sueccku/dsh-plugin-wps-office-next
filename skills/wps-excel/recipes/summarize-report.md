# 配方：汇总报表

适用：多份来源数据 → 一份新的汇总工作簿 / 工作表（含公式、图表、打印设置）。
目标：**原件不动、口径写清、每根数字可追溯、结果可核对**。

## 1. 先固定输入与口径（写进报告开头或备注）

| 项 | 必须定下来的内容 |
|---|---|
| 输入 | 每份源文件/工作表的名字与路径、表头在第几行、数据到第几行 |
| 口径 | 分组维度、聚合方式（SUM/COUNT/AVERAGE…）、空值怎么处理、重复行算不算、日期与数值类型是否干净 |
| 输出 | 输出文件路径（新文件？新工作表？）、表名、放置位置 |
| 公式策略 | 哪几列**必须是公式**（用户要可更新），哪几列可以是静态值 |
| 图表 | 数据源区域（含表头、连续无空行）、图表类型、放置位置 |
| 打印导出 | 纸张与方向、每页重复的表头行、打印区域、是否导出 PDF |

用户已经给了口径就直接用；只有**口径缺失且影响结果**时才问。

## 2. 保护原件

```powershell
# 多份来源：先只读方式打开，或读成结构化数据后立刻关掉
wps_excel_open_workbook {path:"...\\华北.xlsx", readOnly:true}
wps_excel_get_sheet_info {sheet:"明细"}
wps_excel_read_range {range:"A1:H500", sheet:"明细"}
wps_excel_close_workbook {name:"...", save:false}
```

- 汇总**不改来源**：只读打开/读完就关（`save:false`）。
- 输出到新文件：`wps_excel_create_workbook` → 写 → `wps_common_save_as {outputPath}`（**不是**保存到来源文件上）。
- 输出到来源文件里的新工作表：先 `wps_common_save_as` 建工作副本，确认活动目标是副本再写。

## 3. 建输出结构

```
wps_excel_create_sheet {name:"汇总"}            # 或 create_workbook
wps_excel_write_range {range:"A1", sheet:"汇总", data:[["部门","月份","金额","占比"]]}
wps_excel_write_range {range:"A2", sheet:"汇总", data:[...]}     # 明细块一次写完
```

- 表头一次写全，再写数据块；不要一列一列来回写。
- 需要逐行递增引用的列（例如占比 = 本行/总额）**逐格** `set_formula`，或先用 `fill_series` 造出引用骨架再改。
  详见 `references/ranges-and-formulas.md` §2。
- 需要"可更新"的汇总列保留公式；用户明确只要结果时才写静态值。

## 4. 建表 / 透视 / 图表

```
wps_excel_create_list_object {range:"A1:D100", name:"汇总表"}    # 需要结构化引用或自动扩展时
wps_excel_create_pivot_table {sourceRange:"A1:D100", destinationCell:"F2",
                              rowFields:["部门"], valueFields:[{field:"金额",aggregation:"SUM"}]}
wps_excel_create_chart {range:"A1:B6", chartType:"column_clustered", title:"各部门金额"}
wps_excel_set_chart_labels {title:"各部门金额", categoryAxisTitle:"部门", valueAxisTitle:"金额"}
```

- 源区域必须**含表头行、数据连续无空行**。
- 透视表的目标单元格要在数据之外，避免覆盖。

## 5. 打印与导出

```
wps_excel_set_sheet_page_setup {sheet:"汇总", orientation:"landscape", fitToPagesWide:1}
wps_excel_set_sheet_print_titles {sheet:"汇总", printTitleRows:"$1:$1"}
wps_excel_set_sheet_header_footer {sheet:"汇总", centerFooter:"第 &P 页 / 共 &N 页"}
wps_excel_set_print_area {sheet:"汇总", range:"A1:H60"}
wps_common_save
wps_convert_to_pdf {appType:"excel", outputPath:"...\\汇总.pdf"}
```

## 6. 核对（不可省）

| 核对项 | 判据 |
|---|---|
| 总额 | 汇总后的总额 = 各来源分组金额之和（用独立路径再算一次，不要只信插件自己的返回值） |
| 行数 | 输出明细行数 = 来源有效数据行数（去重/空值口径要与第 1 步一致） |
| 公式 | 逐行引用指向本行；关键公式的计算结果正确（`calculate` 后 `evaluate_formula`） |
| 图表/透视 | 回读标题、数据源、位置 |
| 原件 | 来源文件内容与任务开始时逐项一致；来源没有被"顺手保存" |
| 落盘 | 输出文件独立打开可读；PDF 大小合理（不是 200 字节的空壳） |

## 7. 验收标准

- [ ] 口径写在报告里，且与实现一致。
- [ ] 每根关键数字能追到具体来源区域。
- [ ] 用户要求的公式仍是公式（没有被改成静态值）。
- [ ] 来源文件未被修改；输出路径唯一且已落盘。
- [ ] 未验证的部分（视觉渲染、超大表抽样）明确标注。

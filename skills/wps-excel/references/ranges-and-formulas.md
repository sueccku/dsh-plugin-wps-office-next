# 表格：区域、公式与重算的实测行为

只读本文件当你要处理**公式**或**大块区域写入**时。工具名与参数以运行时 schema 为准。

## 1. 坐标与区域写法

- 单元格：`{sheet, row, col}`，行列都从 **1** 开始。
- 区域：A1 记法字符串（`A1:C10`、`E1`）。
- `read_range` 的 `include_header: true` 会把首行当表头单独返回。
- `write_range` 的 `data` 是二维数组，从起始单元格向右下填充。数字写成数字（字符串会被当文本写进单元格）。
- **有界优先**：`A1:A1000` 比 `A:A` 便宜得多，也更不容易把表外内容卷进来。

## 2. `set_formula` 的真实行为（重要）

| 目标 | 行为 |
|---|---|
| 单个单元格 | `Range.Formula = formula`，正常 |
| **多格区域** | 桥**逐格**写入**同一个公式字符串**；`Cells(r,c).Formula = $p.formula`，**不做相对引用调整** |

后果：对 `B2:B5` 一次写 `=SUMIF(明细!$A$2:$A$13,A2,明细!$C$2:$C$13)`，B2:B5 会**全部引用 A2**，看起来像公式算错了。

为什么不能像 Excel 那样广播：真机上多格 `Range.Formula = ...` 会报「this object has no property Formula」，`Range.Rows.Count` / 单参数 `Cells.Item(index)` 也都失败，所以只能按地址解析后逐格写。

**需要逐行递增的引用时：**

1. **逐格调用** `set_formula`（每个格子一个自己的公式字符串）——最直接、可核对。
2. 或用 `fill_series`：它同时接受两种写法
   - 已有起始格：`{range:"B2:B5", direction, type, step, startValue}`
   - 源→目标填充：`{sourceRange:"B2", targetRange:"B2:B5"}`
3. **不要**因为一次区域赋值不满足需求，就把用户要的公式整列改成静态数值——那会改掉"可更新"这个需求本身。

写二维字符串数组是否会被 WPS 解释成公式，取决于内容是否以 `=` 开头；**这件事必须先实测确认再依赖**，不要把"应该会"写进方案。

## 3. 写公式之后的核对（三样都要看）

```
wps_excel_get_formula {cell:"B2"}          # 或 get_cell_info：公式文本
wps_excel_calculate {sheet}                # 强制重算（手动计算模式下必须）
wps_excel_evaluate_formula {formula:"=SUM(C2:C4)"}   # 某个表达式的求值结果
wps_excel_read_range {range:"B2:B5"}       # 目标区域的实际值
```

- **公式文本**：确认真的是公式而不是被当成文本写进去了。
- **引用方向**：逐行公式里每一行的引用是不是指向了自己那一行（这是 `set_formula` 区域写入最容易出错的地方）。
- **计算结果**：公式写完不代表算完。手动计算模式下必须 `calculate`；错误值（`#DIV/0!`、`#REF!`…）要读出来，不要当成"还没算"。

`wps_excel_evaluate_formula` 是在**不写入任何单元格**的前提下求值，适合核对，不适合"把结果放进表里"——后者用 `write_range` 写值。

## 4. 值 / 公式 / 数字格式是三个独立结果

- 设「0.00%」不会把 0.1234 变成 12.34；它只是显示格式。
- `read_range` 读到的是值（不是公式文本）；要公式文本用 `get_formula` / `get_cell_info`。
- `copy_format` 只搬格式，不搬值与公式；`clear_formats` 只清格式，保留内容。
- 表的**结构化引用**（`表名[列名]`）需要先 `create_list_object`；`get_list_objects` 会列出每列可用的结构化引用。

## 5. 公式审计与依赖

- `wps_excel_get_formula_audit {cell}`：看它引用了谁（precedents）、谁引用它（dependents），可选画追踪箭头。
- `wps_excel_diagnose_formula`：诊断错误原因。
- `wps_excel_generate_formula {description}`：按自然语言生成公式（`description` 必填）。**生成的公式必须核对**：引用范围、相对/绝对、边界行。
- `wps_excel_refresh_links` / `wps_excel_refresh_all_data`：刷新外部链接与全部数据连接。

## 6. 大范围操作前先量范围

只有下面这几个动作有**执行前的范围预算**（单线程 STA 宿主，一次动作跑太久会让后面所有调用排队，
客户端 60s 后杀掉宿主）：

| 档 | 上限（单元格数） | 动作 |
|---|---|---|
| `read` | 50,000 | `read_range` |
| `scan` | 200,000 | `find_in_sheet`、`find_replace` |
| `layout` | 200,000 | `auto_fit`（整表 / 行 / 列） |
| `percell` | 20,000 | `clean_data`（逐格 COM 往返） |

超了会返回「范围太大：共 N 个单元格，超过……」并要求缩小 `range`。**其余动作没有这道闸门**，
范围要自己控制（`sort_range`、`remove_duplicates`、`transpose`、`consolidate` 都没有自动上限）。

- 排序/筛选/删除重复/分类汇总前，确认区域**包含完整记录**（含表头、不含表外小计行）。
- 转置/写入前确认**目标区域不会覆盖已有数据**：先读目标块。

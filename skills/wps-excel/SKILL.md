---
name: wps-excel
description: 用 WPS 表格工具读写单元格区域、设置格式与数字格式、写公式并核对计算结果、建表与透视表、生成图表、排序筛选与导出。适用于已有工作簿的修改与新建报表。
whenToUse: 任务针对工作表、工作簿、单元格区域、公式与计算结果、图表、透视表、数据清洗、排序筛选或表格导出时用本技能。
---

# WPS 表格

共享的执行协议（目标确认、保存策略、验证与失败分流）见 `wps-office-next` 技能；
本节是表格特有的决策与流程。**应用技能可能被单独加载**，所以下面这些关键规则在本文件里也写了一遍。

<!-- GENERATED:advertised:start -->
本应用注册 **118** 个工具，standard 档直接广告 **46** 个；
完整工具表与按任务分组索引见同目录 `reference.md`（未广告的用 `wps_call`）。
<!-- GENERATED:advertised:end -->

## 1. 先认清三个层级

**应用 → 工作簿 → 工作表 → 区域**。工具作用的目标必须落到「哪一本、哪一张表、哪一块区域」上：

- 应用与工作簿：`wps_excel_get_open_workbooks` 列出所有已打开的工作簿；`wps_excel_switch_workbook` 切换。
- 工作表：`wps_excel_get_sheet_list`（名字 / 序号 / 是否活动）、`wps_excel_switch_sheet`。
- 区域：A1 记法字符串（`A1:C10`）或单元格 `{sheet, row, col}`（行列都从 **1** 开始）。

**只指定 `sheet` 不一定能消除工作簿歧义**：`sheet` 是在**活动工作簿**里解析的。同时打开多本工作簿时，先 `switch_workbook` 点名，再传 `sheet`。
不传 `sheet` 时操作落在**活动工作表**上（跟窗口焦点走）——多工作簿打开时桥会在 `warnings` 里提醒。**养成显式传 `sheet` 的习惯。**

少数动作的目标不参与这套解析：`copy_sheet` 复制到的是**活动工作簿**（`sheet` 只用来选源表）。多工作簿时先点名目标工作簿。

## 2. 动数据之前先读懂数据

1. `wps_excel_get_sheet_info` 看已用范围、表头与当前单元格；需要更细就 `wps_excel_read_range`。
2. 确认**有界的数据区域**：从哪一行开始、到哪一行结束。空行、小计行、合并单元格都会让隐式区域判断出错。
3. 认清每一列是什么：文本、数字、日期、公式、还是汇总。**公式列不能当普通数值列覆盖。**
4. 排序、筛选、删除重复、转置、分类汇总之前，先确认区域**包含完整记录**（含表头、不含表外小计），否则会打乱行的对应关系。
5. 写之前先读要写的那块现有内容，避免覆盖用户数据。

## 3. 标准工作流

```
定位工作簿 / 工作表 → 读取数据结构 → 确定源区域与目标区域
→ 写值 / 写公式 → 设置格式与数字格式 → 重算并核对 → 保存 / 导出
```

- 汇总类任务先定口径：空值怎么算、重复行算不算、日期与数值类型是否真的干净。**用户已经给定口径就直接用**，不要为每个常规选择再问一遍。
- 整块读写优先：一次 `wps_excel_read_range` / `wps_excel_write_range`，不要逐格往返。
- `wps_excel_write_range` 的 `data` 是二维数组，从起始单元格向右下填充；数字要写成数字而不是字符串。

## 4. 值、公式、数字格式是三件不同的事

它们**不能互相代替**。用户要"可更新"，就必须留公式：

- **写值**：`wps_excel_write_range` / `wps_excel_set_cell_value`。
- **写公式**：`wps_excel_set_formula {range, formula, sheet}`。
- **数字格式**：`wps_excel_set_number_format`（`#,##0.00`、`0.00%`、`yyyy-mm-dd`…）。设了格式不会改变值本身。

`set_formula` 对**多格区域**会把同一个公式字符串写进每个单元格，**不做相对引用调整**（与 Excel 的 `Range.Formula` 广播不同）。所以：
逐行递增的引用要**逐格** `set_formula`，或用 `wps_excel_fill_series`（它同时接受 `range` 与 `sourceRange` + `targetRange` 两种写法）。

**不要因为一次区域赋值不满足需求，就把整列公式改成静态数值** —— 那会改掉用户要的可更新能力。细节与核对方法见 `references/ranges-and-formulas.md`。

## 5. 输出策略：改原数据 vs 新建报表

| 意图 | 做法 |
|---|---|
| 改用户现有工作簿 | 就地改，保存原目标（`wps_common_save`） |
| 保留原数据、另出一份 | `wps_common_save_as {outputPath}` 建工作副本，确认后续改动落在副本上 |
| 新建报表 | `create_sheet`（或新建工作簿）→ 写数据 → 建表/透视/图表 → 设打印与导出 → 核对总额 |

跨工作簿汇总必须显式区分**源**和**目的**：源侧读成结构化数据，目的侧点名工作簿与工作表再写。当前接口**无法把目的工作簿绑定到一次写入上**（`sheet` 在活动工作簿里解析），所以多本打开时不要只靠 `sheet` 就认为安全。

## 6. 核对与保存

- 写值：回读目标区域的关键值（或整个小区域）。
- 写公式：**三样都要看** —— 公式文本（`wps_excel_get_formula` / `get_cell_info`）、引用方向、计算结果（必要时 `wps_excel_calculate` 强制重算后 `wps_excel_evaluate_formula`）。
- 改格式：回读实际字体 / 数字格式 / 边框，不要只看工具回了 success。
- 图表与透视：源区域要**含表头行、数据连续无空行**；建完回读它的标题、数据源与位置。
- 导出：`wps_excel_export_chart_as_image` / `wps_excel_export_range_as_image` 依赖剪贴板，并发或无人值守时可能失败——失败要如实上报，不要把"文件路径存在"当成"图是对的"。
- 收尾：按用户意图决定 `wps_common_save` / `wps_common_save_as`；**用户没要求关闭就不要关闭**工作簿。

## 7. 工作表管理

- `create_sheet` / `copy_sheet` / `move_sheet` 的 `position` 是 **0 基**（0 = 最前）；create / copy 不传 `position` 时追加到末尾。
- `delete_sheet` **必须显式给 `name`**，不会退回到活动工作表；工作簿只剩一张表时会拒绝。
- `rename_sheet` 用 `{oldName, newName}`；`switch_sheet` 用 `{name}`。
- 这些工具的返回值里同时给出 0 基 `position`、1 基 `index` 与工作表总数，可以据此核对。

## 8. 已知限制与坑

- 参数名写错会被**顶层**校验拒绝（并列出可选参数）；**嵌套对象不做递归校验**，对象内部多写字段要自己按 schema 核对。
- 值不合法也会报错：不存在的具名样式、未知条件格式名、越界行列等，错误里会给出可用取值。
- `find_replace` 返回的是**命中的单元格数**（不是出现次数），搜索引擎对大小写与全半角敏感。
- 超链接不会覆盖已有单元格的值：给非空单元格设超链接时，显示文本不会替换原内容。
- 插入图片的路径可以相对（按当前工作目录解析）；文件不存在会明确报 `image file not found`。
- `create_chart` 用 `chartType` 指定类型、`range` 指定数据区域；它**没有** `has_header` 参数。
- 公式里避免整列引用（`A:A`），写成 `A1:A1000` 这类有界区域。
- `generate_formula` 需要 `description` 参数。
- 有**范围预算**的动作会在执行前拦下过大的区域：`read_range`（读，50,000 格）、`find_in_sheet` / `find_replace`（扫描，200,000 格）、`auto_fit`（排版，200,000 格）、`clean_data`（逐格写，20,000 格）。超了会要求给出更小的 `range`。其余动作没有这道闸门，范围要自己控制。

## 9. 按需参考

| 想了解 | 读哪里 |
|---|---|
| 全部表格工具与按任务分组索引 | 同目录 `reference.md` |
| 区域公式、相对引用、`fill_series`、重算与错误值的实测行为 | `references/ranges-and-formulas.md` |
| 汇总报表的完整任务实例（口径、公式策略、图表、打印、核对、原件保护） | `recipes/summarize-report.md` |
| 共享执行协议 | `wps-office-next` 技能 |

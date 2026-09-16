# 破坏性操作清单与前置守卫（S3）

> 目的：破坏性动作动的是客户真金白银的数据。插件必须在动手前后说清「改了什么、失去了多少」，
> 而不是回一个「成功」却没有依据。
> 本文件是 S3 的验收物：**动作 / 影响范围 / 是否回传统计 / 是否抑制弹窗 / 对应测试**。
> 清单数据来源：2026-09-16 对 `mcp/scripts/wps-com.ps1` 的实测枚举（不是记忆）。
> 相关设计见 `docs/stabilization-plan.md` 的 S3，实现记录见 `docs/FIXES.md` 第 53 条。

## 机制

桥头部新增两个**只读**助手（位于 `mcp/scripts/wps-com.ps1` 头部，会一并进入生成物
`host/wps-actions.ps1`）：

- `Get-WpsRangeImpact($range, $maxPreview = 8)` → `{ address, cells, nonEmpty, preview }`
  - `cells`：范围格子数（`Range.Count`）。
  - `nonEmpty`：非空格子数——**真正会被抹掉的数据量**。首选 `WorksheetFunction.CountA`，
    回退 `Application.CountA`，最后才是小范围手工扫描（≤ 10000 格）。
  - `preview`：内容前 8 项示例；**仅当范围 ≤ 4096 格**才读 `Value2`（避免一次读上百万格把宿主拖住）。
  - 每一步失败都降级为 `$null`：**守卫本身绝不把破坏性动作弄失败**。
- `Get-WpsRangeNonEmpty($range)`：上面的计数部分。

TS 侧 `mcp/src/tools/excel/impact.ts` 把 `impact` 压成一句中文，附在动作结果后面，让模型能直接复述：

```
范围 A1:C3 的内容已清除（原有 9 个非空单元格；示例内容：地区 / 产品 / 金额 / 华东 / A）
工作表 "VictimS3" 已成功删除（原有 4 个非空单元格；示例内容：1 / 2 / 3 / 4）
```

## 清单

图例：**✅ 已回传统计**（本批实现并有断言）；**⬜ 待办**（S3 剩余）；**—** 非用户数据（内部/临时对象）。

| 动作 | 桥位置 | 影响范围 | 回传统计 | 弹窗抑制 | 对应测试 |
| --- | --- | --- | --- | --- | --- |
| `clearRange` | 2413–2416 | 清除区域全部/内容/格式/批注 | ✅ impact | 无确认框（未触发） | destructive-guard |
| `clearFormats` | 1666 | 清除区域格式（内容保留） | ✅ impact | 无确认框（未触发） | destructive-guard |
| `deleteRows` | 2454 | 删除整行 | ✅ impact（原有 count 保留） | 无确认框（未触发） | destructive-guard |
| `deleteColumns` | 2467 | 删除整列 | ✅ impact（原有 count 保留） | 无确认框（未触发） | destructive-guard |
| `deleteSheet` | 2188 | 删除整张工作表 | ✅ impact（原有 remaining 保留） | 是（原有 `DisplayAlerts`） | destructive-guard / sheet-ops |
| `deleteListRow` | 2677 | 删除表（ListObject）一行 | ⬜ | 待核 | — |
| `unlistListObject` | 2751 | 表转回普通区域（丢结构） | ⬜ | 待核 | — |
| `deleteTableLine` | 3264 | Word 表格删除一行/一列 | ⬜ | 待核 | — |
| `deleteComment` | 3495 / 3501 | 删除 Word 批注（单条/全部） | ⬜ | 待核 | — |
| `deleteCellComment` | 4125 | 删除单元格批注 | ⬜ | 待核 | — |
| `addCellComment` | 4114 | **覆盖**已有单元格批注 | ⬜ | 待核 | — |
| `deleteNamedRange` | 4088 | 删除命名范围 | ⬜ | 待核 | — |
| `deleteChart` | 3120 | 删除图表 | ⬜ | 待核 | — |
| `clearPivotTable` | 3031 | 清除透视表 | ⬜ | 待核 | — |
| `clearSparkline` | 3097 | 清除迷你图 | ⬜ | 待核 | — |
| `resetPageBreaks` | 2922 | 重置全部分页符（丢打印分页） | ⬜ | 待核 | — |
| `removeConditionalFormat` | 1757 / 1759 | 删除条件格式规则 | ⬜ | 待核 | — |
| `removeDataValidation` | 1820 | 删除数据验证 | ⬜ | 待核 | — |
| `addDataValidation` | 1787 | 覆盖前先删除已有验证 | ⬜ | 待核 | — |
| `deleteSlide` | 5071 | 删除 PPT 幻灯片 | ⬜ | 待核 | — |
| `deleteShape` | 5436 | 删除 PPT 形状 | ⬜ | 待核 | — |
| `deleteTextBox` | 5279 | 删除 PPT 文本框 | ⬜ | 待核 | — |
| `deletePptImage` | 5527 | 删除 PPT 图片 | ⬜ | 待核 | — |
| `replacePptImage` | 5548 | 替换图片（删旧图） | ⬜ | 待核 | — |
| `removeAnimation` | 6068 / 6070 | 删除 PPT 动画（单条/全部） | ⬜ | 待核 | — |
| `cleanData` | 1322 | 只删空行 | — | — | — |
| `updateChart` | 2061 | 删旧数据标签（外观） | — | — | — |
| `exportRangeAsImage` | 2116 / 2121 | 临时图表清理 | — | — | — |
| `evaluateFormula` | 6400 | 临时探针格清理 | — | — | — |

> 计数口径：计划写的是「27 处 `Delete()`」。实测把 `Clear()` / `ClearFormats()` /
> `ClearContents()` / `Unlist()` / `ResetAllPageBreaks()` 一并算上是 **35 个站点 / 29 个动作**；
> 其中 4 个动作属内部或临时对象（`cleanData` 只删空行），不涉及用户数据。

## 已知边界（如实记录）

- `preview` 对整行/整列（> 4096 格）为**空数组**，只回 `cells` 与 `nonEmpty`；这是刻意的性能取舍。
  实测：整列 `cells = 1048576`、整行 `cells = 16384`——范围本身就这么大，**`nonEmpty` 才是有意义的数据量**。
- `nonEmpty` 依赖 WPS 的 `CountA`。实测在本机可用；三级回退保证它失效时只是数值缺失，不会让动作挂住。
- **弹窗抑制尚未逐项实测**：目前只在已知会弹确认框的路径上做了抑制（`deleteSheet`，以及 S1 的打开文档助手）。
  其它破坏性动作是否会弹确认框，是 S3 的剩余工作之一。

## 运行

```powershell
node test/destructive-guard.test.mjs
```

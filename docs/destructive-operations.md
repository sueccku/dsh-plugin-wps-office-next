# 破坏性操作清单与前置守卫（S3）

> 目的：破坏性动作动的是客户真金白银的数据。插件必须在动手前后说清「改了什么、失去了多少」，
> 而不是回一个「成功」却没有依据。
> 本文件是 S3 的验收物：**动作 / 影响范围 / 是否回传统计 / 是否抑制弹窗 / 对应测试**。
> 清单数据来源：2026-09-16 对 `mcp/scripts/wps-com.ps1` 的实测枚举（不是记忆）。
> 相关设计见 `docs/stabilization-plan.md` 的 S3，实现记录见 `docs/FIXES.md` 第 53 / 54 / 55 条。

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

对象级动作（表行、图表、命名范围、迷你图、分页符、批注、幻灯片、形状、图片、动画…）在同一条
`data.impact` 里带上 `{ name, kind, count, rows, detail, address }`。

TS 侧 `mcp/src/tools/impact.ts`（跨 Excel/Word/PPT 共用）把 `impact` 压成一句中文，附在动作结果后面，
让模型能直接复述：

```
范围 A1:C3 的内容已清除（原有 9 个非空单元格；示例内容：地区 / 产品 / 金额 / 华东 / A）
已删除该行（表行「S3Table」；原有 3 个非空单元格；示例内容：华东 / A / 120）
已删除图表 Chart 1（该表还剩 0 张）（图表「Chart 1」）
图片删除成功！（图片「图片 4」）
```

## 清单

图例：**✅ 已回传统计**；**—** 非用户数据（内部/临时对象）。
当前进度：**25 / 25 个用户数据动作已回传统计**（35 个站点，其中 4 个动作只动内部对象）。

| 动作 | 应用 | 影响范围 | 回传统计 | 对应测试 |
| --- | --- | --- | --- | --- |
| `clearRange` | Excel | 清除区域全部/内容/格式/批注 | ✅ 范围 | destructive-guard |
| `clearFormats` | Excel | 清除区域格式（内容保留） | ✅ 范围 | destructive-guard |
| `deleteRows` | Excel | 删除整行 | ✅ 范围 | destructive-guard |
| `deleteColumns` | Excel | 删除整列 | ✅ 范围 | destructive-guard |
| `deleteSheet` | Excel | 删除整张工作表 | ✅ 范围 | destructive-guard / sheet-ops |
| `deleteListRow` | Excel | 删除表（ListObject）一行 | ✅ 该行内容 | destructive-guard |
| `unlistListObject` | Excel | 表转回普通区域（丢结构） | ✅ 名称/范围/行数 | destructive-guard |
| `resetPageBreaks` | Excel | 重置全部分页符 | ✅ 清除前横/纵计数 | destructive-guard |
| `clearPivotTable` | Excel | 清除透视表报表 | ✅ 名称/原区域 | 既有 excel-advanced |
| `clearSparkline` | Excel | 清除迷你图 | ✅ 清除前组数 | destructive-guard |
| `deleteChart` | Excel | 删除图表 | ✅ 图表名 | destructive-guard |
| `removeConditionalFormat` | Excel | 删除条件格式规则 | ✅ 规则条数 | destructive-guard |
| `deleteNamedRange` | Excel | 删除命名范围 | ✅ 原引用地址 | destructive-guard |
| `addDataValidation` | Excel | 覆盖前删除已有验证 | ✅ 是否覆盖 + 原类型 | destructive-guard |
| `removeDataValidation` | Excel | 删除数据验证 | ✅ 原类型 | destructive-guard |
| `addCellComment` | Excel | **覆盖**已有单元格批注 | ✅ 是否覆盖 + 原文 | destructive-guard |
| `deleteCellComment` | Excel | 删除单元格批注 | ✅ 原文 | destructive-guard |
| `deleteTableLine` | Word | 删除表格一行/一列 | ✅ 删除前文字 | destructive-guard |
| `deleteComment` | Word | 删除批注（单条/全部） | ✅ 条数 + 原文 | destructive-guard |
| `deleteSlide` | PPT | 删除幻灯片 | ✅ 页号 + 删除前形状数 | destructive-guard |
| `deleteShape` | PPT | 删除形状 | ✅ 形状名 | destructive-guard |
| `deleteTextBox` | PPT | 删除文本框 | ✅ 名称 + 文字长度 | destructive-guard |
| `deletePptImage` | PPT | 删除图片 | ✅ 图片名 | destructive-guard |
| `replacePptImage` | PPT | 替换图片（删旧图） | ✅ 旧图名 | destructive-guard |
| `removeAnimation` | PPT | 删除动画（单条/全部） | ✅ 条数 | destructive-guard |
| `cleanData` | Excel | 只删空行 | — | — |
| `updateChart` | Excel | 删旧数据标签（外观） | — | — |
| `exportRangeAsImage` | Excel | 临时图表清理 | — | — |
| `evaluateFormula` | Excel | 临时探针格清理 | — | — |

> 计数口径：计划写的是「27 处 `Delete()`」。实测把 `Clear()` / `ClearFormats()` /
> `ClearContents()` / `Unlist()` / `ResetAllPageBreaks()` 一并算上是 **35 个站点 / 29 个动作**；
> 其中 4 个动作属内部或临时对象，不涉及用户数据，剩下 **25 个用户数据动作已全部回传统计**。

## 已知边界（如实记录）

- `preview` 对整行/整列（> 4096 格）为**空数组**，只回 `cells` 与 `nonEmpty`；这是刻意的性能取舍。
  实测：整列 `cells = 1048576`、整行 `cells = 16384`——范围本身就这么大，**`nonEmpty` 才是有意义的数据量**。
- `nonEmpty` 依赖 WPS 的 `CountA`。实测在本机可用；三级回退保证它失效时只是数值缺失，不会让动作挂住。
- WPS 的 `Comment.Text` 是**方法**而不是属性（直接读会拿到方法签名），桥里统一写成 `.Text()`。
- PPT `AddPicture` 需要**绝对路径**（相对路径报 `The specified protocol is unknown`）。
- `clearPivotTable` 的模型可见文案沿用原有报告（名称 + 原区域），未再叠加 impact，避免重复。
- `removeAnimation` 的计数以 WPS 自己报告的 `MainSequence.Count` 为准（实测给一个形状加入场效果后为 2）。
- **弹窗抑制尚未逐项实测**：目前只在已知会弹确认框的路径上做了抑制（`deleteSheet`，以及 S1 的打开文档助手）。
  其它破坏性动作是否会弹确认框，是 S3 的剩余工作之一。

## 运行

```powershell
node test/destructive-guard.test.mjs
```

当前 **45 项**断言（真实 WPS）：范围类 5 + 对象类 8 + 批注/验证/Word/PPT 12，逐个验证 impact 与模型可见摘要。

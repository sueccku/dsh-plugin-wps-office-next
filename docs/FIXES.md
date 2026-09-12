# 修复记录

对应 baseline/known-defects.md 的编号，记录本仓库相对上游 a825336 的修复。
每条都要求有可复跑的验证，不接受“看起来对”。

## 已修复

### 1. 11 个工具在 Windows 上永久失败（缺陷 P1 及同类）

上游工具层引用了 10 个 action，但 wps-com.ps1 里**完全不存在这些字符串**——
源码注释却写着“macOS未实现，仅Windows支持”。结果是这 11 个工具每次都返回
Unknown action，其中 wps_ppt_set_shape_fill 还在我们的 standard 档里被广告。

| action | 涉及工具 | 实现要点 |
|---|---|---|
| setZoom | wps_excel_set_zoom、wps_excel_zoom | ActiveWindow.Zoom，范围校验 10-400 |
| setLineSpacing | wps_word_set_line_spacing | LineSpacingRule=5（倍数），可选 paragraphIndex（0 基） |
| autoSum | wps_excel_auto_sum | 写入 =SUM(range)，回读公式与结果 |
| evaluateFormula | wps_excel_evaluate_formula | Application.Evaluate，失败回落到临时单元格 |
| setSlideTheme | wps_ppt_set_slide_theme | ApplyTemplate；要求真实模板路径，否则明确报错 |
| setSlideSize | wps_ppt_set_slide_size | 像素→磅换算（x0.75），回读实际尺寸 |
| setShapeFill | wps_ppt_set_shape_fill | Fill.Solid + ForeColor.RGB |
| setFontColor | wps_ppt_set_font_color | TextFrame.TextRange.Font.Color.RGB |
| setTextColor | wps_word_set_text_color | 十六进制或颜色名，非法值明确报错 |
| insertSectionBreak | wps_word_insert_section_break | 四种分节符映射，回读节数 |

### 2. closePresentation 无法放弃更改

上游实现直接调用 $pres.Close()，丢弃修改时会弹保存对话框，破坏无人值守场景。
现在与 Excel/Word 一致：saveChanges 为 false 时先置 Saved=$true 再关闭。

### 3. 省略 sheet 时必然失败（缺陷 C4，影响 6 处）

上游把省略的 sheet 强制成索引 0：`$p.sheet -is [int]` 的两个分支代码完全相同，
最终都走 `Sheets.Item(0)`，触发 DISP_E_BADINDEX。而 wps_excel_read_range / write_range 的
schema 里 sheet 是可选参数——也就是说**最常用的读写工具在不指定工作表时必然失败**。

已实测确认：`getRangeData {range:"A1:B2"}` 返回 Invalid index (0x8002000B)。
现在 6 处统一改为：传了就按名/序号取，没传就用 `$excel.ActiveSheet`。
同时修掉工具层的 `sheet || 0` 强制转换。

### 4. 范围读取逐格循环（缺陷 C5）+ 二维数组编组

原实现按 cell 逐个读（O(n*m) 次 COM 往返），根因是绕开 `Range.Value2` 返回 `Object[,]`
时 `ConvertTo-Json` 会把它压平的问题。

现在改为一次 `$range.Value2` 读取，再用 `GetLowerBound()/GetUpperBound()` 显式转成锯齿数组
（不能硬编码 0 或 1 基），并处理单格标量与空区域两种情况。

实测：读取 400x20 = 8000 格耗时 **37ms**；原实现是 8000 次 COM 往返。

## 新发现的 WPS / Office 差异

- **WPS 的 Presentations.Add() 返回 0 页演示文稿**，PowerPoint 返回 1 页。
  任何“新建演示文稿后立刻操作第 1 页”的流程都会 E_FAIL。
  技能文档必须写明：新建后先 add_slide。
- Range.Value2 返回 Object[,]，ConvertTo-Json 会把它压平；
  必须显式转换并用 GetLowerBound() 判断下界（这也解释了上游为何逐格读）。

## 验证

两个测试都在真实 WPS 上跑，各自创建一次性文档、回读校验、不保存关闭：

- node test/new-actions.test.mjs —— 28 项（新增/修复的 10 个 action 与 closePresentation）
- node test/excel-range.test.mjs —— 10 项（sheet 回退、二维编组、单格、8000 格性能）
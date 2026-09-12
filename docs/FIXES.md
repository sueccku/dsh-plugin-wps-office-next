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

## 新发现的 WPS / Office 差异

- **WPS 的 Presentations.Add() 返回 0 页演示文稿**，PowerPoint 返回 1 页。
  任何“新建演示文稿后立刻操作第 1 页”的流程都会 E_FAIL。
  技能文档必须写明：新建后先 add_slide。
- Range.Value2 返回 Object[,]，ConvertTo-Json 会把它压平；
  必须显式转换并用 GetLowerBound() 判断下界（这也解释了上游为何逐格读）。

## 验证

node test/new-actions.test.mjs —— 28 项，全部在真实 WPS 上执行：
创建一次性文档 → 逐项调用并回读校验 → 不保存关闭。

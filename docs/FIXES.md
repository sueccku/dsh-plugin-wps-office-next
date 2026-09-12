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

### 5. 重复工具合并（第一批，3 对）

上游有 18 组「同应用、同一 action」的工具对。其中接口完全等价的先合并：

| 废弃名 | 规范名 | 依据 |
|---|---|---|
| wps_excel_zoom | wps_excel_set_zoom | 参数与语义完全相同（都只收 percent） |
| wps_ppt_add_speaker_notes | wps_ppt_set_slide_notes | 参数与语义完全相同（slideIndex, notes） |
| wps_word_generate_doc_toc | wps_word_generate_toc | levels 是规范参数的子集 |

实现方式：mcp/src/tools/deprecated.ts 登记废弃→规范映射，启动时把废弃名改成转发别名。
旧名字仍然可用（不破坏既有提示词），但不再出现在 wps_help 的目录与搜索里；
wps_status 会汇报合并数量。验证：test/deprecated.test.mjs 8/8。

### 6. 未合并的 15 对：为什么先不动

这些对的参数接口不同（例如 delete_rows 用 startRow、delete_row 用 row；
fill_series 与 auto_fill 是两种不同操作），**机械转发会改变行为**。
更关键的是：继续合并前必须先修底层，否则只是把坏工具合并成一个更权威的坏名字。

## 已确认不可用的整族功能：PPT 动画与切换

实测（真实 WPS，未保存关闭）：

```
SlideShowTransition.EntryEffect = 'fade'      -> FAIL: Index was outside the bounds of the array
SlideShowTransition.EntryEffect = 1793        -> OK
TimeLine.MainSequence.AddEffect(sh,'fadeIn',1) -> FAIL: Cannot convert "fadeIn" (string) to type Object
TimeLine.MainSequence.AddEffect(sh, 10, 1)     -> OK
```

**COM 要求数字枚举，而上游整族传的是英文名**，所以 animate/transition 相关工具全部不可用：
wps_ppt_add_animation、wps_ppt_set_animation、wps_ppt_add_animation_preset、
wps_ppt_add_emphasis_animation、wps_ppt_set_slide_transition、wps_ppt_set_transition、
wps_ppt_apply_transition_to_all。

另外 wps_ppt_set_transition / wps_ppt_set_animation 连参数名都不对：
桥读 effect / shapeName，它们传 transition / shapeIndex。

## 待办：参数契约测试

工具层与桥之间存在系统性的参数名不一致。静态比对（scripts/analyze-param-mismatch.mjs）
会因对象展开与改名产生大量误报，**不足以作为结论**；且实测也证明 handler 会主动补别名
（wps_ppt_insert_ppt_image 同时发 filePath/path/imagePath）。

正确做法是逐工具动态验证：用文档化参数调用，再回读可观测的效果。

### 7. 单元格写入：PowerShell 的 COM 绑定器缓存（高危，已修）

现象：`wps_excel_write_range` 对**任何含数字的数据**都失败，只有纯字符串能写；
`wps_excel_set_cell_value` 同样。而它们是广告中的核心工具。

根因分两层：

1. `ConvertFrom-Json` 产出的数值赋给 COM 的 `Value2` 会抛 `InvalidCastException`。
2. 更隐蔽的是：**PowerShell 5.1 对 COM 成员按首次调用缓存绑定器**。
   同一个 `Value2` 调用点，先赋字符串则之后数字失败，先赋数字则之后字符串失败。
   实测：同一进程内 "Sheets.Item -> OK"（字符串）之后数字三种目标全部失败；
   反过来在模块里先写了数字，之后字符串/布尔全部失败。

修复：写入改走 `$cell.PSObject.Properties['Value2'].Value = $v`。
PSPropertyInfo 每次调用重新绑定，数字/字符串/小数/布尔混合均正常。
区域写入同时改为解包成 typed 矩阵后**一次整体赋值**（O(n*m) 次 COM 往返降为 1 次）。

验证：setCellValue 六种类型往返正确；混合类型区域写入精确往返；
excel-range 12/12；8000 格读取 15ms。

### 8. setCellValue 行列参数解包

`Cells.Item($p.row, $p.col)` 的参数同样是 PSObject 包装的数值，显式 `[int]` 转换后稳定。

### 9. Excel 查找替换：跨应用缺陷 + 静默无效 + 计数造假（已修）

`wps_excel_find_replace` 同时有三个缺陷：

1. **跨应用**：它调用的是 Word 专用的 `findReplace` action（实现里是 `Get-WpsWord` + `$doc.Content.Find`），
   所以"在 Excel 里查找替换"实际作用在 **Word 文档**上。
2. **静默无效**：它发 `find`/`replace`，而桥读 `findText`/`replaceText`，字符串从未被传入。
3. **计数造假**：它期望 `data.count`，而 Word action 返回的是 `replaced` 布尔值，于是永远显示"共替换 0 处"。

修复：

- 桥里新增独立的 **`findReplaceExcel`** action，Excel 工具改调它，不再复用 Word 的实现。
- Word 的 `findReplace` 同时接受 `find`/`replace` 别名（容错）。
- 计数改为一次 `Value2` 读取后在 PowerShell 内精确统计（Excel 的 `Find`/`FindNext` 受 COM 绑定器缓存影响不可靠），
  替换仍交给 `Range.Replace` 执行；返回 `cells` 与 `changed`。
- 顺带修掉 Word 工具消息里的 "替换了 undefined 处"。

回归：test/find-replace.test.mjs 10/10，其中包含**跨应用隔离**用例：
在 Word 文档打开的状态下执行 Excel 查找替换，Word 内容必须保持不变。

### 10. PowerShell COM 稳定性加固（A 方案，已落地）

先澄清一个判断：**这不是 PowerShell 版本问题**。宿主固定使用 Windows PowerShell 5.1
（Windows 10/11 自带、全平台同版本），而踩的坑是 5.1 COM 适配器的固有语义：

1. **成员绑定器按首次调用缓存** —— 同一个调用点先写字符串，之后再写数字就失败（反之亦然）。
2. **成员集合惰性填充** —— `PSObject.Methods['X']` 在集合被枚举过之前会返回 null。

也就是说，这不是"这台好、别台坏"，而是**所有机器上同一颗地雷**，因此修法是确定性的。

落地措施：

| 措施 | 内容 |
|---|---|
| COM 边界 helper | 新增 `ConvertTo-ComValue` / `Set-ComValue` / `Get-ComValue` / `Invoke-ComMethod`，统一解包类型并**枚举成员**而非索引 |
| 解释器固定 | 宿主改用绝对路径 `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`，不再依赖 PATH |
| 版本断言 | 宿主在 ready 帧回报 `psVersion`，非 5.x 直接拒绝启动；`__env` 控制帧暴露 psVersion/pid/exe |
| 机器一致性自检 | `scripts/com-conformance.mjs`：在本机跑 11 项 COM 行为矩阵（数字/字符串/小数/布尔交替写同一单元格、二维矩阵往返、Word 数值参数、PPT 形状与颜色…），不符即报错 |
| 边界 lint | `scripts/lint-com-boundary.mjs`：统计仍在边界层外的 COM 访问点（当前 value-write 63、com-call 4 待随工具修复逐步迁移；colour-write 72 恒为整数、无风险） |

一个实测结论：`Range.Replace` 的 PSObject 方法查找在宿主里不可用（惰性填充），
因此该成员保留直接调用，但**两处调用点共用同一参数签名**（全部显式 `[string]`/`[int]`/`[bool]`），
使缓存绑定器对两者都有效。这是探针验证可行的路径。

证据：本机一致性自检 11/11；全量回归全绿。

### 11. `wps_excel_set_cell_format`：整个格式对象被丢弃（已修）

这是广告中的常用工具，实际行为是：

- 桥只读 `numberFormat`，`format` 与 `sheet` 被忽略，返回却是 `success: true`；
- 工具的 schema **同时**声明了嵌套 `format` 对象和 12 个平铺属性（bold/fontSize/...），
  但 handler 只转发 `format`，所以平铺写法也被静默忽略——同一个功能有两套入口，两套都不可用。

修复：

- 桥实现完整的格式应用：粗体、斜体、下划线、删除线、字体名、字号、字色、背景色、
  水平/垂直对齐、自动换行、数字格式；支持嵌套与平铺两种形状；支持 `sheet` 指定工作表。
- 工具层把平铺属性合并进 `format`，并在返回消息里报告实际生效的属性列表。
- 空格式**明确报错**（`no supported format property was provided`），不再假装成功。
- 顺带扩展 `getCellInfo` 的回读字段（italic、字体色、对齐、换行、sheet），使格式可验证。
- 顺带加固 helper：新增 `Find-ComProperty`，先索引后枚举，应对成员集合惰性填充。

回归：test/cell-format.test.mjs 16/16，包含 10 个属性逐个回读、平铺写法、
空格式必须报错、以及**不串到相邻单元格**。

### 12. 工作表操作组 `delete`/`rename`/`copy`/`move`/`switch` 作用在错误的工作表上（已修）

这一组是本轮最危险的一类缺陷，因为它们的失败方式是**静默改错对象**，不是报错。

根因：工具的 schema 用 `name` / `oldName` 命名参数，桥只读 `sheet`。
参数名不匹配 → 桥取不到名字 → 回退到“当前活动工作表”。于是：

- `delete_sheet(name="报表")` 删掉的是当时恰好激活的那张表；
- `rename_sheet` / `copy_sheet` / `move_sheet` 同样作用在活动表上；
- `switch_sheet` 因为拿不到名字**必然失败**（活动表没变，它自己也不知道要切到哪）。

修复：

- 桥统一用 `Resolve-Worksheet` 解析目标，接受 `sheet` / `name` / `oldName` 三种键名；
- `delete_sheet` 加 `-RequireName`：**没有显式目标就报错**，绝不回退到活动表；
  且拒绝删除工作簿里最后一张表；
- 其余五个工具补齐参数，响应键与工具层期望对齐；
- `set_number_format` 的 `sheet` 参数此前也被忽略，一并修好。

顺带修掉一个**契约自相矛盾**：`position` 在 schema 里写明 0 基，
但桥回显的是 Excel 的 1 基 `Sheet.Index`，工具层又 `+1`，
导致 `position: 0` 的返回消息是“位置: 第2个”——模型照着读会算错。
现在桥同时返回 0 基 `position`、1 基 `index`、`sheetCount`，
工具层只回显 0 基值并附总数；负数位置明确报错；
创建/复制省略 `position` 时按 schema 承诺**追加到末尾**（此前创建反而插到最前）。

回归：test/sheet-ops.test.mjs 21/21，其中最关键的一项是
“删除指定表时，活动表必须原封不动”。

## 新发现的 WPS / Office 差异

- **WPS 的 Presentations.Add() 返回 0 页演示文稿**，PowerPoint 返回 1 页。
  任何“新建演示文稿后立刻操作第 1 页”的流程都会 E_FAIL。
  技能文档必须写明：新建后先 add_slide。
- Range.Value2 返回 Object[,]，ConvertTo-Json 会把它压平；
  必须显式转换并用 GetLowerBound() 判断下界（这也解释了上游为何逐格读）。

## 验证

全部测试都在**真实 WPS** 上跑：各自创建一次性文档、回读校验、不保存关闭。

| 测试 | 项数 | 覆盖 |
| --- | --- | --- |
| test/com-host.test.mjs | 6 | 常驻宿主握手、就绪帧、串行队列、重启 |
| test/plugin.test.mjs | 32 | 插件注册、skills、工具面过滤、预算 |
| test/new-actions.test.mjs | 28 | 本轮新增/修复的 10 个 action、closePresentation |
| test/excel-range.test.mjs | 12 | sheet 回退、二维编组、单格、8000 格性能 |
| test/find-replace.test.mjs | 10 | 查找替换的参数名契约 |
| test/cell-format.test.mjs | 16 | 10 个格式属性逐个回读、平铺写法、空格式报错 |
| test/deprecated.test.mjs | 8 | 弃用名合并与隐藏 |
| test/sheet-ops.test.mjs | 21 | 工作表组目标正确性、0 基 position、删表安全 |

合计 133 项，加 `node scripts/verify.mjs` 22 项门禁（含 45 工具 / 25,000 字节预算）。
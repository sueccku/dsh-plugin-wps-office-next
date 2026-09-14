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

### 13. 参数名不匹配只会静默失效——一类缺陷，不是一堆 bug（已加闸门）

第 9～12 条的根因是同一个：**MCP 不校验参数名，PowerShell 也不会抱怨没人读的属性**。
工具发 'name'、桥读 'sheet'，于是操作落到活动工作表上；工具发 'save'、桥读 'saveChanges'，
于是 'save=false' 被丢掉并弹出模态框。每一个都要靠人工逐工具实验才能发现，这才是真正的成本。

现在改成机器发现：

- 生成器（scripts/build-host-actions.ps1）静态扫描桥的 switch，
  **推导出每个 action 真正读取的参数名表**（含 Resolve-Worksheet / Get-TargetPres /
  Get-WorksheetByParam 这三个共享解析器贡献的键），并写进生成产物。
  读法无法静态确定的 action（当前仅 setCellFormat，因为它遍历嵌套的 format）**排除在表外**，
  因此闸门不可能误拒一个确实会被读取的键。
- 桥在 $p 生成后立刻比对：出现该 action 不读的键就**明确报错**，
  并列出被拒的键与该 action 接受的键。静默失效变成响亮失败。
- 宿主新增只校验不执行的 __validateParams 控制帧；
  scripts/param-contract.mjs 用它在**不碰 COM** 的前提下，把全部 254 个工具的
  handler 实参对象与桥的键表对账，并生成 docs/param-contract.md。

覆盖：251/256 个 action 进入键表（1 个动态跳过），254 个工具中 212 对完成对账。
闸门上线当天即抓出 findReplace 的 replaceMode（见第 14 条），并量出其余 85 处
「工具发了、桥不读」的错配——它们此前全都是静默的。

顺带把最大的一类修掉：约 50 个 Excel action 无视工具传来的 'sheet'、一律操作活动表。
桥新增 Get-WorksheetByParam（只读 $p.sheet，不把 'name' 当表名，因为若干 action 的 name
另有含义），替换了 58 处 '$sheet = $excel.ActiveSheet' 兜底。

### 14. Word 查找替换：**纯查找会删掉所有命中内容**（已修）

这是本轮唯一的**数据丢失**级缺陷。

- 工具在「只查找」时也传 replaceText: ''，桥把它当替换文本，
  于是 Find.Execute(..., 替换为 '') 把每一处命中**原地删除**；
- 桥只返回 { replaced = ... }，而工具读的是 count（永远 undefined），
  所以返回消息里的次数是编造的，「未找到」分支也永远走不到。

修复：

- 桥读 replaceMode；无该标志时，**空替换文本一律按「只查找」处理**——
  不能把「删除命中」当作没写替换文本的默认行为；
- 新增 Get-WordMatchCount：用 Find 走一遍文档统计命中数，
  Wrap=wdFindStop 且有上限，不允许 WPS 的怪癖把常驻宿主挂住；纯查找只计数、不改文档；
- 替换模式先计数再替换，返回真实 count（replaceAll=false 时为 0/1）；
- 工具层按 replaceMode 决定是否发送 replaceText，并如实回报次数。

回归：test/find-replace.test.mjs 14/14，其中新增
**「纯查找不得删除命中内容」**与「无命中时报 0 次」两项。

### 15. Excel 参数契约批次：25 处「工具发了、桥不读」全部修掉（已修）

第 13 条的闸门把 Excel 侧的错配一次列全，共 25 处。按性质分三类处理：

**一、工具发的键名与桥不同（桥接受两种拼写）**

| 工具 | 工具发的 | 桥原本读的 |
| --- | --- | --- |
| set_hyperlink | url / text | address / textToDisplay |
| auto_filter | column | field |
| transpose | source / destination | sourceRange / destinationCell |
| subtotal | columns | totalColumns |
| sort_range | column / ascending | keyColumn / order |
| set_border | borderStyle | style |
| set_data_validation | type / formula | validationType / formula1 |
| hide_row / hide_column | count / hide | （没有对应能力） |
| hide_rows / show_rows | startRow / endRow | row / rows |
| show_columns | startColumn / endColumn | column / columns |
| freeze_panes | freeze | （没有取消冻结） |
| protect_sheet / protect_workbook | protect | （没有取消保护） |
| fill_series | direction | （没有方向） |
| auto_fill | sourceRange / targetRange | （没有自动填充） |
| set_conditional_format | condition / format | operator / value / 颜色 |
| set_cell_style | style（具名样式） | 逐个属性 |

**二、能力确实缺失，补实现**

- `copy_range` 只把 `range` 复制到剪贴板，`source`/`destination` 被丢弃，**从未粘贴到任何地方**；
  改为真正的 `source→destination` 复制；
- `get_cell_comments` 的 `range` 被忽略，总是返回整表批注，现按范围过滤；
- `insert_excel_image` 的 `cell` 被忽略，图片位置参数失效；现按单元格锚定；
- `add_comment` **是 Word 的 action**（用 `$word.Selection.Range`），
  所以 `wps_excel_add_comment` 一直把批注写进**当前 Word 文档**；
  改指向 Excel 自己的 `addCellComment`（原本就存在，只是没人调用），并接受工具公开的 `comment` 键；
- 58 处 Excel action 的兜底工作表改为读 `sheet`（第 13 条），
  其中采用 `Get-WorksheetByParam`，只认 `sheet`、不把 `name` 当表名，
  因为 `set_named_range` 等 action 的 `name` 另有含义。

**三、schema 宣称了却无法实现，删掉**

- `wps_excel_create_chart` 的 `has_header`：桥从未读取，且 Excel 图表没有「首行是否表头」开关，
  无法兑现。删掉比留着假装支持更诚实（schema 因此少 92 字节）。

回归：test/excel-contract-fixes.test.mjs **35 项**，全部在真实 WPS 上写→改→回读，
其中包含「批注落在 Excel 而不是 Word」的跨应用检查。

### 本轮实测到的 WPS 与 Office 差异（新增）

- **`Range("2:2").Hidden` 在 WPS 上恒抛 E_FAIL**，`Rows("2:3").Hidden` / `Columns("B:C").Hidden` 正常。
  上游四个隐藏/显示工具在 WPS 上**从来没有生效过**，现改用 Rows/Columns；
- **`Range.DataSeries` 的 Rowcol 必须留空**：纵向区域传 1（xlRows）静默不填充，
  传 2 或留空（自动判断）才填。上游留空是对的，我一度改成 1 反而弄坏，已回退并加注释；
- **`Shapes.AddPicture` 只接受绝对路径**：相对路径抛裸 E_FAIL。
  桥现在解析相对路径并校验存在性，缺文件时报明确错误；
- **`Hyperlinks.Add` 不会覆盖已有单元格的值**：`TextToDisplay` 只在空单元格上写入。
  测试据此改在空单元格上验证 `url`/`text` 确实传到了桥；
### 16. 打开/另存/转换的路径键名错配，以及「转 PDF 永远导 Excel」（已修）

闸门覆盖所有的工具后（对账从 212 对提升到 226 对），又暴露出同一类错配，而且都在**文件路径**上：

| 工具 | 工具发的 | 桥读的 | 实际后果 |
| --- | --- | --- | --- |
| wps_common_save_as | filePath / outputPath | path | **另存为忽略目标路径** |
| wps_excel_open_workbook | filePath | path | 打开的不是指定文件 |
| wps_word_open_document | filePath | path | 同上 |
| wps_word_insert_image | imagePath / filePath | path | 图片路径丢失 |
| wps_word_set_page_setup | marginTop 等 | topMargin 等 | 页边距全部无效 |

上游的做法是**一次发多个别名赌其中一个会被读到**（例如同时发 `filePath`/`path`/`outputPath`）。
闸门现在会把没被读取的别名直接报错，这类赌博写法无法再蒙混过关。

另外两个能力缺口：

- `convertToPDF` / `convertFormat` 按 **Excel → Word → PPT 取第一个正在运行的应用**，
  而 Excel 几乎总是开着，所以「把这个 Word 转成 PDF」实际导出的是工作簿。
  两个工具现在都有 `app_type`（excel/word/ppt）参数；指定后只考虑该应用，
  指定了却没有对应文档时**明确报错**而不是退回去导 Excel；
- `openAfterExport`（导出后打开）此前被丢弃，现已实现，并在结果里回报 `opened` / `failed: 原因`；
- `wps_word_insert_header` / `insert_footer` 的 `section` 此前被丢弃，
  现在按节写入，节号不存在时报出「文档只有 N 节」。

回归：test/file-ops.test.mjs **11 项**——另存为真的落盘、打开工作簿真的打开、
文件不存在要报错、`app_type=ppt` 时不得偷偷导出 Excel、页眉页脚按节且越界报错。

### 17. 测试自身的文档泄漏（已修）

上面两个测试暴露了一个卫生问题：测试每次 `createWorkbook`/`createDocument` 都不收尾，
累计泄漏了 **23 个工作簿 + 14 个 Word 文档**。

根因是早期 `close_*` 会弹模态框（第 13 条之前的 `save`/`saveChanges` 错配），
所以测试一律不敢关文档。关闭工具修好之后，测试可以正常收尾了：

- 会新建文档的 5 个测试都加了收尾（Word 没有 close 工具，统一走 `wps_call` 门面）；
- 实测跑完全量后 **leftover workbooks=0 / word docs=0 / presentations=0**。
### 18. PPT 参数契约批次：55 处归零，并新增两张生成表（已修）

第 13 条的闸门把 PPT 侧一次性列全 55 处。其中大量是**同一个形状的键名不同**，
逐个改 case 会写出 20 份重复逻辑，于是改为两张由生成器产出的表：

- **别名表**（ActionParamAliases，20 条）：公开参数名 → 该 action 真正读取的键，
  在参数校验之前把值搬到规范键上。别名同时并入「接受的键」集合，因此不可能绕过校验；
- **容器表**（ActionNestedParams，13 条）：工具发的是嵌套对象（style/shadow/border/
  gradient/element/rotation），action 读的是平铺键。展开后再做键校验——
  嵌套里多写一个 action 不读的属性，同样会明确报错。

实际修好的能力：

- **切换效果与动画整族此前完全不可用**：EntryEffect/AddEffect 要的是 PpEntryEffect /
  MsoAnimEffect 数值，而工具传的是 fade/fadeIn 之类的名字，赋值直接抛「索引超出了数组界限」。
  现在有两张名称→数值表，未知名明确报错，并支持直接传数值；
- 动画触发器 trigger、退出类动画（fadeOut/flyOut）、按 shapeIndex 只给一个形状加动画
  （此前 addAnimationPreset 给整页每个形状都加）；
- 仪表盘 max、进度条 value、页码指示器 position、母版/幻灯片背景对象
  （solid/gradient/image，缺色或未知类型明确报错）；
- 3D 旋转：ThreeD.RotationX 是 Single，传 Int32 报「指定的转换无效」，现显式转换；
- 页脚 show、日期时间 autoUpdate/format、配色 slideIndex、统一字体 includeTitle/includeBody、
  复制幻灯片 targetIndex、标题装饰 style 的 5 种样式；
- 图片/文本框的**类别序号**：imageIndex 数图片、textboxIndex 数文本框，与 shapeIndex
  （任意形状）语义不同，因此按类别解析而不是当别名，否则会删错对象；
- wps_ppt_insert_slide_image 原本指向 **Word 的 insertImage**，插图片插进了 Word 文档，改指 insertPptImage；
- 迷你图表 data（sparkline 序列→末值+趋势）、流程图 nodes/connections、组织架构 data（树→层级）
  此前都是「发了不读」。

### 本轮实测到的 WPS 限制（据此删掉参数，而不是假装支持）

- **Adjustments.Item(1) = x 会挂住常驻宿主**：PowerShell 无法给带参数的属性赋值，
  而 WPS 的扇形只暴露 1 个 adjustment，多段环形图因此无法实现；
- **渐变角度 GradientAngle 同样挂起宿主**；
- **组织架构图传入自定义节点会 60s 超时**；
- **图表数据无法安全注入**：需要打开图表内嵌工作簿，常驻宿主可能留下隐藏文档。

对应参数从 schema 移除并写明原因（环形图多段 data、组织架构 data、图表 data、
渐变 angle/type、页码 startFrom），而不是保留一个静默失效的入口。

### 新增的两道门禁

- 生成器写文件前用 Parser::ParseInput 真正解析生成的模块，任一语法错误即拒绝写出
  （此前 Tokenize 放过了坏文件，导致全部 action 一起失效）；
- verify.mjs 断言桥的 action 数量（259）在源文件与生成产物之间一致——
  补丁脚本曾静默吞掉一个 case（slide.unifyFont），当时没有任何检查发现。

### 状态

参数契约 sweep：**A=0、B=0、C=0**（212 对工具/action 全部一致），验证门禁 23 项通过。

PPT 端到端测试 test/ppt-contract-fixes.test.mjs **42 项全部通过**（真实 WPS，写-改-回读）。

排障过程中定位到一个此前一直存在的**根因缺陷**：获取应用对象的兜底链里
`[Activator]::CreateInstance` 会返回"空壳"实例——`Presentations` 看似非空，
但 `Add()` 内部解引用未初始化的文档管理器，报的是
"cannot call a method on a null-valued expression"，与真正的原因相距极远。
现已**移除该兜底**，改为只认 `GetActiveObject` 与 `New-Object`，并加可用性校验与进程内缓存；
`createPresentation` 在失败时会弃用缓存实例重试一次。

同一批还修掉：我新加的 slideIndex 上界校验依赖 `Slides.Count`，
而 WPS 对 COM 新建的文稿**恒报 0**，导致所有合法页码被拒——改为"下界必查、上界仅在计数可信时查"。

### 19. 技能文档同步（并因此发现 4 个缺陷）

四份 SKILL.md 按本轮的真实行为重写，reference.md 重新生成（工具数：表格 82、文字 32、演示 115、
门面与通用 25；standard 档直接广告 43）。主要更正：

- **参数校验的准确表述**：桥校验的是「工具层发给桥的键」，不是「调用方传给工具的键」。
  必填参数写错名会被 schema 挡下（Missing required parameter: X），
  **可选参数写错名仍然是静默忽略**——这一点此前含糊，现在明确写成两条。
  旧文档里「写错名字不会报错」的说法已删除；
- standard 档工具数 31 → 43；
- 表格：sheet 参数已在几乎所有工具上生效（删掉「底层曾取 0 号工作表」的旧警告）、
  工作表 position 是 0 基、delete_sheet 必须显式给名字、常用工具的正确参数名列表；
- 文字：set_line_spacing / set_page_setup / get_paragraphs 均已可用（旧文档说前两者是坏功能），
  find_replace 的「不传 replace_text 就是纯查找」、页眉页脚的 section、页边距单位是整数磅；
- 演示：动画与切换的名称/数值、trigger、imageIndex 与 shapeIndex 的语义差别、
  以及**不支持的能力**（多段环形图、自定义组织架构节点、图表数据注入、渐变角度、页码起始编号）；
  同时写明「新建文稿 Slides.Count 恒报 0，不要据此判断文稿为空」。

写文档需要核实说法，于是逐个实测，又抓到并修掉 4 个缺陷：

1. **段落读取一直失败**：三处 `TrimEnd("\r\n", "\r", "\n")` 是无效写法——
   PowerShell 的 TrimEnd 只接受 char[]，多字符字符串转换直接抛类型错误。
   改为 `TrimEnd([char[]]@([char]13, [char]10))`；
2. **页面设置的四个边距全部 E_FAIL**：action 把入参当厘米乘 28.35（而 schema 写的是磅），
   且 WPS 的边距**只接受整数磅**（小数报「指定的转换无效」，超范围报裸 E_FAIL）。
   现在按磅取整、校验 0-1584 并给出可读的范围错误，且**回读实际生效值**回报给调用方；
3. **98 处演示 action 在没有打开文稿时抛裸空引用**：现在统一返回 no presentation is open；
4. **出错时返回不合法的 MCP 结果**：`JSON.stringify(undefined)` 让整个 tools/call 结果非法，
   客户端看到的是 -32602 协议错误而不是错误消息。改为合法的内容块。

### 20. 「11 个实参不可知的 handler」：其实只有一个不设防（已修）

之前把这件事记成「11 个 handler 的实参静态不可知」，复核后发现这个说法既高估了风险、
又低估了覆盖：

- **9 个是「整包转发」**（handler 把调用方对象原样交给桥，例如 `executeMethod('evaluateFormula', args)`）。
  对这类 handler，**schema 就是桥实际收到的东西**——所以「schema 的键 ⊆ 桥接受的键」
  是一条可静态执行的检查，正好补上解析不到的缺口。这就是新的 **D 类**；
- **2 个是客户端 helper 路径**（read_range / write_range 走 `wpsClient.getRangeData/setRangeData`）。
  它们的键固定在 wps-client.ts 里，于是 sweep 改为从那里读取发送的键，
  这两个工具因此回到常规校验（对账从 226 对升到 228 对）；
- 剩下 1 个 `wps_word_proofread_basic` 是**纯本地 JS**，完全不碰 COM，没什么可校验的。

D 类一上线就抓到一个被缺口掩盖的真实错配：

**`wps_excel_evaluate_formula` 的 `cell`**：schema 宣称「目标单元格（可选）」，
但 action 用 `Application.Evaluate` 求值，**从不读取 cell**。已移除该参数，
并在描述里写明「由 Excel 求值、不写入任何单元格；要在单元格里求值请用 set_formula」。

顺带说明一处仍然存在的设计性例外：`wps_excel_set_cell_format` 的 action 会遍历嵌套的 `format` 对象，
参数名无法静态推导，因此**它不进键表、也不受参数闸门保护**——
这一处靠 test/cell-format.test.mjs 的 16 项回读验证兜底，属于已知例外而不是遗漏。

现状：A/B/C/D 四类全部为 0；未解析 9 个中 8 个由 D 类覆盖，真正无覆盖的只剩
proofread_basic（无 COM）与 set_cell_format（动态 action，另有专门测试）。

### 21. 其余 15 对重复工具合并（已修）

判据用的是「多个工具驱动同一个 COM action」——比按名字相似去猜可靠。全量扫出 18 组，
其中 3 组此前已合并，剩下 15 组这次做完（合并表从 3 条增到 18 条）：

| 旧名（转发） | 规范名 |
| --- | --- |
| insert_slide_image / insert_image | insert_ppt_image |
| auto_fill | fill_series |
| insert_row / hide_row / delete_row / insert_column / delete_column | 各自的复数规范名 |
| set_animation | add_animation |
| set_transition | set_slide_transition |
| set_background | set_slide_background |
| add_chart | insert_ppt_chart |
| duplicate_slide | copy_slide |
| align_objects | align_shapes |
| set_font_style | set_font |

做法：保留**能力更全或已在广告**的那个名字，旧名保留可调用（隐藏于 wps_help），用 paramMap 改名转发。
为了让改名不丢能力，**5 个规范工具被扩成参数并集**：fill_series 接受 sourceRange/targetRange/startValue、
hide_rows 接受 row/rows/count/hide、delete_rows 接受 row、set_slide_background 接受平铺的 color/imagePath、
align_shapes 的 shapeIndices 变可选。

合并过程中又修掉 5 个缺陷：

1. **AddChart 要的是 XlChartType 数值**，而工具传的是名字，直接抛「无法将类型 string 转换为 Object」。
   现在有名称→数值映射，未知名明确报错并列出可用名；
2. **平铺写法的背景不生效**：Get-PptBackgroundSpec 在没有 type 时一律按 solid 处理，
   只给 imagePath 会报「需要 color」。现在按提供的值推断类型（image/gradient/solid）；
3. **ShapeRange.Align 在这里不可用**：常驻宿主里 PowerShell 解析不到该成员，反射调用又会拿到 Object[]
   （Shapes.Range() 被展开）。改为按几何直接对齐，语义一致且确定；
4. **show_rows 不该转发 hide**：它的 action 从不读取该键，转发只会被参数闸门拒绝；
5. 规范工具 align_shapes 的成功消息解引用了可选的 shapeIndices，旧名转发时抛 undefined 错误。

特意**没有**合并的一对：wps_ppt_set_active_target 与 wps_ppt_get_open_presentations 共用 action，
但前者是「锁定目标文稿」这一独立能力（后者只是它的校验手段），属同 action 不同功能，不是重复。

回归：test/merged-tools.test.mjs **31 项**——每个旧名用**自己的**参数名调用，并核实效果
（插入行/列、隐藏后再显示、自动填充扩展、三张图片落位、动画/切换/背景/图表、无索引对齐、
复制幻灯片、字体转发）。全量 13 个测试文件 270 项 + verify 23 项通过，跑完无残留文档。

### 22. 高层场景封装从工具层下沉到技能层（已做）

上一节判据用的是同一个 `action`，这一节的判据是**同一个目的**：这些工具的唯一用途是
「照固定套路画一张信息图」，而不是提供原子能力。共 19 个：

- 信息图：KPI 卡片、时间线、流程图、组织架构图、迷你图表、网格布局；
- 图表件：环形图、仪表盘、进度条；
- 装饰与版式：标题装饰、页码指示器、自动排版、智能分布；
- 美化与主题：配色方案、自动美化（单页/全部）、带样式表格、3D 文字、演示主题。

做法：

1. 从工具注册表摘除这 19 个工具（工具目录 254 → 235）。
   **广告面不变（仍是 43）**——它们本来就都不在 standard 档里；
2. 在 `skills/wps-ppt/SKILL.md` 新增「组合配方」一节，用**基础工具**给出每张图的画法：
   坐标怎么算、用哪个形状、什么时候用文本框、配色怎么统一。配方是给模型看的，可按场景调整；
3. 配方里写明了已知限制，避免重新踩坑：多段占比**不要用扇形**
   （WPS 的扇形只暴露一个 adjustment，改角度会挂住 COM 宿主），改用堆叠条 + 图例。

关于「能力是否丢失」：桥里对应的 action **保留**，所以仍然可以用
`wps_call {tool:"wps_execute_method", args:{method:"createKpiCards", ...}}` 调用，
参数见同节表格；变化的是工具层不再替它们背书。

证据：`test/ppt-contract-fixes.test.mjs` 里原本依赖这些封装的 9 项断言，
**改写成用基础工具实现同样图形的配方断言**（进度条填充宽度是否为轨道的 60%、外圈+内圈+中心文本、
方块+箭头、节点+轴+标签、卡片形状与文字……），测试项从 42 增到 **57**，全部在真实 WPS 上通过。
这同时证明「去掉封装、让模型按配方组合」是可行的，而不是把功能删掉了事。

现状：工具目录 235（广告 43），桥 action 259 不变，契约对账 209 对，A/B/C/D 四类均为 0。

**源码同步清理（已完成）**：上一步只摘了注册表、代码仍留在源文件里，于是又做了一次删除：

- 删掉 38 个块（19 个工具的定义 + handler），共 **1711 行**；
- `data-viz.ts` 删空后整文件移除（其余四个文件仍有别的工具，保留）；
- 清掉文件头与 `ppt/index.ts` 注释里对已删工具的清单（本项目约定：改了文件就要更新头部注释）；
- `noUnusedLocals` 会拦住遗留的未用导入，构建因此充当了这次清理的校验器。

结果：源码与工具面一致——工具目录 235（广告 43），桥 action 259 不变（仍可用 wps_call 调用），
契约对账 209 对，A/B/C/D 均为 0，13 个测试文件 285 项 + verify 23 项通过，跑完无残留文档。

### 23. 消灭静默 catch：49 处改为「完成操作 + 如实报告」（已修）

桥里原有 **60 处 `catch { }`**。它们不一定都危险，但共同点是：出错时调用方**什么也看不到**，
于是「改了但没生效」和「改好了」在结果上无法区分。

做法不是一律改成报错——有些确实属于「尽力而为」（例如某个属性在当前 WPS 版本上不存在，
失败不应中断整个操作）。取中间路线：**操作继续完成，但把失败收集进结果**。

- 桥新增 `$script:WpsWarnings` / `Add-WpsWarning` / `Clear-WpsWarnings`；
- 生成器在每个 action 开始前清空，`Output-Json` 在结果里附加 `warnings`，
  并**同时写进 `data`**，这样只回传 `data` 的透传型工具也能看到；
- 49 处转为 `catch { Add-WpsWarning $_.Exception.Message }`；
- 保留静默的 11 处，都是**失败属预期**的：COM 实例探测（GetActiveObject / New-Object 尝试）、
  `Visible` 设置、以及 `DisplayAlerts` 的读取与还原（还原动作本身若失败不应掩盖真正的错误）。

验证：`test/warnings.test.mjs` **7 项**——用一个不存在的样式名制造真实的尽力而为失败，
断言「操作仍然成功」+「`warnings` 非空且带错误原文」+「`data.warnings` 同样存在」+
「下一次干净调用不带任何警告」（证明是逐调用清空）。

一个顺带的发现：验证触发点时才确认，透视表/图表的两处查找**在 WPS 上是静默返回 null 而非抛错**，
所以它们本来就不会产生警告——「没有警告」也可能意味着「根本没有失败」，两者不能混为一谈。

技能同步：`skills/wps-office-next/SKILL.md` 的失败处理一节增加一条——结果里可能出现 `warnings`，
涉及本次操作时要如实转述给用户，不要因为它 `success` 为真就忽略。

### 24. 第一次真实端到端验证：一次任务暴露 5 个缺陷（已修）

前 23 条都是读代码或单点测试找出来的。为了知道**模型真的用起来是什么样**，跑了一次真实端到端：
新建 headless profile（`dsh --profile wpse2e --from-default-profile headless`）→
`dsh plugin --profile wpse2e add D:\dsh\a` → 给一个跨应用任务，事后把会话日志逐帧解码，
复盘每一次工具调用（`$DSH_HOME/sessions/--D-dsh-a--/session-*.jsonl.zstd` 是**追加式拼接的
多个 zstd 帧**，Node 只解第一帧，必须按帧边界切分才能拿到全部事件）。

**任务**：打开 `sales.xlsx`（12 行明细）→ 按「地区」汇总写入新工作表「汇总」并降序 →
画簇状柱形图 → 新建 Word 文档写结论另存为 `结论.docx` → 保存并关闭这两个文件。

**结果**：72 秒，exit 0。两个产物**另起 COM 会话独立复核**：「汇总」A1:B5 = 华北 330 / 华南 280 /
华东 240 / 西南 165；图表 ChartType=51（簇状柱形）、带标题、无图例、1 个系列；`结论.docx` 3 段
（标题 1 + 正文）145 字符；收尾后已打开工作簿 0、文档 0。

**模型侧表现**（这部分是本轮最想知道的）：3 个技能按需触发（router → excel → word）；
用 `wps_help {tool:...}` 取隐藏工具 schema；用 `wps_batch` 合并调用；用
`export_chart_as_image` + `read_image` **自己导出图片看图表画对没有**；全程没自写 COM 脚本；
最终报告的数字与实测一致。19 次工具调用，其中 7 次 `wps_help`、6 次 `wps_call`。

它自报的 4 处异常里，有 4 条都指向真问题：`wps_help` 分词不中、Word 没有「新建文档」直达工具、
`get_open_workbooks` 打印 `[object Object]`、`set_formula` 回读 `null`。下面 5 条即这些问题的修复。

### 25. wps_help 自由文本检索：整串子串 → 分词 + CJK 二元组（已修）

`wps_help` 的 `query` 把整条查询当成一个子串去 `name/description.includes(query)`，于是模型最自然的
两种问法都返回 0 命中：`"新建 文档 create new"`、`"close workbook 关闭"`。它只能退回 grep 技能文档找工具名，
多花了 4 次调用才够到目标。

现在按 token 打分（名称命中 4 分 / 描述命中 2 分 / 整串命中 10–12 分），并对中文再走一遍**二元组覆盖率**
（≥50% 才计分）：中文没有空格，`"关闭工作簿"` 并不是描述 `"关闭指定的Excel工作簿，可选是否保存"` 的子串，
但 4 个二元组命中 3 个。命中 0 时返回 `hint` 指路，而不是把「没搜到」当成「不存在」。

验证：`test/word-lifecycle.test.mjs` 前 5 项（中英混排、无空格中文、精确名排第一、0 命中带 hint）。

### 26. get_open_workbooks 输出 `[object Object]`（已修）

桥返回的是对象数组（`{name,path,sheets,active}`），工具层却声明成 `string[]` 直接 `join('\n')`，
于是**每一个已打开的工作簿都显示成 `[object Object]`**，这个工具实际上无法用。
现在逐项格式化成 `名称 | 完整路径 | N 个工作表 | 当前活动`；未落盘的工作簿 FullName 等于名字，只显示一次。

### 27. get_active_document 输出 `undefined`，并把文档名当路径（已修）

固定模板 `页数: ${d.pageCount}` 必然打印 `页数: undefined`（桥从来不返回 pageCount），
未保存的文档 `FullName` 其实是它的名字 `文字文稿1`，却被打上「路径」标签。
现在只打印桥真正返回的字段，并补上真实 `pageCount`（`ComputeStatistics(2)`，失败则记 warning 而不是编一个数），
未落盘文档显示 `路径: (尚未保存到磁盘)`。

### 28. set_formula：回读 `null` + 多格广播在常驻宿主里直接失败（已修）

两件事：

1. 桥只回 `{success:true}`，工具层却写 `计算结果: ${JSON.stringify(data ?? null)}`——**每一次**公式写入
   都显示「计算结果: null」，读起来像算错了。现在单格回读真实值，多格明确标注「（区域首格）」。
2. 顺手加的回归测试立刻抓出更硬的问题：`Range("D1:D2").Formula = x` 在**常驻宿主**里直接失败
   （`在此对象上找不到属性"Formula"`），而 Excel 会把它广播到整个区域。实测三种取格方式在宿主里都不可靠
   （一参 `Range.Cells.Item(i)` 报「索引超出了数组界限」，`Range.Row/Column/Rows.Count` 报「值不在预期的范围内」），
   最终改为**自己解析 A1 地址**（新增 `Get-AddressSpan`）+ 工作表双参 `Cells.Item(r,c)` 逐格写入。

值得记下的一条经验：同一个 COM 调用在进程内 `New-Object` 与常驻宿主里的行为**不一样**——
探针脚本跑通不等于宿主里跑通。三处失败里有两处只在宿主里出现。

验证：`word-lifecycle.test.mjs` 的 2 项公式检查（单格 10、多格「区域首格: 10」，且结果不含 `null`）。

### 29. Word 缺少「新建/关闭文档」工具（已修）

Excel 有 `create_workbook`/`close_workbook`、PPT 有 `create_presentation`，**Word 连隐藏的建/关工具都没有**：
桥里的 `createDocument`/`closeDocument` 两个 action 只能靠 `wps_execute_method` 拼 COM 方法名去够。
e2e 里模型正是这么绕过去的。

现在补上 `wps_word_create_document` 与 `wps_word_close_document`；前者**加入 standard 广告位**
（Word 是第 2 优先级，而它是唯一连隐藏工具都查不到的路径），后者保持隐藏、由修好的 `wps_help` 检索得到。
桥侧 `createDocument` 同时改为回报新文档名、失败时明确报错，不再只回一个 `{success:true}`。

预算：广告面 43 → **44 工具 / 23,593 字节**（上限 45 / 25,000），注册目录 235 → 237，桥 action 仍 259。

### 30. e2e 一键化：一条命令跑完并断言（已落地）

第 24 条的两轮验证是手工做的：造 fixture、跑任务、写脚本解会话日志、再用眼睛看结果。手工步骤永远不会失败，
也就拦不住回归。现在固化成 `scripts/e2e.mjs`：

    node scripts/e2e.mjs --profile <name>            # 19 项检查，约 65 秒
    node scripts/e2e.mjs --profile <name> --setup    # profile 不存在就先建好并装上本包，再跑

它做五件事，每一步的证据都留在 `test/.artifacts/e2e/<run>/`：

1. **自己造 fixture**（裸 COM，刻意不走本插件——用具自己的 bug 去伪造输入，测出来的东西没意义）；
2. 跑一个真实 headless 任务（`DSH_PERMISSION_MODE=danger-full-access`，并显式清掉 `DSH_TOOLS_MODE`，
   保证走的是**广告工具面**而不是 PTC 的 `run_code`）；
3. 找这次运行的会话日志并**逐帧解码**后打印轨迹；
4. 用裸 COM 重开产物核对：工作表、降序合计、簇状柱形图（ChartType 51）、文档里的各地区合计；
5. 行为断言：收尾后 0 个残留文档；**操作类结果**里没有 7 个已知缺陷标记；模型没有在自己写的 pwsh 里碰 COM
   （`New-Object -ComObject` / `Ket.Application` 等）；用了 `wps_status`、≥8 次 wps 工具
   （含 `wps_batch` / `wps_call` 的批内成员）、≥2 个技能，且 Word 文档是插件工具建的。

**跑通的过程中它也抓出三处自己的 bug**，照实记下——这三个坑对任何写 e2e 的人都会遇到：

- 会话目录**少看了一层**（真实结构是 `sessions/<encoded-cwd>/<session-id>/session.v3.jsonl.zstd`），
  于是「找不到会话」让 4 个行为断言全部 FAIL。**这正是它该有的样子**：找不到证据就失败，而不是默认通过；
- 缺陷标记扫到了**技能文档正文**（技能里本来就在教「遇到 `unknown parameter(` 怎么办」），于是误报；
  改为只扫 `wps_*` 的操作结果、并排除 `wps_help`（它返回的就是文档）；
- 只在**顶层**工具名里找 `wps_word_create_document`，而它是在 `wps_batch` 里被调用的，于是误判「没用插件建文档」；
  改为展开 `wps_batch.calls[].tool` 与 `wps_call.tool`。
- `--setup` 用 `join()` 拼出 `@deepseek-ai\dsh-base` 去匹配 dump-config 输出，而那份输出用的是正斜杠，
  于是明明建好了却报失败；改为直接校验 profile 的 `package.json`。

另外做了**负向验证**：把 `--timeout` 压到 5 秒，得到 **10 项 FAIL、exit 1**——确认这个门禁会真的失败，
而不只是「打印一堆 PASS」。

实测：`E2E OK (19 checks) in 65s`，exit 0；`--setup` 全流程（建 profile → 装包 → 跑验收）
也实测通过：**46 秒、19 项全绿**。

### 31. CI：只跑不需要 WPS 的那部分（已落地）

`.github/workflows/ci.yml`（GitHub Actions，`windows-latest`）跑六件事，全部不需要 WPS：
tsc 构建 + `mcp/dist` 漂移检查、宿主生成器 + `host/wps-actions.ps1` 漂移检查、技能表重生成 +
`skills/**` 漂移检查、`verify.mjs --static`（18 项）、参数契约对账 + 报告漂移检查、
以及两个静态测试文件（`plugin.test.mjs` 32 项、`com-host.test.mjs` 6 项）。

两处取舍值得记下来：

- **用漂移检查而不是重新生成**：`mcp/dist`、`host/wps-actions.ps1`、`skills/**/reference.md` 都是入库产物。
  CI 重新生成它们并断言工作区没有变化——「改了源码忘了重建」会立刻失败，而不是等到用户装上来才发现。
  这能成立的前提是产物可复现：`mcp/package-lock.json` 把 typescript 锁在 5.9.3，dist 的 sourcemap 里也没有绝对路径。
- **`verify.mjs --static`**：原来 23 项里有 5 项要真实 WPS（`wps_status`、真实派发、`wps_batch`），
  在没装 WPS 的 runner 上必然失败。现在把这 5 项归入完整模式，静态模式跑 18 项——
  **跑多少项由模式决定**，而不是「反正 CI 上会红就先注释掉」。

实测：workflow 的每一条命令都在本机按同样顺序干跑过一遍，**11 步 0 失败**、三处漂移检查全干净；
`npm ci --dry-run` 通过（锁文件与 package.json 一致），workflow 的 YAML 也过了解析器。

### 32. P0 清理：删死代码、拆废弃名、恢复主题工具（已落地）

按 docs/tool-roadmap.md 的 P0 执行，每一步都可复跑。

**P0-1 删 11 个 builtin**（保留 `wps_execute_method` 逃生舱）：6 个与 pro 工具重复、`wps_check_connection`、
以及 4 个没人用的缓存工具。顺带删掉 `ToolRegistry.dataCache`——**tsc 的 noUnusedLocals 当场抓到了这个残留**，
构建再次充当清理的校验器。

**P0-2 18 个废弃名从「注册工具」改成「派发别名」**：以前它们是注销后再注册的转发器，每个仍占一个注册位
与一份重复 schema；现在注册表里根本没有它们。`tools/call`、`wps_call`、`wps_batch` 三处派发前统一查
`DEPRECATED_TOOLS`，解析成规范工具并改名参数；`wps_help {tool:"旧名"}` 仍回答 `deprecated: true` 与 canonical。
旧名字照样能调（实测 `wps_excel_zoom` 转发后返回 130%）。

**P0-3 删桥里的死代码**：28 个 case 条目 / 1039 行——22 个场景封装 action（KPI 卡/仪表盘/时间线/流程图/
3D 文字/网格/自动美化等，工具早已删除）、3 个点号重复实现（`slide.add`/`slide.unifyFont`/`slide.beautify`）、
以及 4 个重复 case 标签的**后一份**（PowerShell switch 只认第一份，后一份永不可达）。
动刀前先 dry-run 打印每块的起止行与首末行，确认无重叠、且每个名字在桥里只出现在自己的标签处；
删完由生成器验证解析与数量（switch_cases 259 → 231）。

**P0-4 恢复 `wps_ppt_set_slide_theme`**：FIXES 第 1 条把它列为已实现，第 22 条的场景清理又把它当「美化」
删掉了——文档与现实打架。D3 决定保留「主题」，所以恢复工具，并把 schema 写准：`theme` 是**模板文件路径**
（.thmx/.potx/.pptx，必须存在），不再沿用上游那句误导的「主题名称」。

**踩到并修掉的一个坑**：我按「名字出现在重复 case 标签里」就顺手删了 `set3DRotation` 的**容器表**条目，
但那个 action 是活的（`wps_ppt_set_3d_rotation` 发的是嵌套 `rotation` 对象）——表项一删，
`ppt-contract-fixes` 的 3D 旋转用例立刻变红，参数契约 A 类也从 0 变 1。教训：
**重复标签只说明有一份是死的，不说明整个 action 是死的**；删表项要看「还有没有工具在驱动这个 action」，
而不是看名字。

**回归**：310 项测试 + verify 23 项 + 一键 e2e + CI 全绿；参数契约 197 对、A/B/C/D 均为 0（UNPARSED 9 → 6，
少掉的正是已删的废弃 handler）；注册 209（广告 44 / 隐藏 165 / 别名 18）。

**顺带**：生成器别名表 20 → 17、容器表 13 → 12；`skills/*/reference.md` 全部重生成；
`mcp/src/tools/index.ts` 的头部枚举（长期过期、写的是 235 个工具与 12 个内置）改成指针——
清单本来就该由生成物承担。

### 33. P1 第一步：操作规格（spec）成为工具面的真源（已落地）

把「工具契约」从三处人肉维护（TS 定义 + 生成器两张表 + 精选数组）收敛成一份声明式 spec：
`mcp/src/spec/types.ts`（形状）、`mcp/src/spec/operations.ts`（209 条，由 `scripts/extract-spec.mjs`
从今天的事实 bootstrap）、`scripts/gen-tool-surface.mjs`（spec → 四份产物）。

**验收标准是「逐字节复现今天的工具面」**，`test/spec-reproduction.test.mjs` 直接跑这件事：

| 检查 | 结果 |
|---|---|
| 工具数 / 工具名 | 209 / 完全一致 |
| 209 个 schema 深度相等（键排序后） | 全部一致 |
| 序列化字节数 | **121,308 = 121,308** |
| 广告集（44 个） | 完全一致 |
| spec 的 action 都在宿主键表里 | 202 个 action，1 个已知例外（setCellFormat 动态键） |

过程中被数据纠正了两次，都记下来：

- 老 surface **本身并不一致**：35 个工具写了 `required: []`，14 个干脆没有这个键。所以 spec 必须
  **原样保留** `required`（undefined = 原 schema 没有这个键），否则字节数对不上（差 196 = 14 × 14）。
  这也说明「迁移先复现、再谈收拾」的顺序是对的：不一致本身是先复现出来才看得见的。
- 我原以为「工具参数名 = 桥键（经生成器别名表）」，实测发现真正的重命名发生在 **handler 代码**里
  （`createChart` 收蛇形 `data_range`、桥读驼峰 `dataRange`）。这条记成账：
  **58 处「参数名 ≠ 桥键」分布在 33 个 action 上**，正是 P1-4 要归零的第二套方言。

**账本式门禁**：`rename debt ≤ 57`、`untooled actions ≤ 29` 两条断言把债务写进测试——
**涨了会红，降了要手动改数字**。29 个「桥里有、没工具」的 action 同时就是 P2 的待办清单
（autoFit*、命名范围、条件格式、数据验证、getComments、getBookmarks、getDocumentStats…）。

**新增产物**（`spec/`，入库、进 CI 漂移检查）：`tool-definitions.json`（模型看到的 surface）、
`action-keys.json`（每个 action 接受的键，将来供宿主生成器读取，取代「用正则扫桥源码」）、
`advertised.json`（广告集）、`signatures.json`（紧凑签名索引 `range*, sheet?`——给 `wps_help` 用，
能把一次 schema 查询省掉）。

**当前债务（下一步的目标）**：raw schema 片段 32/549、带别名的工具 15、带容器的工具 12、
重命名 58 处、未工具化 action 29 个。

### 34. P1-3：共享的调用点解析器 —— spec 说出「工具真正发什么键」（已落地）

P1 第一步量出的 58 处「参数名 ≠ 桥键」里，有一大半不是真差异，而是**我的推导方式不对**：
工具的参数名要经过 handler 代码才变成桥键（`data_range` → `dataRange`），而我只做了「参数名 → 别名表」这一跳。

这一步把 `scripts/param-contract.mjs` 里那个成熟的解析器抽成 `scripts/lib/tool-action-map.mjs`
（工具源码 → 它驱动的 action + 它实际发送的键，含 `sentKeys` 的展开与「读不出」分类），
`param-contract.mjs` 改为 import 它——**抽完立刻验证：参数契约输出与 `docs/param-contract.md` 逐字节未变**。

提取器随后用「发送键」反推桥键（形状归一：去下划线 + 小写，覆盖蛇形/驼峰），于是：

| 指标 | 之前 | 现在 |
|---|---|---|
| 重命名债务（参数名 ≠ 桥键） | 57 | **27** |
| 工具面复现（schema / 字节数 / 广告集） | 通过 | 仍然通过（209 / 121,308 / 44） |
| action 键表与宿主表不一致的键 | 57 | 27（就是上面那 27 个同义改名） |

剩下的 27 处是**同义改名**（`filePath` → `path`、`style_name` → `style`、`font_name` → `font`…）：
公开名与桥键确实不同，只能靠 P1-4 统一（改 schema 参数名，需要时用已有的废弃别名机制兼容旧名）。
测试额度随手收紧：`RENAME_DEBT` 57 → **27**（棘轮：只能降不能升）。

方法上的收获：**先把解析器抽成共享模块，再改推导方向**，比在两个脚本里各写一份强——
`param-contract` 的地基是「handler 真正发了什么」，spec 的 `action-keys.json` 也该建在同一块地基上，
否则两边会慢慢漂开（这正是 P0 之前三张人手表的老毛病）。

### 35. P1-4a：给每个参数一个明确的去向（已落地）

接着 P1-3 剩下的 27 处「参数名 ≠ 桥键」逐个查清后发现：**它们不是一类东西，而是三种去向**，
而 spec 之前只认识一种。

| 去向 | 例子 | 数量 |
|---|---|---|
| **bridge**：发给桥（公开名可能与桥键不同） | `chart_type` → 桥读 `chartType` | 509 |
| **local**：handler 自己消费，从不到桥 | `read_range` 的 `include_header`、`insert_text` 的 `new_paragraph`、`generate_formula` 的 `description`/`target_cell`、`set_active_target` 的 `name`/`clear`、`get_document_text` 的 `start`/`end` | 22 |
| **container / 折进单个参数**：被 handler 打包后只以一个键到达桥 | `wps_ppt_beautify` 把 `color_scheme`/`font`/`beautify_all` 折进它的 `style` 参数 | 3 |

`types.ts` 增加 `ParamSpec.kind`；提取器按「发送键形状匹配 → 宿主键表回退 → 显式声明」三级判定，
**未归类必须为 0**（这正是过去「静默忽略」的老毛病在 spec 层的对应物）。显式声明了
`LOCAL_PARAMS`（8 条）与 `FOLDED_INTO_ONE_ARG`（3 条）——**显式胜过猜测**：宁可留一张会被 review 的表，
也不要让提取器把「没人读的参数」和「handler 自己用的参数」混为一谈。

同时把验收收紧：

- 新增断言 **「每个 bridge 参数都落在桥真正读的键上」——严格 0 例外**（之前只是"债务 ≤ N"）；
- `alias debt`（公开名 ≠ 桥键，**62 处**）成为新的棘轮，目标 0；
- 未归类参数严格为 0。

过程中修掉两个我自己的错：把 `beautify` 的三个参数当成容器成员（其实 handler 把它们折进 `style` 一个参数，
桥只读 `style`）；以及生成器对 local / container 参数没有区别对待（会把 local 参数也写进 action 键表）。

工具面仍逐字节复现（209 schema / 121,308 字节 / 广告集 44），验收 **11 项全绿**。

### 36. P1-4b：两张手写表搬进 spec，宿主产物逐字节未变（已落地）

先试了"机械改名把 62 处别名对齐"，dry-run 立刻暴露两个**不能用改名处理**的情况：
`setBackgroundImage` 一个桥键对两个公开名（`imagePath`/`filePath`）、`addAnimation` 的
`shapeName`/`shapeIndex` 有语义纠缠（桥两者都读，含义不同）。**于是放弃机械改名，改做更根本的事**：
把手写表的数据搬进 spec，由生成器产出——行为不变，但真源变了。

- 新增 `mcp/src/spec/aliases.ts`：声明式地记录「公开参数名 → 桥读取的键」与「嵌套容器名」
  （由一次性脚本从生成器的两张表导入，含那些 schema 没声明、但历史上被接受的兼容拼写）。
- `gen-tool-surface.mjs` 产出 `spec/param-aliases.json` 与 `spec/param-containers.json`；
- `build-host-actions.ps1` 删掉两张手写表，改读这两个 JSON（PS 5.1 有 `ConvertFrom-Json`）。

**验收是「宿主产物逐字节未变」**：重新生成 `host/wps-actions.ps1` 后 `git diff` 为空，
生成器计数也一致（`param_aliases=17 containers=12`、`switch_cases=231`、`guard_installed=1`）。
这证明搬家没有改变任何运行时行为——比"测试还绿"更强的证据。

**新增一条不变量**（验收 12 项）：两种改名机制不许混淆——
handler 自己改名时桥侧不需要声明；handler **原样传公开名**时，桥侧必须有声明。
现在 16 处 pass-through 改名全部有声明，另有 2 条"没人发送的兼容拼写"被显式记账。

**三张人手表的状态**：aliases ✅ 归零（数据进 spec）、containers ✅ 归零、`$helperKeys` **仍在**
——它是"从桥源码推导键表"的辅助信息，只有把键表本身也改成 spec 产出才会消失（那是 P1-5 的伴生工作）。

### 37. P1-5：第三张手写表进 spec，动态 action 未声明就失败（已落地）

P1 收尾。`build-host-actions.ps1` 里的 `$helperKeys`（8 条：共享解析函数各读哪些键）是最后一张手写表。
它决定守卫接受哪些键，所以搬进 spec 时必须证明行为不变：

- 新增 `mcp/src/spec/bridge-helpers.ts`（一次性从手写表导入）→ 生成器产出 `spec/param-helpers.json`
  → 宿主生成器改读 JSON；
- **验收仍是「宿主产物逐字节未变」**：`git diff --numstat host/wps-actions.ps1` 为空，计数一致
  （`exit_remaining=0 switch_cases=231 functions=48 param_aliases=17 containers=12`）。

**同时给「读不出参数的 action」上了闸**：过去生成器遇到静态提取不出的 action 就静默跳过（今天 1 个：
`setCellFormat`），于是它**不受守卫保护**而没人知道。现在例外必须写进 `mcp/src/spec/aliases.ts` 的
`dynamicParamActions`（带原因），产出 `spec/param-dynamic.json`；**未声明的动态 action 会让生成器直接失败**。
负向验证：把声明清空后，生成器报
`dynamic actions missing from spec/param-dynamic.json: setCellFormat - declare each one (with a reason) in mcp/src/spec/aliases.ts`。
验收里那条硬编码的例外清单也改成从 spec 读取。

**一条值得记住的失败**：我原本想**从函数体自动推导**这 8 条键（那样连声明都不需要）。JS 原型显示
8 条里只有 6 条能精确复现——`Resolve-Worksheet` 的键藏在一个
`foreach ($key in @('sheet','name','oldName'))` 数组里；移植到 PowerShell 时我连犯三个错：
两次切错替换区域，一次在循环里复用了 `$body` 这个变量名，**把生成器正在拼装的 dispatch 主体覆盖掉了**。
最后是生成器自己的计数器 `exit_remaining=505` 把它抓住的（正常应为 0）。
教训：**当"更聪明的自动化"反复出错时，先用已验证过的机制把目标达成**（这里就是"数据进 spec"），
把自动推导留作有测试保护的后续改进。

**三张人手表状态**：aliases ✅、containers ✅、helperKeys ✅ —— 全部归零，每一次都以宿主产物逐字节未变作证。

### 38. P2 第一波：8 个 Excel「缺失的另一半」挂上工具，顺带挖出两个在常驻宿主里从不生效的 action（已落地）

P2 开始把 Excel 做深。第一波只挂**桥里已经实现、却没有工具出口**的能力，零新 COM 代码：

| 工具 | action | 说明 |
|---|---|---|
| `wps_excel_get_sheet_info` | getExcelContext | 工作簿/工作表名、已用范围、表头、当前单元格 |
| `wps_excel_auto_fit` / `_columns` / `_rows` | autoFitAll / autoFitColumn / autoFitRow | 三个显式字面量 action（静态提取器依赖这个写法） |
| `wps_excel_set_wrap_text` | wrapText | 打开/关闭自动换行 |
| `wps_excel_find_in_sheet` | findInSheet | 查找并报告命中地址，不改内容 |
| `wps_excel_get_named_ranges` / `_delete_named_range` | getNamedRanges / deleteNamedRange | 命名范围此前只有写 |

**但「这些 action 既然在桥里，挂出来就能用」这个假设是错的：两个 action 在常驻宿主里根本不可能生效**，
而且此前没有任何测试覆盖它们。验收 `test/excel-missing-halves.test.mjs`（18 项，真实 WPS）第一次跑就抓到：

1. **`Range.Find()` 结果的 `Address()` 在常驻宿主里永远失败**（0 参与 2 参都抛），
   而 `UsedRange.Address()`、`Selection.Address()` 都正常。旧实现先取 `$found.Address()` 再用它做
   `FindNext` 的比较基准，第一步就抛，于是 `find_in_sheet`/`find_replace` 的多格路径是死的。
   改为：范围来自调用方的 A1 串（缺省时用 `UsedRange.Address()`），一次 `Value2` 读回整块后自行扫描，
   与 `getRangeData` 用同一个访问器。
2. **PowerShell 的逗号优先级高于 `+`**。我按 C 系语言直觉写了 `$matrix[$mRow + $r, $mCol + $c]`，
   它被解析成 `$mRow + ($r, $mCol) + $c` → `int + Object[]` → 每条索引都抛
   `[System.Object[]] 不包含名为 op_Addition 的方法`。因为索引外面套着 `try { } catch { continue }`，
   异常被**逐格吞掉**，对外表现是「成功、0 命中」：单格范围（标量路径）查得到，多格范围永远查不到。
   修法是把下标列表括起来：`$matrix[($mRow + $r), ($mCol + $c)]`。

第 2 条是本轮最值得记住的一课：**裸 `catch { continue }` 会把「整段逻辑失效」伪装成「没有结果」**，
而空结果在这个仓库里太容易被当成正常。定位它的过程也值得记：先加诊断字段、经 error 通道把
rank/下界/首格/异常消息打出来，才从「循环体没执行」纠正到「循环体执行了、每格都抛」。

**验收**：18/18 绿，覆盖单格（C4）、显式范围（A1:C4）、不传范围用已用范围三条路径；
生成器计数不变（`switch_cases=231`）——本波没有增删 action，只改了实现。

**顺带清掉两笔契约债**：`find_in_sheet` 原本在 handler 里先调 `getExcelContext` 再调 `findInSheet`，
参数契约把它记成 D 类（透传声明了 action 读不到的键）；把范围解析移回桥后它是纯透传，D 归零。
另外 5 个新 handler 原来写 `executeMethod('x', args)`，静态读不出实参；改成显式键对象后，
被校验的参数对从 200 升到 205，`UNPARSED` 回到基线 6（都是既有工具）。

**D1 提前生效**：本波把广告面从 44 推到 48 工具 / 25,097 字节，越过旧的 45 / 25,000 门禁。
按已锁定的 D1 把上限同步调整为 **60 工具 / 32,000 字节**（`scripts/verify.mjs` 与 `test/deprecated.test.mjs`），
P5-2 到时只做最终复测。门禁仍然存在，只是跟着决定走——这正是它该有的行为。
### 39. P2 第一波余项：再挂 10 个 Excel 工具，又挖出两个「从来没生效」的 action（已落地）

第一波的下一半，仍然是**只挂桥里已有、没有出口**的能力：

| 工具 | action | 说明 |
|---|---|---|
| `wps_excel_copy_format` | copyFormat | 格式刷：只搬格式，不动值与公式 |
| `wps_excel_clear_formats` | clearFormats | 清格式留内容 |
| `wps_excel_get_conditional_formats` / `_remove_conditional_format` | getConditionalFormats / removeConditionalFormat | 条件格式的读与删（新增侧早有 set） |
| `wps_excel_get_data_validations` / `_remove_data_validation` | getDataValidations / removeDataValidation | 数据验证的读与删 |
| `wps_excel_refresh_links` | refreshLinks | 刷新外部链接，没有链接时如实报 0 条 |
| `wps_excel_consolidate` | consolidate | 合并计算 |
| `wps_excel_calculate` | calculateSheet | 强制重算（整簿或单表） |
| `wps_excel_group_columns` | groupColumns | 列分组（行分组早有工具） |

刻意**不挂**两个：`getActiveWorkbook`（`get_sheet_info` 已给工作簿与全部工作表名，只有 FullName 是新信息）
与 `unfreezePanes`（`wps_excel_freeze_panes { freeze: false }` 已覆盖）。同一能力挂两个工具，正是这一轮要减掉的冗余。

验收 `test/excel-missing-halves-2.test.mjs`（30 项，真实 WPS）又抓到**两个从未生效的 action**：

1. **`consolidate` 的函数常量是假的**。桥里写的是 `@{ sum = 9; count = 2; average = 1; max = 4; min = 5 }`，
   而 `Range.Consolidate` 要的是 `XlConsolidationFunction`：sum=-4157、count=-4112、average=-4106、
   max=-4136、min=-4139。9 不是这个枚举的合法值，每次调用都抛 HRESULT 0x800A03EC。
2. **`subtotal` 用了同一张假表**（`Range.Subtotal` 的 Function 也是 XlConsolidationFunction），
   而它**注册成工具、却一条测试都没有**——典型的有出口所以看起来没问题：没人调用就没人发现。

两处都换成真常量，并给 `subtotal` 补了验收（写 4 行数据 → 分组求和 → 已用范围真的多出汇总行）。

**还有一个 WPS 差异，是靠裸 COM 量出来的**：`Range.Consolidate` 在 WPS 里**只认 R1C1 引用，
而且表名必须带引号**——`'S1'!R1C1:R4C1` 成功，`S1!A1:A4` 与不带引号的含空格表名都**静默不写任何东西**
（不抛异常）。Excel 两种引用都认。桥里因此新增 `ConvertTo-ConsolidateSource`，把调用方的 A1 串
转成 `'表名'!R{行}C{列}`；解析不了的引用原样透传，所以 R1C1 输入仍然可用。

同一轮还把 `consolidate` 的默认值改对了：`topRow` 原本默认 **true**（按标签匹配），纯数字表会**什么都不写**
（量过：topRow=true 时 D1 空，false 时 D1=1+10）。默认改成 false（按位置逐格相加），
并在描述里写明合并计算是**按位置**写入（D1=来源1!A1+来源2!A1），不是一个总和。
`subtotal` 的 `groupBy`/`columns` 描述也改成真实语义：它们是 **range 内的序号**（从 1 开始），不是列名——
原来写「分组列标识、要汇总的列标识列表」，照字面传 "A" 会被 `[int]` 转换炸掉。

**验收**：30/30 绿。广告面 48 → **51 工具 / 26,652 字节**（`copy_format`/`clear_formats`/`calculate`
进精选档，其余注册可调用 + `wps_help` 可见）；注册 217 → **227**，桥 action 仍 **231**；
参数契约被校验的对 205 → **215**，A/B/C/D 仍全 0。
### 40. P2-2：表（ListObject）全族，8 个 action + 8 个工具（新 COM 代码）

P2 第二波是**新 COM 代码**——第一波只是把已有 action 挂出出口，这一波桥里此前没有任何 ListObject 能力：

| 工具 | action | 说明 |
|---|---|---|
| `wps_excel_create_list_object` | createListObject | 区域 → 表（表头/筛选/结构化引用），可命名、可指定首行是否标题与表格样式 |
| `wps_excel_get_list_objects` | getListObjects | 列出工作簿或某个表上的全部表：范围、行列数、列名、样式、总计行开关、结构化引用 |
| `wps_excel_add_list_row` | addListRow | 追加一行并可写入该行的值 |
| `wps_excel_delete_list_row` | deleteListRow | 删除第 N 行（表体行，从 1 开始） |
| `wps_excel_update_list_object` | updateListObject | 改名 / 换样式 / 显隐表头行 / 显隐筛选按钮 |
| `wps_excel_set_list_object_totals` | setListObjectTotals | 开关总计行 + 指定某列汇总方式（写 SUBTOTAL） |
| `wps_excel_resize_list_object` | resizeListObject | 调整表覆盖的范围 |
| `wps_excel_unlist_list_object` | unlistListObject | 转回普通区域（数据与格式保留） |

**动手前先用裸 COM 量支持面**。WPS 的 ListObject 实现是完整的：建表、ListColumns/ListRows、
DataBodyRange/HeaderRowRange/TotalsRowRange、TableStyle 赋值、Resize、Unlist 都可用，
总计行写出来的就是 `=SUBTOTAL(109,[Amount])`——结构化引用真的生效。所以这一族能直接做，不需要绕路。

**枚举值不凭记忆**（FIXES 39 刚吃过一次）：`XlTotalsCalculation` 实测 1=sum 2=average 3=count
4=countNums 5=max 6=min 7=stdDev 8=var，9/10 直接报“值不在预期的范围内”。总量一次、再写代码。

**又抓到一个 WPS 差异**：`ListRows.Delete()` / `Add()` 之后，**手里那个 ListObject 仍返回变更前的几何
信息**——`ListRows.Count` 还是旧值、`Range.Address()` 还是旧地址，表现就是“删掉一行，工具却报告表还是
4 行”。修法是在结构性变更后**重新解析一次表对象**再报告。这个缺陷只在“把读回来的结构给人看”的工具里
才会暴露：桥里第一波那些只回 `success` 的 action 永远发现不了它。

**验收**：`test/excel-list-object.test.mjs` **25 项**（真实 WPS）：建表 → 读结构（含结构化引用）→
加行（回读 A5:C5 确认真的落进表里）→ 删行 → 总计行（读 C5 公式确认 `SUBTOTAL(109,[Amount])`，
换函数后变 `SUBTOTAL(103,…)`）→ 关总计行 → 换样式 → 改名（并确认旧名不再解析）→ 调整范围 →
转回区域 → 数据仍在。

**一条写测试的教训**：第一版我把同一个 action 在断言里调用了两次（`ok(await call(...)) && text.includes(...)`），
第二次调用当然失败，5 个检查因此误报。改成“先取结果、再断言”后 25/25——**测量本身也要被检查**。

**数字**：桥 action 231 → **239**、注册工具 227 → **235**、广告面 51 → **54 工具 / 28,793 字节**
（`create` / `get` / `add_row` 进精选档）。`verify.mjs` 的 `EXPECTED_ACTIONS` 台账按规矩**故意**改到 239。
### 41. P2-3：Excel 页面设置 / 打印 / 外观 / 公式审计，8 个 action + 8 个工具

侦察先发现一件事：**Excel 此前完全没有页面设置能力**。桥里确实有 `setPageSetup` 与 `insertPageBreak`，
但两个都是 **Word** 的 action（由 word/format.ts 与 word/content.ts 驱动），Excel 侧只有 `setPrintArea`。
工具面在这一块是空的。

| 工具 | action | 说明 |
|---|---|---|
| `wps_excel_get_sheet_settings` | getSheetSettings | 一次读全：方向/纸张/页边距/缩放或按页适配/居中/网格线/打印区域/打印标题/页眉页脚/可见性/标签色/手动分页符 |
| `wps_excel_set_sheet_page_setup` | setSheetPageSetup | 方向、纸张、六个页边距（磅）、缩放或按页适配、居中、打印网格线与行列标题 |
| `wps_excel_set_sheet_print_titles` | setSheetPrintTitles | 每页重复的行/列（`$1:$1`、`$A:$A`） |
| `wps_excel_set_sheet_header_footer` | setSheetHeaderFooter | 页眉页脚六个位置，支持 `&P`/`&N`/`&D`/`&F`/`&A` 域代码 |
| `wps_excel_set_sheet_appearance` | setSheetAppearance | 可见 / 隐藏 / 深度隐藏（veryHidden）+ 标签色 |
| `wps_excel_set_outline_levels` | setOutlineLevels | 分级显示展开到第几级 + 汇总行/列位置 |
| `wps_excel_reset_page_breaks` | resetPageBreaks | 清除手动分页符 |
| `wps_excel_get_formula_audit` | getFormulaAudit | 引用来源 / 被引用 / 直接引用，可选画出追踪箭头 |

**刻意不做的两件事**（这是决定，不是遗漏）：**打印**（`PrintOut`）是物理副作用，模型不该随手触发；
**打印预览**（`PrintPreview`）会开模态窗口，把常驻宿主卡住——整条自动化链路都在等它。
两者都改成「把设置调好、由人来打印」。要的话可以单独评估（例如放进一个显式、会警告的工具）。

**测试抓到两处「成功但说错」**，都在我自己的新代码里：

1. **清掉标签色之后读回的不是「默认」**：`Tab.Color` 返回 0（黑），快照会显示成 `标签色: 0`，
   用户分不清「黑色」与「没设过」。修法是同时读 `Tab.ColorIndex`，等于 -4142（xlColorIndexNone）
   才显示 `(默认)`。
2. **常量格上 `Range.Formula` 返回的是值本身**（例如 `Region`），所以「这不是公式」不能靠
   `formula === ''` 判断。快照里加 `HasFormula`，文案改成 `不是公式，是常量（当前值: Region）`。

两条的共同点是：工具没报错，只是把事情**描述错了**——只有断言输出文本的测试才抓得到这类问题。

**验收**：`test/excel-page-setup.test.mjs` **24 项**（真实 WPS）：读快照 → 横向 A3 → 页边距与居中 →
按页适配 → 缩放 90% → 非法纸张被拒 → 打印标题 → 页眉页脚 → 标签色（含清回默认）→ 隐藏/深度隐藏/恢复
→ 非法可见性被拒 → 分级显示 → 清除分页符 → 公式审计（引用来源 2 处、被引用 1 处、常量格如实说明、
缺 `cell` 被拒）。写测试时还发现自己漏传 `sheet`：那几次设置其实落在「新建后成为活动表」的 Extra 上——
工具行为没错，是测试没锁住目标，已改成显式传 sheet。

**数字**：桥 action 239 → **247**、注册工具 235 → **243**、广告面 54 → **56 工具 / 30,885 字节**
（离 D1 上限 32,000 只剩 1,115 字节——P5-2 的重定已经不是可选项）。
### 42. D1 上限第二次上抬：60 / 32,000 → 70 / 40,000（用户决定）

P2-3 结束时广告面是 **56 工具 / 30,885 字节**，离 D1 的上限只剩 1,115 字节，而后面还有 P2-4（高级项）、
P3（Word 长尾，档位是「全都要」）与 P4。把选择摆出来之后（A 抬上限 / B 守住 60 并降级低频工具 /
C 新工具一律不广告），用户选了 **A**。

- 新上限：**70 工具 / 40,000 字节**（`scripts/verify.mjs` 与 `test/deprecated.test.mjs` 同步）；
- 理由沿用 D1 当初的推理：广告面是「一站式」的瓶颈，请求前缀可被 prompt cache 复用，边际成本低；
  `minimal` 档继续留给成本敏感场景（4 工具 / 1,348 字节）；
- P5-2 仍然负责「最终复测」，但不再需要在这一轮做取舍。

这条本身是小改动，但它是一次**有意识的放宽**：门禁数字跟着决定走，每次变动都留档——
下次有人看到 70 这个数，能查到它是怎么来的、为什么不是 60。
### 43. P2-4：高级项 9 个工具，并修好一个「一直建不出透视表」的既有工具

P2-4 依旧先量后写（每一族都用裸 COM 试过支持面）：

| 工具 | action | 说明 |
|---|---|---|
| `wps_excel_get_pivot_tables` | getPivotTables | 列出透视表：名字、区域、行字段、数据字段 |
| `wps_excel_refresh_pivot_tables` | refreshPivotTables | 刷新一张，或一张表上的全部 |
| `wps_excel_clear_pivot_table` | clearPivotTable | 清除报表区域（见下：刻意不说成「删除」） |
| `wps_excel_refresh_all_data` | refreshAllData | 工作簿全部刷新（外部数据 + 透视表） |
| `wps_excel_goal_seek` | goalSeek | 单变量求解 |
| `wps_excel_add_sparkline` / `_clear_sparkline` | addSparkline / clearSparkline | 迷你图（line/column/winloss，可标点） |
| `wps_excel_delete_chart` | deleteChart | 按名或序号删图表；只有一张时可省略 |
| `wps_excel_set_chart_labels` | setChartLabels | 图表标题 + 横轴/纵轴标题 |

**最大的收获不是新工具，而是一个既有工具其实从来没工作过。** 写验收要先建一张透视表，用的是
**本来就存在、而且还在广告面里**的 `wps_excel_create_pivot_table`——它报 `HRESULT 0x800A03EC`；
而同一串调用在**裸 COM 里完全正常**。逐步定位后，根因在桥里的一个 helper：

```powershell
function Get-RangeFromAddress($workbook, [string]$address) {
    return $workbook.ActiveSheet.Range($address)   # ← 这里
}
```

`Range` 是**可枚举**的：多格 Range 在函数返回时经过 PowerShell 输出流会被**展开成单格数组**
（A1:C7 → 21 个 range）。于是 `PivotCaches().Create(1, $range)` 拿到的是 21 元素的 `System.Object[]`，
WPS 回 `0x800A03EC`。加前置逗号 `return ,$range` 才会把 Range 本身返回。
验证方式是三种形状并排试：不经函数返回的 `$wb.ActiveSheet.Range(...)` 成功、带 Version 参数的失败；
修完后同一个 `Create` 立刻给出**有意义的**业务错误（探针数据没写表头），说明 Range 已被正确接受。

为什么一直没人发现：这个 helper **只被** `createPivotTable` 调用，而透视表工具**没有任何测试**。
这和 P2 里那三个「从未生效的 action」是同一类问题——**没有断言真实输出效果的测试，就没有人会知道**。
现在 `test/excel-advanced.test.mjs` 会真的建一张透视表并读回它的行字段。

**顺手补上错误上报**：`createPivotTable` 原本没有 try/catch，出错只剩一个 HRESULT；现在按步骤名
上报（`createCache: …` / `rowFields: …`），正是靠它把范围缩到 `Create` 那一行的。

**两处刻意推迟**（按实测，不是遗漏）：

- **切片器**：`SlicerCaches.Add2` 能建出缓存（名字如「切片器_Region」）也能 Delete，但
  `SlicerCaches(1).Slicers.Count` 始终是 0——用户可见的切片器并没有真的出现，接上去只会是
  「成功但没效果」。
- **场景管理器**：WPS 把 `Worksheet.Scenarios` 暴露成一个**方法**（要写 `$s.Scenarios()` 才拿到集合，
  直接 `$s.Scenarios.Add` 会报「PSMethod 不含 Add」）。集合本身能用，但默认方案与 Show 的效果在 COM 上
  读不干净。

两者的探针脚本都留在 `test/.artifacts/e2e/p24-probe*.ps1`，要接的时候不用重新摸。

**另一处诚实的措辞**：`clear_pivot_table` 不叫 delete——WPS 清掉 `TableRange2` 之后 PivotTable 对象
**仍留在集合里**（再读它会 E_FAIL），所以工具报告「报表已清除，对象会留到保存/重开」并回读剩余数量，
不谎称删除。

**验收**：`test/excel-advanced.test.mjs` **24 项**（真实 WPS）：建表 → 列表（含行字段）→ 刷新
（单张/全部/错名被拒）→ 清除（含剩余对象数）→ 全部刷新 → 单变量求解（A10=25、B10=50，并回读）→
迷你图（加 line/column、清、缺参数被拒）→ 图表标题三件套 → 删图 → 再删被拒。

**数字**：桥 action 247 → **256**、注册工具 243 → **252**、广告面 56 → **59 工具 / 32,924 字节**
（`goal_seek`、`clear_pivot_table`、`set_chart_labels` 进精选档）。
### 44. P2-5：一键 e2e 扩成两个真实场景（第二个覆盖 ListObject + 条件格式 + 打印）

`scripts/e2e.mjs` 原来只跑一个「Excel 汇总 + Word 结论」场景（19 项检查）。P2-5 往**同一次** headless
运行里加了第二个独立的 Excel 任务：新 fixture `orders.xlsx`（订单表 A1:D13，12 行订单，金额里 6 个大于 500），
要求模型把 A1:D13 变成真正的表（ListObject，表名 `Orders`）、给「金额」列加「大于 500 标红」的条件格式、
设成横向 A4 打印（页脚居中「第 &P 页 / 共 &N 页」、第 1 行每页重复），最后保存关闭。

独立验证（裸 COM 重开文件读）新增 9 项：订单工作簿存在，以及表对象存在 / 表名 / 覆盖 A1:D13 /
条件格式条数 / 横向 A4 / 页脚含 `&P` / 打印标题非空 / 读得到订单工作簿。检查数 **19 → 28**。

**验证逻辑先单独验过，再去赌模型**：把 `verifyScript()` 从 e2e.mjs 里抽出来，对一份「手工驱动插件工具
做好」的 orders.xlsx 跑一遍，读回
`{name: Orders, range: $A$1:$D$13, listCount: 1, formatCount: 1, orientation: 2, paperSize: 9,
centerFooter: 第 &P 页 / 共 &N 页, printTitleRows: $1:$1}`——**先证明「正确的产物能被认出来」**，
否则一个写错的验证脚本会让整件事失去意义。

**一处刻意的宽松**：条件格式落在哪个 Range 上是模型的选择（`D2:D13`、`D:D`、`D1:D13` 都合理），
而 `FormatConditions` 是**按 range 对象**而不是按表统计的，所以验证脚本对几个候选区域各读一次取最大值。
这不是放水，是避免把「模型的合理选择」判成插件缺陷。

**首次完整运行就通过**：`E2E OK (28 checks) in 94s`——96 次 `wps_*` 调用、2 次技能加载、
4 条 pwsh 命令里没有一条自己写 COM，第二场景的产物与预期逐项一致。默认超时从 300 提到 420 秒。
### 45. P3-1/P3-2：Word 深水区第一二波——挂出 4 个已有能力 + 表格族 9 个新 action

P3 依旧「先裸 COM 量遍 D2 清单，再动手」。第一波把桥里已有、却没有出口的 4 个 Word 能力挂出来：
`get_bookmarks`、`get_comments`、`get_document_stats`、`insert_hyperlink`。

第二波是**表格族的新 COM 代码**（9 个 action + 9 个工具）：

| 工具 | action | 说明 |
|---|---|---|
| `wps_word_get_tables` | getDocumentTables | 列出表格：序号、行列数、样式、文本预览 |
| `wps_word_get_table_data` | getTableData | 按行列读出全部单元格文本（合并格留空） |
| `wps_word_set_table_cell` | setTableCell | 写单元格（行列从 1 开始） |
| `wps_word_add_table_lines` | addTableLines | 加行/加列（kind + count + position） |
| `wps_word_delete_table_line` | deleteTableLine | 删行/删列 |
| `wps_word_merge_table_cells` | mergeTableCells | 矩形区域合并 |
| `wps_word_split_table_cell` | splitTableCell | 单元格拆分 |
| `wps_word_set_table_format` | setTableFormat | 样式 / 边框 / 自适应 / 表头底纹 |
| `wps_word_convert_table_to_text` | convertTableToText | 表格转文本 |

**两处实测缺口（写清楚，不硬做）**：

- **水印**：Word 的通行做法是在页眉里放一个 WordArt/TextEffect，但 WPS 的**页眉 Shapes 集合不接受任何图形**
  ——`AddTextEffect` / `AddShape` / `AddTextbox` 都返回了对象，`Shapes.Count` 却始终是 0；
  `Selection.HeaderFooter` 还是 null。没有可靠路径，本轮不做。
- **文档属性**：`$doc.BuiltInDocumentProperties` / `CustomDocumentProperties` 在 WPS 里是**坏壳**——
  属性本身非 null，但 `.Count` 读出来是空、`.Item('Title')` 与 `GetType()` 直接抛
  「Object reference not set」。既写不进去也读不出来，本轮不做。

两者的探针脚本都留在 `test/.artifacts/e2e/word-probe*.ps1`，要接的时候不用重新摸。

**一处工程细节值得记**：`mergeTableCells` 最初被生成器判成「动态 action」并**拒绝构建**——原因不是代码，
而是我在注释里写了 `$p` 加 `.` 加 `$key` 这样的字面量，正好命中生成器的动态模式。守卫按设计工作了：
它读不出键集的动作必须显式声明，不许静默跳过。把注释改成文字描述即可。

**验收**：`test/word-deep.test.mjs` **27 项**（真实 WPS Writer）：书签/批注/统计/超链接，以及表格的
读结构 → 写单元格 → 加行加列 → 删行 → 合并 → 拆分 → 设置外观 → 转回文本，外加 5 项负向
（缺参数、越界、非法颜色）。

**数字**：桥 action 256 → **265**、注册工具 252 → **265**、广告面 59 → **62 工具 / 33,836 字节**
（`get_document_stats` / `get_tables` / `get_table_data` 进精选档）；未工具化 action 台账 11 → **7**
（剩下的都是刻意的重复实现或 P4 的项）。
### 46. P3-3：Word 文档生产族，6 个 action + 6 个工具（新 COM 代码）

| 工具 | action | 说明 |
|---|---|---|
| `wps_word_insert_page_numbers` | insertPageNumbers | 给某节的页眉/页脚插页码（左/中/右、是否首页显示） |
| `wps_word_set_columns` | setColumns | 分栏：栏数、栏间距、分隔线；count=1 即取消分栏 |
| `wps_word_get_revisions` | getRevisions | 列出修订（插入/删除/替换…）并报告修订跟踪开关 |
| `wps_word_accept_revisions` | acceptRevisions | 接受一处或全部修订 |
| `wps_word_reject_revisions` | rejectRevisions | 拒绝一处或全部修订 |
| `wps_word_delete_comment` | deleteComment | 删除一条或全部批注 |

**一处「成功但说错」被自己的验收抓到**：单栏时 WPS 的 `TextColumns.Spacing` 返回哨兵值 **9999999**，
工具原样印成「栏间距 9999999 磅」——用户会以为出了 bug。改成只在 count > 1 时才报告间距。
这类问题不会让任何断言失败（断言只看「现在是 1 栏」），是**读输出文本**时才发现的。

**验收**：`test/word-produce.test.mjs` **19 项**（真实 WPS Writer）：页码落页脚/页眉并对齐、非法节号被拒、
两栏↔一栏、非法栏数被拒、修订跟踪开启→改文档→列出修订（类型「插入」、作者）→拒绝一处→再改→全部接受
→确认没有残留、两条批注→按序号删一条→无序号删全部→确认清空。

**数字**：桥 action 265 → **271**、注册工具 265 → **271**、广告面 62 → **64 工具 / 34,754 字节**
（`insert_page_numbers` / `get_revisions` 进精选档）。
### 47. P3-4：Word 长尾（内容控件 / 脚注尾注 / 索引 / 交叉引用 / 邮件合并）——P3 完成

| 工具 | action | 说明 |
|---|---|---|
| `wps_word_get_content_controls` | getContentControls | 列出内容控件：类型、标题、标签、当前文本 |
| `wps_word_add_content_control` | addContentControl | 插入内容控件（富文本/纯文本/复选框/下拉/日期/图片） |
| `wps_word_add_footnote` / `_add_endnote` | addFootnote / addEndnote | 插入脚注 / 尾注 |
| `wps_word_get_notes` | getNotes | 列出脚注与尾注及正文 |
| `wps_word_insert_index` | insertIndex | 文档末尾插入索引 |
| `wps_word_insert_cross_reference` | insertCrossReference | 交叉引用（枚举照实透传） |
| `wps_word_mail_merge` | mailMerge | CSV 邮件合并 → 生成新文档（母版不动） |

**一个真实的根因，被验收逼出来**：加过脚注之后，Word 会把光标留在**注释正文**里
（`Selection.StoryType` 不再是正文）。此后再执行「往光标处插内容」的动作——插尾注、插内容控件、
插超链接、邮件合并插字段——全都插进了注释，正文里什么都没有。表现是邮件合并「成功」了、
生成的文档里却没有数据行，尾注干脆报「值不在预期的范围内」。

修法是新增一个共享解析器 `Get-MainTextRange`：光标在正文就用光标，**不在正文就落到正文末尾并如实告警**
（`Add-WpsWarning`），而不是悄悄插到别处。P3 动过的 5 个用光标的动作全部改走它。
这条也解释了为什么「一次只验证一个动作」不够——**动作之间会互相改变状态**，必须串起来跑。

**两处小修正**：内容控件返回的是数字类型而不是名字（工具层要靠数字映射中文名）；
另外发现桥文件里**混着 CRLF 与 LF**（新插入的块是 LF、老区域是 CRLF），按行匹配的脚本必须留意。

**验收**：`test/word-longtail.test.mjs` **20 项**（真实 WPS Writer）：内容控件增读与非法类型被拒、
脚注/尾注增读（正文从 `.Reference.Text` 读——`.Range.Text` 在 WPS 里是空的）、插入索引、
按书签插交叉引用、缺参数被拒、真实 CSV 邮件合并（缺文件/缺参数被拒 → 生成新文档并含两行数据）。

**数字**：桥 action 271 → **279**、注册工具 271 → **279**、广告面 64 → **66 工具 / 35,568 字节**
（`mail_merge`、`get_notes` 进精选档）。测试 **489 → 509 项 / 24 文件**。
### 48. P4：PPT 做减法第一波——四个碎片 setter 合并成一个，3D 族按 D3 删除

P4 是**减法**，与前三阶段的加法不同。第一波做两件事：

1. **P4-2 聚合碎片 setter**：`set_shape_shadow` / `set_shape_gradient` / `set_shape_border` /
   `set_shape_transparency` 四个工具（各带一套同义却不同名的键：`color`/`weight`/`style`/`color1`/`color2`/
   `stops`/`transparency`…）合并为 **`wps_ppt_set_shape_effect`**：阴影、边框、渐变、透明度四族改用
   **带前缀的明确键**（`shadowColor`、`borderWidth`、`gradientColor1`、`transparency`…），只改给出来的项，
   并回报 `applied` 列表。四个 action 也一并从桥里删除——只删工具不删 action，那些 action 会变成
   「没有工具驱动」，把未工具化台账从 7 顶上去。
2. **D3 明确放弃的 3D 族删除**：`set_3d_rotation` / `set_3d_depth` / `set_3d_material` 三个工具与
   三个 action 一并删除。

**加了一个「空操作」闸门**：合并后的工具在**一个效果都没给**时直接报错（`nothing to apply`），
而不是回一个「成功」却什么都没做——这正是 P2/P3 反复遇到的「成功但没做事」的预防。

**过程里踩的坑**（值得记）：第一版删除脚本用「块文本长度」算删除区间，遇到 CRLF 就切错位置，
生成出 69 个 parse error；改成按行正则匹配（`^    }$` / `^};$`，都允许 `\r?`）后一次成功。
教训：**用字符偏移做文本外科手术时，必须先确认换行符**——这个仓库的文件里 CRLF 与 LF 是混着的。

**验收**：`test/ppt-slimming.test.mjs` **16 项**（真实 WPS Presentation）：合并工具一次应用四族效果并
逐项回报、只给一项时只报那一项、什么都不给被拒、非法线型被拒，以及**七个被删的工具名逐个确认已无法解析**。
既有 `ppt-contract-fixes.test.mjs` 里 4 项针对旧工具的检查按新契约改写（3D 那项删除），56/56 绿。

**数字**：PPT 工具 **88 → 82**、注册工具 279 → **273**、桥 action 279 → **273**；广告面不变
（被删的都不在精选档）。测试 **509 → 524 项 / 25 文件**。
### 49. P4 第二波：三组合并（动画 / 表格样式 / 页脚三件套）

按用户要求，把剩下三组碎片 setter 也合并掉：

| 合并前 | 合并后 | 合并方式 |
|---|---|---|
| `add_animation` + `add_animation_preset` + `add_emphasis_animation` | **`wps_ppt_add_animation`** | 给 `preset` 走整页预设；给 `effectKind=emphasis` 走强调；否则入场/退场 |
| `set_table_style` + `set_table_cell_style` + `set_table_row_style` | **`wps_ppt_set_table_format`** | 用 `row`/`col` 定作用域：两者都有＝单元格，只有 row＝整行，都没有但有样式键＝整张表；并吸收原 `set_table_style` 的位置尺寸职责 |
| `set_slide_number` + `set_ppt_footer` + `set_ppt_date_time` | **`wps_ppt_set_slide_footer`** | 一次设页码 / 页脚文字 / 日期，只改给出来的项 |

**两个被自己的验收抓到的缺陷**（都是合并过程中引入的）：

1. **`setAnimation` 的 preset/emphasis 分支漏了 `exit`**——输出完还继续走到入场分支，于是
   `effectKind=emphasis` + `effect=pulse` 会被入场分支当成非法效果再报一次错，宿主回给工具的是后一个错误。
   教训：**合并分支时，每个提前返回的分支都要 `exit`**。
2. **页脚工具没报告「页脚本身」的可见性**——旧测试里那条「footer reports hidden」要的正是它，
   而我合并时只报告了页码与日期。补上 `footerVisible` 后通过。这是「旧断言在保护你没写坏新代码」的例子。

**顺带**：`ALIAS_DEBT` 台账 62 → **59**（三组合并消掉 3 处改名），棘轮跟着收紧。

**验收**：`test/ppt-slimming.test.mjs` 扩到 **32 项**——合并后的表格格式化三种作用域（单元格/整行/仅位置）
与空操作拒绝、页脚一次设三件套与空操作拒绝、动画缺形状被拒，加上 **15 个被删工具名逐个确认已无法解析**。
既有 `ppt-contract-fixes` 里 6 项按新契约改写，57/57 绿。

**数字**：PPT 工具 82 → **76**（P4 合计 88 → 76）、注册 273 → **267**、桥 action 273 → **267**；
广告面不变（66 / 35,568 字节）。测试 524 → **541 项 / 25 文件**。
## 新发现的 WPS / Office 差异

- **WPS 的 Presentations.Add() 返回 0 页演示文稿**，PowerPoint 返回 1 页。
  任何“新建演示文稿后立刻操作第 1 页”的流程都会 E_FAIL。
  技能文档必须写明：新建后先 add_slide。
- Range.Value2 返回 Object[,]，ConvertTo-Json 会把它压平；
  必须显式转换并用 GetLowerBound() 判断下界（这也解释了上游为何逐格读）。- **`Find()` 结果的 `Address()` 在常驻宿主里不可用**（0 参/2 参都抛），但 `UsedRange.Address()` 与
  `Selection.Address()` 正常。凡是「用 Find 定位、再读地址」的写法在这里都是死的（第 38 条）。
- **PowerShell 的 `,` 比 `+` 绑得更紧**：`$a[$i + 1, $j + 1]` 不是「两个下标」，而是 `$i + (1, $j) + 1`，
  结果是 `Object[]` 没有 `op_Addition`。二维组取下标的算式必须写成 `$a[($i + 1), ($j + 1)]`。
- **`Range.Consolidate` 只认 R1C1，且表名必须带引号**：`'S1'!R1C1:R4C1` 能用，
  `S1!A1:A4` 与不带引号的含空格表名都**静默不写任何东西**（不抛异常）。Excel 两种引用都认。
  跨应用搬运数据时，这种「不报错的空操作」比报错更难查（第 39 条）。
- **表对象在结构性变更后不刷新**：`ListRows.Delete()`/`Add()` 之后，手里那个 ListObject 仍返回旧的
  `ListRows.Count` 与旧的 `Range.Address()`，要重新取一次对象才是新结构。Excel 会当场更新（第 40 条）。

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

| test/close-safety.test.mjs | 14 | 关闭文档不得弹模态框、不得泄漏工作簿/演示文稿 |

| test/excel-contract-fixes.test.mjs | 35 | 25 处参数错配逐个真实验证、跨应用批注污染 |

| test/file-ops.test.mjs | 11 | 另存/打开的路径、按应用转换、页眉页脚分节 |
| test/merged-tools.test.mjs | 31 | 合并后的转发、参数改名与隐藏 |
| test/warnings.test.mjs | 7 | warnings 机制：尽力而为的失败如实回传 |
| test/ppt-contract-fixes.test.mjs | 56 | PPT 参数契约逐项修复（P4 后 3D 那项改为测合并工具） |
| test/ppt-slimming.test.mjs | 32 | P4 合并后的形状效果/表格/页脚工具 + 15 个被删工具确认消失 |
| test/word-lifecycle.test.mjs | 17 | e2e 暴露的 5 个缺陷（第 24～29 条） |
| test/spec-reproduction.test.mjs | 12 | P1 验收：spec 逐字节复现模型可见面 |
| test/excel-missing-halves.test.mjs | 18 | P2 第一波：工作表信息、自动尺寸 ×3、自动换行、查找定位、命名范围读删 |
| test/excel-missing-halves-2.test.mjs | 30 | P2 第一波余项：格式刷/清格式、条件格式与数据验证读删、重算、外部链接、合并计算、列分组、分类汇总 |
| test/excel-list-object.test.mjs | 25 | P2-2 表（ListObject）：建表/读结构/增删行/总计行/样式/改名/范围/转回区域 |
| test/excel-page-setup.test.mjs | 24 | P2-3 页面设置/打印标题/页眉页脚/外观/分级显示/分页符/公式审计 |
| test/excel-advanced.test.mjs | 24 | P2-4 透视表列表/刷新/清除、全部刷新、单变量求解、迷你图、图表标题与删除 |
| test/word-deep.test.mjs | 27 | P3 书签/批注/统计/超链接 + 表格读写编辑与外观 |
| test/word-produce.test.mjs | 19 | P3-3 页码/分栏/修订接受拒绝/批注删除 |
| test/word-longtail.test.mjs | 20 | P3-4 内容控件/脚注尾注/索引/交叉引用/CSV 邮件合并 |

合计 **541 项**（25 个测试文件），加 `node scripts/verify.mjs` **23 项**门禁（含 70 工具 / 40,000 字节预算与 action 数量三方一致）。

另有 node scripts/param-contract.mjs：零副作用地把 255 对工具/action 的参数契约对账一遍，
结果写入 docs/param-contract.md。A/B/C/D 四类静默失效**均为 0**；剩下的 1 处「桥无键表」（`setCellFormat`，
动态键闸门跳过）与 6 处「handler 实参静态读不出」都在报告里逐名列出，不做隐藏。
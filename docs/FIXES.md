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

| test/close-safety.test.mjs | 14 | 关闭文档不得弹模态框、不得泄漏工作簿/演示文稿 |

| test/excel-contract-fixes.test.mjs | 35 | 25 处参数错配逐个真实验证、跨应用批注污染 |

| test/file-ops.test.mjs | 11 | 另存/打开的路径、按应用转换、页眉页脚分节 |

合计 197 项，加 node scripts/verify.mjs 22 项门禁（含 45 工具 / 25,000 字节预算）。

另有 node scripts/param-contract.mjs：零副作用地把 212 对工具/action 的参数契约对账一遍，
结果写入 docs/param-contract.md。它不参与通过/失败门禁（当前仍有 85 处待修），
但把「还有多少静默失效」变成了一个随时可查、只会变小的数字。
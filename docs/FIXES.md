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
| test/merged-tools.test.mjs | 31 | 合并后的转发、参数改名与隐藏 |
| test/warnings.test.mjs | 7 | warnings 机制：尽力而为的失败如实回传 |
| test/ppt-contract-fixes.test.mjs | 57 | PPT 参数契约逐项修复 |
| test/word-lifecycle.test.mjs | 17 | e2e 暴露的 5 个缺陷（第 24～29 条） |

合计 **310 项**（15 个测试文件），加 `node scripts/verify.mjs` **23 项**门禁（含 45 工具 / 25,000 字节预算与 action 数量三方一致）。

另有 node scripts/param-contract.mjs：零副作用地把 211 对工具/action 的参数契约对账一遍，
结果写入 docs/param-contract.md。A/B/C/D 四类静默失效**均为 0**；剩下的 1 处「桥无键表」（`setCellFormat`，
动态键闸门跳过）与 9 处「handler 实参静态读不出」都在报告里逐名列出，不做隐藏。
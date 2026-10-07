# 找 bug 工单（第一轮 · W1 断言审计 + W4 日志挖掘）

> **来源**：D1-A —— 只做审计、只出工单（2026-10-04）。
> **进度**：第 1 批（W4-1 + W6）与第 2 批（W1-1 + W1-2）**已完成** → FIXES **94 / 93**。
> **第 3 批（W1-3/W1-4）进行中**：已建共用 oracle 模块 `test/lib/oracle.mjs`，完成 **1 / 40** 个文件
> （`destructive-guard`，45 → 53 条断言，破坏路径首次有真值回读）。
> 审计口径与原始数据都在本文，便于复核；修的时候按 **D2-B 攒批**，一批修完统一验证。
> 工单编号规则：`W<工作流>-<序号>`。每条都标了「怎么验」与「预期产出」。

---

## 0. 本轮实测口径（可复核）

| 指标 | 值 | 怎么量的 |
| --- | --- | --- |
| 静态 `check(` | **932** | `test/*.test.mjs` 里 `check(` 出现次数（运行时合计 1085，差额来自循环） |
| 带"近似特征"的断言 | **365** | 含 `includes()` / 正则 / `>0` / `index[0]` |
| `index[0]` 断言 | **23** | **P5 的盲区原型**（只查第 0 个，缺陷打的是相邻那个） |
| `>0` / `>=0` 断言 | **27** | 只证明"非空" |
| `includes()` 断言 | **184** | 只证明"输出里有 X" |
| 正则断言 | **144** | 同上，且容易写宽 |
| **有独立 oracle 的测试文件** | **8 / 48** | 其余 **40 个只信插件自己的回报** |
| 会话日志语料 | **87 个 / 68 MB** | `~/.dsh/sessions/**/session.v4.jsonl.zstd` |
| 其中的 wps 真实调用 | **430 次 / 13 个会话 / 18 次失败** | 逐帧 zstd 解 + `tool/ptc-dispatch` 配对 |
| 5 个参数契约 UNPARSED 工具 | 见 W6 | `node scripts/param-contract.mjs` |

**判据（P5 的教训）**：套件全绿 ≠ 正确。P5 漏过去是因为断言查 `styles[0]`（第 1 段），
而缺陷打的是**紧邻的上一段**。所以本轮审计的重点是**"这条断言到底证明了什么"**，不是覆盖率。

---

## W4 · 来自真实用法（优先级最高：有现场、可复现）

### W4-1 ✅ 已修（FIXES 94）· `wps_excel_export_chart_as_image` 对不存在的图表名返回**裸 E_FAIL**

> **结果**：与同族 `delete_chart` 完全对齐 —— 收名字或序号、整张表只有一张图时可省略 `chartName`、
> 找不到时**列出表上现有的图表名**、`.Export()` 补上保护。回归 `test/contract-gaps.test.mjs` 26/26。

**现场（真实会话，2026-09-12）**：模型建图后导出，拿到的是一句无法据此恢复的报错。

**2026-10-04 在 0.6.2 上复现**（新建工作簿 → 建 `Chart 1` → 导出）：

| 调用 | 结果 |
| --- | --- |
| `export_chart_as_image { chartName: "Chart 1" }` | ✅ 成功 |
| `export_chart_as_image { chartName: "图表 1" }` | ✅ 成功（WPS 两种名字都认） |
| `export_chart_as_image { chartName: "NoSuchChart" }` | ❌ `对 COM 组件的调用返回了错误 HRESULT E_FAIL。（动作：exportChartAsImage）下一步：确认操作对象存在、名称或序号正确，并且 WPS 处于可操作状态后重试。` |
| **同族对照** `delete_chart { chart: "NoSuchChart" }` | ❌ `chart not found on this sheet（给 chart 名称，或在只有一张图时省略）` ← **同一个桥，消息质量完全不同** |

**三个问题**：
1. 原因被漏成裸 HRESULT，且「下一步」把原因**误导**成"WPS 可能不健康"——真实原因是名字不对。
2. **没有"列出图表"的工具**：`wps_help` 搜「图表」返回 10 个工具，`create/delete/export/update/set_labels/add_sparkline` 都在，**独缺查询型**。工具自己的描述还写着"可通过 `create_chart` 返回值或界面查看"——跨会话或多图表时**无从得知名字**。
3. `chartName` **必填**，而同族的 `delete_chart` 允许"只有一张图时省略"。同一族两套规矩。

**怎么验**：新建工作簿 → 建图 → 用不存在的名字导出，看是否给出"表上有哪些图"。
**预期产出**：报错改名 + 至少一条可发现路径（列表工具，或失败时回报已有图表名/序号，或与 `delete_chart` 一样支持省略/序号）。

### W4-2 ⛔ 已排除（记录以免重复投入）· `export_range_as_image` 的 JSON 污染

**现场**：2026-09-12 两次 `导出区域为图片出错: PowerShell 输出解析失败（非有效JSON）: True True False True {"success":true,...}`
—— 典型的**宿主 stdout 被多打的布尔值污染**。

**2026-10-04 复测 0.6.2**：`export_range_as_image { range: "B2" }` **成功**，写出 880 字节真图。
**结论：老版本的问题，现已不复现**。记录在此，免得下次又当新 bug 挖一遍。
（教训：**计划也是假设** —— 日志里的失败必须先在当前版本复测才能立项。）

### W4-3 真实用法画像（给排期用）

- 调用最多的门面：`wps_call` **162** 次、`wps_help` **55** 次 —— 模型大量走"先查后调"。
- **经 `wps_call` 调用的目标里，`wps_word_insert_text` 占 27 次**（它**已在广告面**上）；
  其次是 `close_workbook` 15、`close_document` 13、`get_paragraphs` 13、`create_document` 9。
  → 值得单独查一次"模型为什么绕门面"：是广告面描述不够、还是模型习惯（**这是一个可查的问题，不是结论**）。
- 18 次失败里：**3 次环境**（WPS 没开）、**11 次是本次会话故意触发的契约验证**、**2 次即 W4-1/W4-2**。
  → 真实失败率很低（4%），**说明"问题在质量不在可用性"**，与 W1 的判断一致。
- 最长耗时：`wps_status` 1731ms、`wps_common_get_app_info` 1857ms、`wps_excel_get_series` 类 768ms —— 无挂起前兆。

---

## W1 · 断言证明力（批量、纯静态，适合攒批一起修）

> **通用升级配方**：把"断言插件自己的话"换成"读回真值再比对"。
> 真值来源优先级：① 磁盘文件用外部库读（openpyxl / python-docx / python-pptx）② 裸 COM 读 ③ 才轮到插件自己的读接口。
> **禁止**用被验证工具自己的读接口当 oracle（读写同源会一起错）。

### W1-1 ✅ 已执行（FIXES 93）· 23 条 `index[0]` 断言 → "精确集合/全量比对"

> **结果**：23 条里 **12 条真弱已升级**；11 条经甄别为**假阳性**（`[0]` 在详情字符串里）或**已被 `length === 1` 守卫**。

**为什么**：这正是 P5 漏过去的形状。只断言第 0 个元素，相邻元素被改坏时不会红。
**例子**：`excel-page-setup.test.mjs:63` 「tab colour is applied (BGR 39423 for #FF9900)」——只查第一个 sheet 的颜色。
**怎么做**：把 `x[0] === expected` 换成 `JSON.stringify(x) === JSON.stringify(expectedAll)`。
**怎么验**：改完**先在旧代码上跑一次**，确认新断言**会红**（否则它证明不了任何东西）。

### W1-2 ✅ 已执行（FIXES 93）· 27 条 `>0` 断言 → 精确计数/精确值

> **结果**：**7 条已升级**；其余 20 条判定为**故意的烟测**（覆盖矩阵的 "no hang" 检查，属 W1-3 范畴）
> 或**本来就在断言存在性**（pid > 0、实例数 > 0、预算 > 0）—— 改反而错。详见 FIXES 93。

**为什么**：`>0` 只证明"非空"。历史上"回 success 但结果不对"正是这种断言放行的。
**例子**：`destructive-guard.test.mjs:123` 「deleteNamedRange reports the reference it removed」只断言非空。
**怎么做**：断言**具体条数**与**具体名字**；破坏性操作还要断言**剩下的东西还在**。

### W1-3 · 40 个文件没有独立 oracle → 按「破坏力 × 调用频次」排优先级补

**缺口最大的（建议优先）**：
1. **`destructive-guard.test.mjs`（46 条断言，0 条 oracle）** —— 破坏性操作是出事代价最高的面，
   却只用"回报文本"验证。**这是最该补真值 oracle 的文件。**
2. `excel-page-setup.test.mjs`（25 条 / 20 条 includes）—— 打印设置类，回报写了 ≠ 文件里真是那样。
3. `excel-advanced`、`excel-missing-halves-2`、`word-deep`、`word-produce`、`excel-list-object` —— 都是"写入型"操作。
**怎么做**：在合适的地方存盘（`.xlsx`/`.docx`/`.pptx`），用 openpyxl / python-docx / python-pptx 读回真值；
本机已装这三个库；本轮已用它验证过 Excel 的数字格式/底色/条件格式/列宽**确实落盘**。

### W1-4 · `destructive-guard` 的 12 条"只断言回报文本" → 补"删除前的真值 + 删除后的真值"

**例子**：`clear_range` / `delete_rows` / `delete_sheet` / `remove_conditional_format` /
`deleteListRow` / `clear_sparkline` / `delete_chart` / `removeDataValidation` 等。
**怎么做**：断言"报告说删了 N 行" **且** "读回确实少了 N 行" **且** "没被删的还在"。

---

### W4-4 ✅ 已修（FIXES 100）· Word 与 Excel 的批注预览不一致

Word 的 `Comments.Item(i).Range.Text` 带结尾批注标记（`\r`），Excel 的 `Comment.Text()` 不带。
桥侧三处统一按既有约定清洗（同 `deleteTableLine`）。
**验证**：`contract-gaps` 31/31（+5 条，直接比对 `data.impact.preview` 的原始字节）。

**原始记录（保留背景）**：

**来源**：执行 W1-2 时加强断言逼出来的（FIXES 93）。
`deleteComment`（Word）回报的预览是 `"S3 批注内容\r"` —— **带段落标记**；
`deleteCellComment`（Excel）回报的是 `"第二条批注"` —— **不带**。
影响面是给人看的"将删除什么"预览文本（多一个换行/回车），**轻微**。要统一应在桥侧 trim 后再进 preview。
**怎么验**：两边各建一条批注、删掉、比对 `data.impact.preview` 的原始字节。

### W4-5 ✅ 已处理（FIXES 101）· 插入类工具不返回对象名 —— **原记录一半已过时，另一半是真缺陷**

**前提修正**：工单说 `wps_ppt_add_shape` / `wps_ppt_insert_ppt_image` 都不回对象名。实测：
- `add_shape` **早就回了**（`形状名称: 矩形 1`，slide-ops.ts:784）—— 这部分记录已过时。
- `insert_ppt_image` 是**真缺陷**：桥本来就返回 `name`（`$pic.Name`，wps-com.ps1:6397），
  但工具层声明并读取的是 `imageIndex` —— **桥从不返回那个字段**，所以那句「图片索引: N」
  **永远是死代码**，而真正有用的名字被丢掉了；`replace_ppt_image` 是**按形状名定位**的，于是链不起来。

**修法**：改回读桥真的给的字段（2 行，不涉及桥与管线）。
**验证**：`形状名称: 图片 1`，且与 `wps_ppt_get_shapes` 报的名字**逐字一致**；
`contract-gaps` 34/34（+3 条：回了名字 / 名字与 get_shapes 一致 / 死代码那行已消失）。

> 结论：W4-5 不需要作为"改进项"排期 —— 它要么已经做了，要么是个可以当场修掉的字段对不上。

**原始记录（保留背景）**：

`wps_ppt_add_shape` / `wps_ppt_insert_ppt_image` **不返回**对象的 `name`（形状实际叫「矩形 1」、图片叫「图片 4」，
都是 WPS 自动命名）。所以"拿插入时的返回值与删除时的 `impact.name` 逐字比对"这条写法**不成立**；
本轮已把断言改为"必须报出**可辨认的标识**（种类/图片 + 序号）"。
若以后想让模型能可靠地引用刚建的对象，**让 add 类工具回传对象名**是个值得考虑的改进（不是缺陷）。

## W1-3 / W1-4 · 第 3 批进度（进行中）

**共用 oracle 模块：`test/lib/oracle.mjs`** —— 三条独立路径收成一个模块，后续 39 个文件照着用即可：
1. 外部库读**落盘文件**（openpyxl / python-docx / python-pptx）—— 最强
2. 裸 COM 直接问 WPS（`comGrid` / `comSheetNames` / `comParagraphs`）
3. 文件头（`pngInfo`：字节数 + 真实像素尺寸）

| # | 文件 | 状态 |
| --- | --- | --- |
| 1 | `destructive-guard.test.mjs` | ✅ 45 → **53** 条，破坏路径加了 7 条真值回读（清空/删行/删列后逐格比对、清格式不动内容、删表只删那张、批注真的没了、区块左移正确） |
| 2 | `excel-page-setup.test.mjs` | ✅ 25 → **35** 条，加 9 条"以磁盘文件为准"的断言（openpyxl 读方向/纸张/页边距/居中/打印标题/页眉页脚/缩放/标签色/可见性）—— **第一次运行就挖出并修掉 2 个真 bug（W1-5）** |
| 3 | `excel-list-object.test.mjs` | ✅ 26 → **36** 条，三个存盘点读 `xl/tables/*.xml`（建表后 A1:C5 与列名、改名+调整后 SalesData/A1:C6/样式、unlist 后文件里**真的没有表**且数据还在） |
| 4 | `word-deep.test.mjs` | ✅ 28 → **35** 条，存盘后用 python-docx 读文件（表真的在/3 行 4 列/写进去的单元格文字在；转成文本后**表真的没了**而文字还在） |
| 5 | `word-produce.test.mjs` | ✅ 20 → **27** 条，存盘后直读 docx 的 XML part（页脚/页眉里**真的有 PAGE 域**、分栏真的回到 1 栏、修订标记 ins/del 归零、批注 part 清空） |
| 6 | `excel-advanced.test.mjs` | ✅ 25 → **30** 条，裸 COM 读求解后的 A10/B10 真值（`[['25','50']]`）、chart part 里真的有新标题（含坐标轴）、清掉一组后文件里**恰好剩 1 组**迷你图 |
| 7 | `excel-missing-halves-2.test.mjs` | ✅ 31 → **38** 条，裸 COM 读每格粗体（格式刷真的刷上、清格式真的清掉）、文件里条件格式真的删干净、合并计算结果 D1=11/D2=22、SUBTOTAL 公式真的落盘 |
| 8 | `excel-contract-fixes.test.mjs` | ✅ 31 → **41** 条，补上 4 处"只断言调用没报错"的文件级核对（数据验证/超链接/边框/条件格式）—— **当场挖出并修掉 `set_border` 的静默空转（FIXES 98）** |
| 9 | `ppt-coverage.test.mjs` | ✅ 86 → **92** 条，存盘后读 pptx：**页数与工具汇报跨源对照**（不写死 —— 实测两次运行分别是 3 页和 4 页）、文件里有测试写入的内容与替换标记、导出 PNG 是真实 1280×720 |
| 10 | `excel-coverage.test.mjs` | ✅ 64 → **68** 条，导出图片必须**真尺寸**（419×116 / 962×602，不是 1x1 占位）、存盘工作簿里真的还有数据 |
| 11 | `word-common-coverage.test.mjs` | ✅ 50 → **53** 条，重存后读 docx：插入的两段文字真的在、插入的图片**真的成了内嵌图形**（`inlineShapes=2`） |
| 12 | `open-safety.test.mjs` | ✅ 21 → **23** 条，裸 COM 跨源核对：加密打开失败后 WPS 里**确实没有**半开着的文档/工作簿（`raw="0"`；带纯数字闸门，防 `Number('')===0` 假通过） |
| 13 | `close-safety.test.mjs` | ✅ 14 → **16** 条，裸 COM 跨源核对：关完之后 Excel / 演示实例里**确实没有**残留（`raw="0"`） |
| 14 | `warning-channel.test.mjs` | ✅ 12 → **14** 条，两条**前置条件**经裸 COM 交叉验证（"确实开了两个工作簿/两个文档"，`raw="2"`）—— 前置条件不成立的话整串警告断言都是白测 |
| 15 | `error-wording.test.mjs` | ✅ 28 → **31** 条，26 个"必须失败"的调用经裸 COM 核对**零副作用**（打开的东西没有变多；只比增量，不假设机器上本来没开文档） |
| 16 | `orphan-reclaim.test.mjs` | ⚠️ **文件本身已经很硬**（用操作系统级进程计数做核对，天然独立）——但顺手查出一个**数据丢失级测试缺口，见 W1-7** |
| 17 | `deprecated.test.mjs` | ✅ 7 → **10** 条，规范名不只验证"能解析"，还**真调一次**并用裸 COM 读回（`set_zoom(85)` → `ActiveWindow.Zoom === 85`）—— 别名删除的两头断风险 |
| 18 | `merged-tools.test.mjs` | ✅ 19 → **23** 条，Excel 段原来六个裸 `ok()` 换成**逐格真值**：插删行列的净效果回到原样、`hide_rows` 的 `Hidden` 属性真的是 `True`/`False` |
| 19 | `range-limits.test.mjs` | ✅ 14 → **15** 条，被拒绝的 `clean_data` 之后逐格确认**数据还在**（`[["1","2","3"],["4","5","6"]]`）—— "拒绝"不等于"没动手" |
| 20 | `warnings.test.mjs` | ✅ 7 → **8** 条，"样式失败但文字照样插入"的**实际效果**经裸 COM 确认（正文里真有 `warning probe`）—— 部分成功最容易退化成"什么都没做、警告照报" |
| 21 | `target-ambiguity.test.mjs` | ✅ 13 → **15** 条，两处**前置计数**经裸 COM 交叉验证（"确实开着两个工作簿/两个演示"，`raw=2`） |
| 22 | `find-replace.test.mjs` | ✅ 14 → **15** 条，替换结果**以磁盘 xlsx 为准**再核一遍（`[["ZZZ","beta","ZZZ",...],["gamma","ZZZ","delta",...]]`）—— 原来全部回读走插件自己的 read_range |
| 23 | `ppt-contract-fixes.test.mjs` | ✅ 61 → **64** 条。**全库裸 ok() 最多的文件（38 条）**，补三个端到端锚点：转场真的写进 slide part（count=6）、形状填充真的在（`28A745` count=2）、文本框文字真的在 |
| 24 | `excel-missing-halves.test.mjs` | ✅ 18 → **20** 条，命名区域的 set/delete 补**文件级**一对核对：`xl/workbook.xml` 里真的出现（count=1）→ 删除后**真的不在**（count=0） |
| 25 | `word-longtail.test.mjs` | ✅ 20 → **24** 条，脚注/尾注/内容控件补**文件级**核对：文字真的在 `word/footnotes.xml`、`word/endnotes.xml`，`<w:sdt>` 真的在 `document.xml` |
| 26 | `ppt-slimming.test.mjs` | ✅ 32 → **36** 条，阴影真的写成 `<a:outerShdw>`、表格真的成 `<a:tbl>`、页脚文字真的在 slide part |

### 已复核、**判定无缺口**（不硬塞断言，避免数字虚胖）

| 文件 | 为什么不需要补 |
| --- | --- |
| `cell-format` | 17 条断言全是精确数值回读（bold/italic/size/name/color/对齐的整数值） |
| `sheet-ops` | 22 条断言比对**真实表序数组**，不是"调用成功" |
| `file-ops` | 查 `existsSync`（文件真的落盘）+ 真实打开计数 |
| `encrypted-preflight` | 查时序（<10s）+ 文案 + 扩展名清单，是行为验证 |
| `alerts-gate` | 源码静态分析器（含"真桥通过闸门"一条） |
| `arg-shape-guard` | 功能性校验（错误类型必须被拒、正确的必须通过） |
| `com-host` | 直接驱动宿主：冷/热 ping、并发、**杀掉后恢复** |
| `error-contract` | 契约文档 + 拒绝行为双重核对 |
| `honest-reporting` | 本身就是"诚实汇报"语义（已保存 vs 未保存、错密码必须失败） |
| `word-lifecycle` | 工具检索排序 + 生命周期，全部真实回读 |
| `word-range-format` | **52 条断言、0 条裸 ok()**（前面轮次已加固过） |
| `confirm-dialog` | 6 条断言，监听器 + 退出码核对（FIXES 92 已修过误报） |
| `silent-catch` | 5 条断言，空 catch 账本静态核对 |
| `error-wording` / `open-safety` / `close-safety` / `warning-channel` / `target-ambiguity` | 已在前面的轮次补强 |

> 结论：第 3 批 40 个文件里，真正"只断言调用没报错"的重灾区全部处理完毕；
> 两个真 bug（FIXES 97/98）都是从这类断言后面挖出来的。剩下一处已知未施工项是 **W1-7**（未保存内容护栏零覆盖）。

### W1-7 ✅ 已施工（`test/orphan-reclaim.test.mjs` 场景 4）· "有未保存内容就不许退出实例"这条护栏

**结论：护栏有效，已在真机上钉死。** 直接驱动宿主（`createDocument` → `insertText` → `__shutdown`）：

```
PASS W1-7: the document really is unsaved   Saved="False"
PASS W1-7: shutdown does not claim to have closed the word instance   []
PASS W1-7: the instance holding unsaved work is STILL RUNNING   apps=5
```

**过程中发现的一个边界事实（已写进测试注释）**：WPS 里**刚创建的空文档 `Saved` 就是 `True`**。
所以护栏会把"新建但一个字都没写"的文档当作干净、照常退出实例 —— 这不算数据丢失（内容本来就是空的），
但解释了护栏的边界：**只有写入过内容（`Saved=False`）才受保护**。

> 这条也救了我的第一版探针：当时只 `createDocument` 就断言"实例应当仍在"，跑出 `apps=0` 一度像是
> 重大 bug；正是那句 `Saved === "False"` 的前提检查证明**是我的前提错了**，不是护栏坏了。
> —— 教训：断言"某保护应当生效"之前，先断言**被保护的状态真的存在**。

**原工单（保留背景）**：

- **代码位置**：`host/wps-actions.ps1:130-153` 的 `Close-WpsAppsStartedByUs`（FIXES 57）。
  判定逻辑：逐个实例看 `$app.Workbooks/Documents/Presentations` 里有没有 `Saved == $false`；
  **读状态抛异常时按"脏"处理**（宁可不退出）；脏 → `continue`（不 `Quit()`）。
- **测试现状**：整个 `test/` 目录里**没有任何一条**断言覆盖这条分支。
  `orphan-reclaim.test.mjs` 只在**文件头注释**里写了"only when … holds no unsaved work"，从未验证。
- **风险等级**：**数据丢失**。这条护栏回归 = 会话退出时把用户没保存的文档/工作簿一起丢掉，
  而且是静默的（用户只会发现"我的文档没了"）。
- **怎么测**（已勘明，可直接施工）：`test/host-lease.test.mjs:25-60` 的 `startHost()` 模式可以直接驱动
  `host/wps-com-host.ps1`：
  1. 起一个宿主（`WPS_OFFICE_CLIENT_PID` 指向本进程）；
  2. 发 `{"id":1,"action":"createDocument","params":{}}` → 得到一份**未保存**的 Word 文档；
  3. 发 `{"id":2,"action":"__shutdown","params":{}}`；
  4. 断言 **WPS Word 进程仍在**（护栏生效）。
- **施工时的两个坑**：
  - **收尾**必须把这台实例清掉，否则残留实例会通过单实例租约把后续测试全部挡在门外。
    不要用 `$w.Quit()`（会波及用户自己开着的文档）——用探针自己记下的 PID 精确杀。
  - 要把"我们起的实例"和"本来就开着的"分开，断言用**增量**而不是绝对值。

### W1-5 ✅ 已修（FIXES 97）· 打印缩放的两个静默丢失

由上面第 2 个文件的新 oracle 挖出：设了 `zoom=90`，工具回报「90%」，**存盘后文件里连 `scale` 都没有**。
逐项二分定位到两个独立成因：**① 设 zoom 时没清 FitToPages（互斥只实现了一半）；② `reset_page_breaks`
会把缩放一起重置**（WPS 行为，但工具不该顺手改）。两个都已修并有文件级断言兜住。

**建模块时踩到并已修掉的三个坑（写进模块注释，免得后面重踩）**：
- **ProgID 分应用**：Word=`Kwps.Application`、表格=`Ket.Application`、演示=`Kwpp.Application`。
  拿 Word 的 ProgID 连表格 → `MK_E_UNAVAILABLE`。
- **失败必须返回 `null`，不能返回空数组** —— 否则 `before=[] / after=[]` 会**静默通过**（我第一版就是这样，跑出了两条假 PASS）。
- **别用 `ConvertTo-Json` 序列化嵌套数组**：PowerShell 会把每行变成 `{value:[...],Count:N}` 这种怪形状；
  手工拼 JSON 才稳定。
- **别在一个 shell 里背靠背手跑多个 WPS 测试文件**：前一个测试的宿主还没退干净，后一个就被单实例租约挡住
  （实测：连着跑三个文件得到 52/53、24/35、23/36 的假失败；**单独跑全部通过**，整轮 `run-tests.ps1` 也全绿）。
  整轮脚本会在文件之间回收孤儿并节流 —— 要么用它，要么文件之间留出间隔。
- **读 xlsx 里的表（ListObject）别走 openpyxl 的 `ws.tables`**（本机实测遍历拿不到东西），
  直接解析 zip 里的 `xl/tables/*.xml`；而且要用 **ElementTree 而不是正则** ——
  正则里的反斜杠会被多层模板字符串吃掉，我为此白跑了一轮。

### W5-1 ✅ 已修（FIXES 96）· 租约的 PID 复用会让死宿主「看起来还活着」

**实测**：租约里 `hostPid=41592`，但系统里**没有任何** `wps-com-host.ps1` 进程 —— 那个 PID 已被无关进程复用。
`Test-ProcessAlive` 只看 PID 是否存在于进程表，于是判定"占用者还活着"，新宿主**永久拒绝接管**，
报「检测到另一个 DSH 会话正在控制 WPS」（而其实一个宿主都没有）。

**影响**：使用者只要遇到一次，就得手工删 `%USERPROFILE%\.wps-office-mcp\com-host.json` 才能继续用插件。
**已修**：租约里本来就记着宿主的 `startedUtc`，用它做**身份核对**（PID + 启动时刻 ±2s）即可分辨复用；
存活判定与接管时的 `Stop-Process` 都改用它（原实现会在号被复用时**杀掉一个无关进程**）；
拒绝文案补上自救办法（直接给出要删的租约文件路径）。
回归：`test/host-lease.test.mjs` +2 条（复用号不再拦住接管、且无关进程没被杀），20/20 通过。

> 顺带记一条操作坑：按命令行过滤进程时（`CommandLine -like '*wps-com-host.ps1*'`）**会匹配到自己这条命令**，
> 不加 `Name -eq 'powershell.exe'` 限定就会把自己的 shell 杀掉 —— 实测把 DSH 的作业运行器都搞崩了。

---

## W6 · 契约盲区（5 个工具的参数契约无任何静态验证）— ✅ 已执行（FIXES 94）

> **结果**：3 个工具真查出缺陷并已修 —— `wps_word_set_paragraph`（静默忽略 range / 非法枚举静默变左对齐 /
> 空调用假成功 / 6 个参数不可见 / 行距 0 被跳过）、`wps_excel_evaluate_formula`（Excel 错误值以裸 CVErr
> 负数当结果返回）、`wps_excel_set_print_area`（失败只说"设置失败"）。
> `text_to_columns` 与 `proofread_basic` 的坏输入未发现明确缺陷，但暴露出一条**系统性**问题（见 W6-4）。

### W6-4 ✅ 已修（FIXES 99）· 未知参数被**静默丢弃**

**修法**：前提（schema 与桥键表对齐）已由 `param-contract` 保证（A/B/C/D 四类基线均为 0），
于是在 `tool-registry.ts` 的 `validateArguments` 里把「让给桥去报」的 `continue` 换成**响亮拒绝**：
列出全部可选参数名，拼写接近还会提示「是不是想传 X」。
**验证**：工具层传 `sep` 从「静默按默认执行」变成响亮拒绝；合法参数与直通路径均不受影响；
`arg-shape-guard` 10/10（+4 条回归）；**整轮 49 个文件零误伤**。

**原始记录（保留背景）**：

**实测**：`wps_excel_text_to_columns { range:"A1:A3", sep:"," }` → **不报错**，按默认逗号执行，
回报还写着「分隔符: ","」—— 调用方以为自己传的 `sep` 生效了。

**根因**：桥只对**有键表**的 action 拒绝未知键（`unknown parameter(s) ... accepted: ...`），
而 TS handler 显式拼参数时（如 `{ range, delimiter, sheet }`）会**先把未知键丢掉**，桥根本看不到它。

**为什么不在本批修**：这是一条**跨 268 个工具**的规则，要动 `tool-registry` 的校验层；
而当前多个工具的 schema 本身就不完整（W6-1 刚补了 6 个），一刀切严格拒绝会误伤。
**建议**：先把各工具 schema 与桥的键表对齐，再加"未知参数一律响亮拒绝"，并配 regression。

## W6 原始探针记录（保留）

`param-contract` 报 UNPARSED 的 5 个：`wps_excel_text_to_columns`、`wps_excel_evaluate_formula`、
`wps_excel_set_print_area`、`wps_word_set_paragraph`、`wps_word_proofread_basic`。

**为什么值得查**：这类工具的参数**没有任何门禁看着**，"参数被静默忽略"是高发缺陷。
其中 `wps_word_set_paragraph` **在广告面上**，优先级最高。

**怎么做**：逐个构造三种调用 —— ①参数名写错（如 `lineSpacing` 写成 `line_spacing`）
②漏传可选参数 ③类型错（数字传字符串）—— 看是否**响亮失败**，还是**回 success 但没生效**。
**怎么验**：改完属性后用裸 COM 读回真值。

---

## W5 · 生命周期（留作第二批，需要多文档/多工作簿环境）

- ~~**C7 残余**~~：`transpose` 与 `copySheet` 的**目的地**解析 —— ✅ **已测试化**（本轮）。
  桥里 `Get-WorksheetByParam`（633 行）只认 `$p.sheet`，且**只在活动工作簿里解析**（`$excel.ActiveWorkbook`），
  也**没有 workbook 参数** —— 跨工作簿寻址这件事**做不到**。所以不是"没修"，而是**契约上就不支持**。
  已在 `test/target-ambiguity.test.mjs` 钉死三条可复现断言（两个簿各写标记 → 在活动簿转置）：
  ① 结果**只**落在活动簿（`["BOOK2","seed"]`）；② **另一个簿绝不被碰**（`["",""]`）；
  ③ 不指定 sheet 时照旧给出歧义警告。**从口头记录变成了证据。**
- 打开→编辑→**另存失败**→关闭 的组合；未保存就关；长动作超时后宿主状态；陈旧接管。

---

## 建议的批次划分

| 批次 | 内容 | 为什么这样切 |
| --- | --- | --- |
| ~~**第 1 批**~~ | ~~W4-1 + W6~~ **✅ 已完成（FIXES 94）** | 3 个工具修掉 6 处缺陷；新增 `contract-gaps.test.mjs` 26/26；广告面 schema 46,867 → 47,630 字节 |
| ~~**第 2 批**~~ | ~~W1-1 + W1-2（50 条弱断言）~~ **✅ 已完成（FIXES 93）** | 22 处升级、8 个文件真机全绿、断言总数不变 |
| **第 3 批** | W1-3 + W1-4（给破坏性面补真值 oracle） | 工程量大但价值最高；必须先做 W1-1/2 练过手 |
| **第 4 批** | W5 生命周期 + C7 残余测试化 | 需要多文档环境，单独排 |

> 每批的收口动作（项目既有纪律）：契约管线按 §4 顺序 → `lint` → `verify --static` →
> `param-contract` → `gen-numbers` → 定向回归 → 整轮（~15 分钟）→ 一个 FIXES 号。

## 未采信的假设（记录，避免重复挖）

- ~~`export_range_as_image` 的 JSON 污染~~：**0.6.2 不复现**（W4-2）。
- ~~`run-tests.ps1` 的 perFile 是 PowerShell 脚本缓存问题~~：实为**缺 BOM 的编码问题**，已修（FIXES 92 / P8）。
- 环境类失败（"WPS Excel not running"）**不是缺陷**：WPS 没开时的正确报错。

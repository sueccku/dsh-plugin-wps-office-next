# 场景 S2：项目验收报告定稿（Word 单文档深度）

> 面向 A3「产品扩张」唤醒规则的一份**真实用户场景卡**：谁 × 手上什么文件 × 要什么产出 × 现在手工怎么做。
> 用途：拿它当测试输入，看插件在一条完整真实任务链上能不能走通、哪里会假成功。
> 相关：[HANDOFF.md §3.1](../HANDOFF.md)（A3 唤醒规则）、[tool-roadmap.md §0](../tool-roadmap.md)（类别 B/C 边界）、[S1 场景卡](S1-excel-monthly-report.md)（Excel 域的同类走查）。

| 项 | 内容 |
| --- | --- |
| 场景编号 | **S2** |
| 域 | Word（文字） |
| 建立日期 | 2026-10-10 |
| 状态 | **场景卡已生成，等待用户批准后开始真机走查**（门禁 D6）；尚无任何执行结果 |
| 选定依据 | 用户在本轮决策 **D1 = B（Word 单文档深度）**；S1 已覆盖 Excel 域，Word 长尾（批注 / 修订 / 目录 / 页眉页脚）此前**没有真机链路证据** |

---

## 1. 一句话场景

**系统集成公司的项目助理**，手上有一份三家单位拼起来的《2026年智慧园区项目验收报告》**初稿**（Word），
要在**验收会前一晚**把它变成**能直接发甲方的定稿 + PDF**；
现在他得逐个改标题样式、手工敲目录、逐条处理批注和修订、对不上就翻表核数字、全文查找替换、再调页眉页脚，
**约 50 分钟，且最容易漏掉「修订没接受就发出去」**。

---

## 2. 用户原话（测试输入，可原样发给模型）

> 帮我把我这份《2026年智慧园区项目验收报告》初稿定稿，文件在这儿：
> `D:\dsh\a\test\.artifacts\scenario-s2\input\2026年智慧园区项目验收报告_初稿.docx`
>
> 明天上午就要交给甲方，按下面的来：
> 1. **另存为** `2026年智慧园区项目验收报告_定稿.docx`（放同一个文件夹），**初稿一个字都别动**。
> 2. 标题现在是"加粗大字"，不是真标题，帮我套上样式：文档大标题用「标题」，一级标题（一、二、三…）用「标题 1」，
>    二级（1.1、1.2…）用「标题 2」——这样后面目录才做得出来。
> 3. 正文统一：**宋体、小四、1.5 倍行距、首行缩进 2 字符**（表格里的字不用动）。
> 4. 在大标题后面、正文前面插一个**自动目录**（三级）。
> 5. 页眉写 `2026年智慧园区项目验收报告`；页脚要 `第 X 页 共 Y 页` —— 现在是手敲的"第 1 页"死字，换成真的页码。
> 6. 文档里 3 条批注都是过程记录，定稿不该留：**该采纳的采纳完，全部删掉**。
> 7. 还有 2 处修订（一插一删），**全部接受**，定稿不能带修订痕迹；顺手把"修订跟踪"**关掉**（现在是开着的）。
> 8. **数字以「验收数据表」为准**：正文里"合同金额"那句和表里对不上，按表改正文，**表里的数字一个都不许改**；
>    表头加粗、加个浅色底纹。
> 9. 全文替换：`待定` → `已确认`；旧简称 `智园` 统一成 `智慧园区` —— **注意"智园云平台"是产品名，不能改**。
> 10. 定稿最后**一共多少页、多少字**，你用工具读一下告我，别估。
> 11. 导出一份 PDF：`2026年智慧园区项目验收报告_定稿.pdf`，同目录。
> 12. 弄完别把文档开着，列一下当前还开着哪些文档给我看。

---

## 3. 手上文件：真实脏点清单

fixture 由 `test/.artifacts/scenario-s2/tools/make_fixture.py` 生成（**独立于插件**，python-docx 1.2.0 直写 +
对 `document.xml` 做原始 XML 注入以造修订痕迹），真值另存 `test/.artifacts/scenario-s2/expected.json`。

| # | 脏点 | 数量 | 为什么真实 |
| --- | --- | --- | --- |
| 1 | 各级标题都是**直接格式**（加粗+字号），段落样式全是「正文」 | 7 段 | 三方拼稿最常见的状态：看着像标题，目录做不出来 |
| 2 | **修订痕迹**：1 处插入（"（含软硬件）"）+ 1 处删除（"临时"） | 2 处 | 复核稿必然带修订 |
| 3 | **修订跟踪开关处于打开**状态 | 1 | 定稿要求关掉，否则后续改动继续留痕 |
| 4 | **批注 3 条**（作者 2 人；1 条已过期、1 条要求改简称、1 条是排版提醒） | 3 条 | 审阅意见混着"已过期"的 |
| 5 | 正文"合同金额 **1180.00** 万元"与「验收数据表」里的 **1280.00** 万元**打对不上** | 2 处 | 改稿只改了一处，另一处忘了 |
| 6 | 全文散着 `待定` | 3 处 | 未定稿标记 |
| 7 | 旧公司简称 `智园` | 5 处 = 4 处简称 + **1 处产品名「智园云平台」（必须保留）** | 陷阱：全局替换会改坏产品名 |
| 8 | 页脚是**静态文本**"第 1 页"；页眉是旧文字"内部资料" | 2 | 手敲页码，改页数就不对 |
| 9 | 表格是裸表：表头无加粗、无底纹、无框线 | 1 张（4 行 × 3 列） | 初稿表格 |
| 10 | 段落里全角空格 / 行首多余空格混排 | 若干 | 手工拼稿 |

真值（`expected.json`）：`待定`=3、`智园` 子串=5（其中 `智园云平台`=1）、合同金额=**1280.00 万元**、
被删除字 `临时`=1、插入内容 `（含软硬件）`=1、批注=3、修订=2、标题层级=标题/标题 1×3/标题 2×3、表格=1。

---

## 4. 现在手工怎么做（痛点）

1. 逐个选中大标题/各级小标题 → 套「标题」「标题 1」「标题 2」；
2. Ctrl+A 设宋体小四、1.5 倍行距、首行缩进 2 字符（还要避开表格里的字）；
3. 引用 → 目录 → 自动目录（三级）；
4. 页眉页脚：删掉旧文字，插页码域，再手敲"共 Y 页"（很多人这里做错，插的是死数字）；
5. 审阅：逐条读 3 条批注 → 按意见改 → 删除批注；2 处修订逐条接受；再关掉修订跟踪；
6. 翻到「验收数据表」核对正文金额，手改正文（不敢动表）；
7. 查找替换"待定"→"已确认"、"智园"→"智慧园区"（**最容易把"智园云平台"一起改掉**，改完还得逐个回改）；
8. 另存定稿、导出 PDF、检查有没有漏掉修订痕迹。

**耗时约 50 分钟；每次改稿都要重来一遍；换个人做，漏项各不相同**（尤其"修订到底接受没接受"）。

---

## 5. 验收判据（机器可核验）

核验走**独立通道**：`test/.artifacts/scenario-s2/tools/verify.py`（python-docx 读段落/样式 + 直读 `word/document.xml`、
`word/comments.xml`、`word/settings.xml`、页眉页脚 part + PDF 头解析），**不读插件的任何自述**。
验证器本身必须先通过 known-good / broken 自检（见 §7）。

| ID | 判据 | 档位 |
| --- | --- | --- |
| F1 | `2026年智慧园区项目验收报告_定稿.docx` 存在 | 硬 |
| F2 | 初稿 sha256 未变（"初稿一个字都别动"） | 硬 |
| F3 | `..._定稿.pdf` 存在、文件头是 `%PDF-`、页数 ≥ 1 | 软（导出链路项目已知易失手） |
| T1 | 段落样式：大标题=标题、一级=标题 1（3 段）、二级=标题 2（3 段） | 硬 |
| T2 | 正文段落：宋体（含 `w:eastAsia`）+ 小四（12pt）+ 1.5 倍行距 + 首行缩进 2 字符 | 硬 |
| T3 | 存在自动目录（TOC 域 / `docPartGallery` TOC 的 SDT 任一） | 硬 |
| H1 | 页眉含 `2026年智慧园区项目验收报告` | 硬 |
| H2 | 页脚同时含 PAGE 与 NUMPAGES 域（不是静态数字） | 硬 |
| R1 | `word/comments.xml` 批注数 = 0 | 硬 |
| R2 | `w:ins` / `w:del` 计数 = 0（定稿无修订痕迹） | 硬 |
| R3 | `settings.xml` 里修订跟踪为关闭（`w:trackChanges` 不存在或 val=0） | 硬 |
| D1 | 接受修订的结果正确：文中含 `（含软硬件）`、不含被删的 `临时` | 硬 |
| D2 | 正文合同金额句解析出的数值 == **1280.00**（口径以表为准） | 硬 |
| D3 | 表格数字未被改动（合同金额仍 1280.00，其余单元格与真值一致） | 硬 |
| D4 | 表格表头行加粗 | 硬 |
| D5 | 表格表头有底纹（`w:shd`） | 软 |
| X1 | `待定` 出现 0 次、`已确认` 出现 3 次 | 硬 |
| X2 | `智园` 子串 = 1（只剩产品名）、`智慧园区云平台` = 0 次、4 处旧简称已改 | 硬 |
| S1 | 插件回报的页数/字数都是**数字**，且验证器独立统计的字符数与之同量级（差异记录，不做相等断言） | 软 |
| O1 | 收尾无遗留打开的文档（`wps_word_get_open_documents` 为空，且 WPS 进程里没有该文档） | 硬 |
| N1 | 无假成功：每个 `success: true` 的步骤都能在文件里找到对应事实；`error`/`warnings` 逐条记录 | 硬 |

「硬」不过 = 任务没达成；「软」不过单独记录，不掩盖硬项结论。

---

## 6. 期望覆盖的能力

`wps_word_open_document` → `wps_word_get_paragraphs` → `wps_word_apply_style` → `wps_word_set_font` →
`wps_word_set_paragraph` → `wps_word_generate_toc` → `wps_word_insert_header` → `wps_word_insert_footer` /
`wps_word_insert_page_numbers` → `wps_word_get_comments` → `wps_word_delete_comment` → `wps_word_get_revisions` →
`wps_word_accept_revisions` → `wps_word_get_track_changes_status` / `wps_word_enable_track_changes` →
`wps_word_get_table_data` → `wps_word_set_table_cell` / `wps_word_replace_range` → `wps_word_set_table_format` →
`wps_word_find_replace` / `wps_word_find_in_document` → `wps_word_get_document_stats` →
`wps_common_save_as` → `wps_convert_to_pdf` → `wps_word_close_document` → `wps_word_get_open_documents`。

即：**这是 Word 长尾工具（批注/修订/目录/页眉页脚/表格外观）第一次被放进一条真实任务链里跑**。

---

## 7. 验证器自检（先证明核验认得出对错）

顺序与 S1 相同，**只有 known-good 全绿且 broken 在对应项变红，才允许让模型真机跑场景**：

```powershell
$py = C:\Users\qwer\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe
& $py test\.artifacts\scenario-s2\tools\make_fixture.py          # 造脏初稿 + expected.json
& $py test\.artifacts\scenario-s2\tools\make_known_good.py      # 造 known-good.docx / broken.docx
& $py test\.artifacts\scenario-s2\tools\verify.py test\.artifacts\scenario-s2\selfcheck\known-good.docx --pdf ...\known-good.pdf
& $py test\.artifacts\scenario-s2\tools\verify.py test\.artifacts\scenario-s2\selfcheck\broken.docx
```

- known-good：**全部硬检查通过**（exit 0）
- broken（注入：修订未接受、批注没删、金额没改、无目录、页脚静态数字、产品名被改坏）：**对应硬项变红**（exit 1）

**已知的验证器陷阱（必须先在自检里处理掉，否则会变成对插件的假指控）**：

| # | 陷阱 | 处理 |
| --- | --- | --- |
| 1 | 标题样式名：openpyxl 式的"英文名/中文名"之争——python-docx 读回可能是 `Heading 1`，WPS 界面显示「标题 1」 | 两种名字都收，并记录原始 `w:styleId` |
| 2 | 首行缩进：WPS 写的是**字符**单位 `w:firstLineChars="200"`，python-docx 的 `first_line_indent` 读不到 | 直读 XML，接受 `firstLineChars ≥ 200` 或折算后的 `firstLine` |
| 3 | 行距 1.5：可能是 `w:line="360" w:lineRule="auto"`，也可能是 `lineRule="multiple"` + 1.5 | 两种写法都归一化成倍率再判 |
| 4 | 宋体：中文字体常只写在 `w:rFonts/@w:eastAsia`，`run.font.name` 为空 | 直读 `eastAsia`/`ascii`，任一为宋体即算通过 |
| 5 | 目录识别：TOC 可能是 `w:fldSimple`、`w:instrText`，也可能是 SDT | 三种都认；`TOC \o "1-3"` 的级别写死与否不影响判定 |
| 6 | 页脚域：`第 X 页 共 Y 页` 可能是两个域，也可能是一个 `PAGE \* MERGEFORMAT` + 文本 + `NUMPAGES` | 只要 PAGE 与 NUMPAGES 指令都出现即通过 |
| 7 | 批注 part 不存在 ≠ 批注为 0（可能残留空 `comments.xml`） | 两种情况都算"无批注"，但分别记录 |

> S1 的教训是：验证器本身的缺陷会造成比判红更糟的**假指控**。本场景的 7 条陷阱在写验证器时就要堵住。

---

## 8. 已知不可判项

- 「手工 50 分钟」是场景设定，不是断言。
- 目录"好不好看"、字体在视觉上是否真的生效、PDF 与 docx 是否视觉一致：不判。
- 页数：`get_document_stats` 回报的页数是 WPS 排版结果，独立通道**无法复核**（python-docx 不分页）→ 只判"是数字"，不做相等断言。
- 批注"意见有没有被采纳"不可机器判定 → 只判**批注已删** + 正文是否按对应意见改（金额、简称两条可判，排版那条只判格式）。
- PDF 页数与 docx 页数是否一致：不判。

---

## 9. 待验证的疑点（**假设，不是结论**）

跑之前先把这些列为"测试要回答的问题"，避免事后按结果编解释：

| # | 疑点 | 为什么怀疑 |
| --- | --- | --- |
| H1 | `wps_word_accept_revisions` 在 WPS 上是否真把 `w:ins`/`w:del` 从文件里清掉（而不是只在界面上隐藏） | WPS 的 `Revisions` 集合行为与 Word 有差异，项目里已记录多处 COM 空壳 |
| H2 | `wps_word_generate_toc` 产出的是**真 TOC 域**还是静态文本 | 静态目录在验收判据里不算"自动目录" |
| H3 | 页脚"共 Y 页"能否做出来：`insert_page_numbers` 只管页码，"NUMPAGES" 可能要 `insert_footer` + 域 | 广告面里没有单独的"总页数"工具 |
| H4 | `find_replace` 无法"排除产品名" → 模型必须自己想招（哨兵替换 / 分段替换） | 这是本场景**最可能失分**的点：粗暴全局替换会把 `智园云平台` 改坏 |
| H5 | 另存为后再改，会不会把**初稿**也改掉（违反"初稿别动"） | S1 已验证过一次，但那是 Excel 通道 |
| H6 | 修订跟踪开关：`get_track_changes_status` 读回与 `enable_track_changes(false)` 是否一致 | 开关类能力历史上出现过读回失真 |

疑点被证伪也照样如实记录（说明这条链路没问题），不为了"有发现"而凑结论。

---

## 10. 执行准备（批准后才做）

1. 建 `.artifacts/scenario-s2/{input,tools,selfcheck,run,artifacts}`；
2. `make_fixture.py` 生成脏初稿 + `expected.json`，`make_known_good.py` 造两份对照产物；
3. `verify.py` 自检：known-good 全绿 / broken 对应项变红（**这一步不过就不进真机**）；
4. 真机：`wpse2e` profile（`link:D:/dsh/a`）跑 headless DSH 任务，prompt 就是 §2 原文，WPS 由测试侧自动启动（D4）；
5. 独立核验 + 交叉验证插件自述，写走查报告 `docs/scenarios/S2-真机走查报告.md`；
6. 发现缺陷按 **D3 = 能修就全修**（改 `mcp/` 代码并补回归，回归必须在修复前先红）；
7. 结果**先不提交**（D5），证据与结论先给用户看。

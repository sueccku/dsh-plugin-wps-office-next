# 80 个真实 WPS 场景综合评估与修改建议（v2）

> 日期：2026-10-19（Asia/Shanghai）　基线：v0.9.0 检出（`D:\dsh\a`）
> 数据：第一轮 20 个场景（见 [scenario-review-270-tools.md](scenario-review-270-tools.md)）＋ 第二轮 60 个新场景。
> 方法：网上调研（真实用户原话 + 可打开来源）→ 按模型可见面（95 广告工具 + `wps_help/wps_call/wps_batch`）做纸面仿真 → 对工具名/参数/桥实现做 schema 与源码核验。
> 状态：**未跑 WPS 真机、未跑仓库测试**（遵循 [AGENTS.md](../AGENTS.md) §1 的 R13）。本文是设计层评估，真机项见 §9。
> 过程产物：`.tmp-eval/r2-*.md`、`sim2-*.md`、`v2-source-audit.md`、`v2-sim-audit.md`（本机、未入库）。
> 核验：第二轮来源核验 54 条（43 confirmed / 8 incorrect / 3 unverifiable）；仿真核验 54 条（43 confirmed / 7 incorrect / 4 unverifiable）。本文只采用核验后的数字。

---

## 0. 结论摘要

### 0.1 一句话

**80 个真实场景里，没有 1 个能"顺手完成"；但真正被"工具面没做"挡住的只有 22 个，另外 17 个是平台/产品边界、33 个是"半步工具/要拼装"、8 个可行但绕。**
v1 的结论"270 作为 COM 能力底座合理、作为任务面不合理"被 60 个新场景进一步证实；同时暴露出 v1 没看到的三类问题：**Word 的"域/结构/保护"整族缺失、PPT 的"母版/主题/字体/节/导出"整族缺失、Excel 大量"半步工具"**。

### 0.2 80 条场景的判定

| 判定 | 条数 | 说明 |
| --- | ---: | --- |
| A 顺畅 | **0** | — |
| B 可行但绕 | **8** | v1 2 + 新增 6 |
| C 需拼凑 | **33** | v1 13 + 新增 20 |
| D 做不到 | **39** | 其中 **22 条是真缺口**、**17 条是平台/决策边界** |
| E 风险（叠加） | **19** | 可与 A–D 叠加，不占 80 的名额 |

把边界剔掉后，**插件能力范围内的 63 条**是：

| 范围 | A | B | C | D |
| --- | ---: | ---: | ---: | ---: |
| 插件范围内（63） | 0 | 8 | 33 | **22** |
| 平台/决策边界（17） | 0 | 0 | 0 | 17 |
| 合计（80） | 0 | 8 | 33 | 39 |

- **22 条真缺口**：Excel 7、Word 9、PPT 4、v1 2（详见 §4）。
- **17 条边界**：PDF 4、云/手机/AI/OLE/密码/水印/文档属性/媒体/SmartArt/打印/批量文件等（详见 §4.3）。
- **33 条 C**：不是"没工具"，而是"有工具但只做了一半、或要 3–8 倍调用拼"。
- **8 条 B**：在现有工具面下可做，但要绕/要 help/要验证。

### 0.3 成本画像

- 第二轮 60 条：**约 555 次顶层调用、118 次 `wps_help`、188 次隐藏工具执行、422 个 `wps_batch` 子调用**；抽样 417 个 `(tool,args)` **无 schema 违规**。
- 平均 **9.25 次顶层调用/条、1.97 次 help/条**（轻量仿真口径，不能与 v1 的 19.6/5.3 直接比较；v1 是完整调用链，v2 是 5–12 步摘要）。
- 隐藏工具仍是固定税：60 条里用到的隐藏工具 44 个；很多场景仍是"先 help 再 call"两步。
- 批量任务线性放大：XA9 300 个文件、XA11 50 份、WD10 100 人 ≈ 500 次调用——没有批量编排时"工具再对也不可用"。

### 0.4 六个关键发现（相对 v1 的增量）

1. **"半步工具"比"完全没工具"更危险**：有锁无"允许编辑区域/隐藏公式"（XL12）、有刷新无"列出/改源/断开"（XL15）、有单列排序无多键（XL17）、有 `set_hyperlink` 无读取（XL18）、有保护无"允许编辑区域"、有母版工具无"新增版式"（PT1）。模型会以为能闭环，实际卡在最后一步。
2. **Word 的"域"是整族缺失，不是单个功能**：题注（SEQ）、交叉引用、图表目录（TOC \c）、表格公式（=SUM(ABOVE)）、页码/总页数都依赖"插域 + 更新域"；270 里只有 `insert_page_numbers` 能造 PAGE/NUMPAGES，**没有通用 `get_fields/update_fields`** → WD5/6/7/13 四条同卡一处（v2 核验：4 条）。
3. **"文档内位置"是 v1 完全没覆盖的新瓶颈**：`insert_section_break`/`insert_image`/`add_content_control`/`insert_cross_reference`/`insert_text(cursor)` 都"插到当前选区"，唯一能移动选区的竟是 `find_replace(selectFound:true)` 的副产物 → WD2/5/10/11 四条（v2 核验：4 条）。
4. **PPT 缺的是"结构/交付链路"，不是"建页"**：母版新版式/占位符锁定、主题色/主题字体编辑、字体回读/嵌入、图片压缩、节/隐藏页、备注页/隐藏页/缩放的导出选项——PT1/2/5/6/9/10 六条（v2 核验：6 条）。
5. **`wps_help` 出现"同名不同义"误导**：用户说"题注"，help 命中 `apply_style`（因为样式名叫"题注"）；"保护"只命中 Excel 保护；"公式/排序/边框"全命中 Excel 工具。**这比 0 命中更危险**——模型会拿一个"看起来对"的错工具去执行。
6. **边界必须显式声明**：PDF 是"零工具"（只有单向 `convert_to_pdf`）、手机端/AI/OLE/云文档/密码解密/水印/文档属性/媒体/SmartArt 都不在桌面 COM 插件面内。17 条边界里，很多"D"是模型先 `wps_execute_method` 试 3–4 次才确认的——一句 `wps_status.scope` 就能省掉。

### 0.5 修改建议（详见 §6）

按"影响面 × 修复成本"排序，v2 建议做 **13 件事**：

**P0（不改 270 的规模）**
1. `wps_help` 2.0：去 60 行硬截断/加分页；补 20 个新词；**同名不同义消歧**；场景→工具/配方入口。
2. 安全默认值 + 影响面：`open_*` 默认只读、`close_*` 默认不保存、依赖 batch 默认停；全局改动给 `dryRun/impact` 与"备份副本"配方。
3. 契约补齐：桥已有 schema 未暴露的键（10+）；`read_range` 分页与 50k/200k 预算；`get_document_text` 假参数；`wps_batch` 2,000 字截断；`save_as.createDirs`。
4. 80 场景技能配方 + 边界声明（PDF/手机/AI/OLE/水印/文档属性/媒体/打印/密码/宏）。

**P1（能力族闭环；遵守 D1/D2/D3/C 类）**
5. Word「域/编号」族：`get_fields`/`update_fields` + 题注/图表目录/表格公式/交叉引用增强（4 条 D）。
6. Word「结构/保护/审阅/合并」族：选区定位、页眉页脚高级、protect/restrict、文档比较、按作者修订、内容控件参数、邮件合并标签/图片/一对多（约 8 条）。
7. PPT「结构/主题/字体/图片/节/导出」族：新版式/占位符、主题色/字体、字体回读/嵌入、图片压缩、隐藏页/节、备注页/隐藏页/缩放导出（约 8 条）。
8. Excel「半步工具闭环」：允许编辑区域/隐藏公式、外链列出/改源/断开、视图回读、超链接读改、多键排序/高级筛选、分页符/打印区域读（约 6 条）。
9. Excel「高级分析/导入」：图标集、组合图/次轴/趋势线、透视计算字段/日期分组、图表列表/回读、CSV 导入参数、固定宽度分列、公式填充/错误扫描（约 7 条）。
10. 跨应用一次成型：暴露 `word_insert_table.data`、新增 `ppt_set_table_data`、区域→Word/PPT 表格（2 条 C + v1 表格逐格成本）。
11. 备注/讲义门面：`get_slide_notes{all}` + `hasNotes` + 导出备注到 Word（XA4 从 B 到更顺）。

**P2（产品/治理）**
12. 文件级批量：不把通用文件系统工具塞进 270（Harness 已有）；只补 `save_as.createDirs`，批量转换/套模板写成 `wps_batch` 配方 + 失败清单。
13. 80 场景基准 + 广告面换血：用 80 条的真实轨迹做 R14 证据；只在 ≤100/60,000 内换血；不无证据扩容。

---

## 1. 方法与数据

### 1.1 数据规模

| 轮次 | 场景数 | 覆盖 | 仿真方式 | 来源核验 |
| --- | ---: | --- | --- | --- |
| v1 | 20 | Excel 8 / Word 6 / PPT 4 / 跨应用 2 | 完整调用链（391 顶层 / 105 help） | 16 主来源全部打开 |
| v2 | 60 | Excel 20 / Word 15 / PPT 10 / 跨应用+文件+边界 15 | 轻量仿真（555 顶层 / 118 help / 0 schema 违规） | 26 URL 抽查，26/26 打开 |

**60 条新增不是 60 个完全独立任务**（核验结论）：

- 与 v1 **完全重复 0 条、近似 22 条、不同 38 条**；其中 4 条重叠偏重（PT1/PT2/PT4≈P1/P4，XL12≈E8）。
- 内部还有 1 对同题：XL14 与 XA15（十万行卡顿）。
- 按任务簇粗算约 **75 个独立任务**；汇总时按"同簇去重"看趋势，不把 80 当 80 个独立证据。
- 证据等级合计 **A 52 / B 8 / C 0**，但六个调研文件对 A 的定义不一致（PT/XA 要求真实用户帖，XL/WD 允许教程/官方页），且存在跨场景来源复用；**52 个 A 不能当 52 个独立强证据**。核验未发现编造引文/标题/热度。

### 1.2 仿真口径

- 模型起点：95 个广告工具 + `wps_status/wps_help/wps_call/wps_batch`；隐藏工具先 `wps_help` 再 `wps_call`/`wps_batch`。
- 每条写：用户原话 → 起点 → 5–12 步流程（工具名 + 关键参数，隐藏工具标 `[隐]` + help 词）→ 计数 → 摩擦 → A–E → 建议。
- 判定：A 顺畅 / B 可行但绕 / C 需拼凑 / D 做不到 / E 风险（叠加）。
- 工具名/参数对 `spec/tool-definitions.json` 校验；判"做不到"时读 `mcp/scripts/wps-com.ps1` / `host/wps-actions.ps1`。
- **未跑 WPS 真机**：COM 赋值、弹框、性能、WPS 版本差异都可能改变个别判定；§9 列了 18 项待真机。

### 1.3 核验与修正（只列影响结论的）

1. **WD13–WD15 在 sim2-c 与 sim2-d 重复**：合并时 sim2-c 只取 WD1–WD12，WD13–WD15 以 sim2-d 为准；否则会得 63 条。合并后可采信：60 条 = A0/B6/C19/D35/E9。
2. **PT6 不是"裁剪/压缩都没有"**：`wps_ppt_set_image_style` 已支持 crop*，`replace_ppt_image` 支持换图；缺的是压缩/批量/体积回读。**PT6 从 D 降为 C**（本文按 C 计）。
3. **XL7 的 D 附条件**：CSV 能被 `open_workbook` 打开，但编码/前导零/列格式不可控；若真机 OpenText/QueryTable 可兜底，可降 C。
4. **WD8 的 COM 支持性未验证**：WPS 是否有 `CompareDocuments` 未在真机确认；工具面确实 0 个比较工具，故仍判 D（缺工具），但"补桥即可"待真机。
5. **跨应用 4 条建议（XA1–XA4）不是"必须补的桥能力"**：按 C 类决策，区域→Word/PPT、Word 大纲→PPT、PPT 备注→Word 应走技能配方/产品决策；桥只补"一次写入 data"这类原子参数。
6. **`wps_execute_method` 不是任意 COM**：只转发桥 action 名，遇到没有 action 的能力一律 `Unknown action`；所以 D 场景没有"裸 COM 兜底"。
7. **help 命中数有 10 处小错**（如"隐藏幻灯片 0"实为 58、"合并 0"实为 10），但结论不变：命中全是噪声/反向工具；本文只引用定性结论。
8. **文件系统/批量编排属于"产品决策"**：核验把 XA9/XA10/XA11 标为"新桥能力候选，需产品决策"，不算 270 设计失败。

---

## 2. 80 场景结果（按域）

| 域 | v1 | v2 新增 | 新增 B/C/D | 新增的主缺口 |
| --- | ---: | ---: | --- | --- |
| Excel | 8 | 20 | B5 / C7 / D8 | 公式填充/错误诊断、图标集、CSV 导入、透视计算字段、组合图次轴、允许编辑区域、外链、大文件诊断、多键排序/高级筛选 |
| Word | 6 | 15 | B0 / C4 / D11 | 域/题注/交叉引用/图表目录、奇偶页眉/页眉横线、文档比较、标签/图片合并、内容控件、按作者修订、表格公式/排序、限制编辑 |
| PPT | 4 | 10 | B0 / C4 / D6 | 母版版式/占位符、主题色/字体、字体嵌入/回读、图片压缩、节/隐藏页、导出选项、备注 |
| 跨应用/文件/边界 | 2 | 15 | B1 / C5 / D9 | PDF 合并/拆分/转 Office/水印（边界）、批量转换/改名/套模板、OLE、手机、AI、大文件/长文 |

> 逐条 60 场景见附录 A；v1 的 20 条见 [v1 报告](scenario-review-270-tools.md)。

**读法**：
- Excel 新增里 B 最多（5/20）：公式/文本函数能救；但 D 也最多（8）：高级分析/治理类"半步工具"集中。
- Word 新增 D 11/15，且 9 条是真缺口——这是本轮最值得投入的方向。
- PPT 新增 D 6/10，其中 2 条是 D3 边界（SmartArt/媒体），4 条是真缺口（母版/主题/字体/节/导出）；另 4 条 C（PT3/4/6/10）。
- 跨应用/文件/边界 9 条 D 里，大部分是平台边界；真正要改的是批量编排/备注门面。

---

## 3. 判定账本：22 真缺口 / 33 C / 17 边界

### 3.1 22 条真缺口（插件能力范围内、补了就能从 D 变 C/B）

| 域 | 条数 | 场景 |
| --- | ---: | --- |
| Word | 9 | WD3 奇偶页/首页不同页眉、WD4 页眉横线、WD5 题注/编号、WD6 交叉引用、WD7 图表目录、WD8 文档比较、WD9 邮件合并标签/信封、WD13 Word 表格公式/排序、WD15 限制编辑/保护 |
| Excel | 7 | XL5 图标集、XL7 CSV 导入参数、XL9 透视计算字段/日期分组、XL10 组合图/次轴/趋势线、XL12 允许编辑区域/隐藏公式、XL14 大文件诊断、XL15 外链列出/改源/断开 |
| PPT | 4 | PT1 母版版式/占位符、PT2 主题色/主题字体、PT5 字体嵌入、PT9 节/隐藏页/导出选项 |
| v1 | 2 | E2 公式条件格式、E7 图表日期轴/空值 |

> XL7、WD8、PT5、PT1/PT2 的"补桥即可"还需真机确认 COM 能力；不能把"计划"当"已支持"。

### 3.2 33 条 C：半步工具 / 要拼装

| 类型 | 代表 | 典型"只做了一半" |
| --- | --- | --- |
| Excel 半步 | XL3/XL6/XL16/XL17/XL18/XL19/XL20 | 有单列排序无多键；有批注无超链接读；有合并计算无跨簿按列对齐；有 refresh_links 无断开 |
| Word 半步 | WD2/10/11/12 | 有分节无横向"应用于本节"；有 content control 无下拉/日期；有修订无按作者筛 |
| PPT 半步 | PT3/PT4/PT6/PT10 | 有尺寸无 scaleContent；有 unify_font 无"只换字体"；有裁剪无压缩/体积回读；有备注无导出 |
| 跨应用 | XA1/2/3/11/15 | 有 set_table_cell 无 data；有 add_slide 无大纲导入；批量套模板无循环 |
| v1 | E1/E3/E5/E6/W1/W2/W3/W5/W6/P1–P4 | 见 v1 §2 |

### 3.3 17 条边界（明确不做 / 平台不在范围）

| 类型 | 场景 | 处置 |
| --- | --- | --- |
| 打印/打印预览 | F1、WD9 的打印部分 | 已决策不做（FIXES 1169-1171）；设置好 + 导 PDF 由人打印 |
| PDF 合并/拆分/转 Office/水印 | XA5–XA8 | 插件无 PDF 宿主；不补，写边界 + 外部工具 |
| 密码遗忘/解密 | XL13 | WPS 不支持；明确拒绝 |
| 水印 | WD1 | WPS 页眉 Shapes 空壳；明确拒绝 |
| 文档属性/PDF 元数据 | WD14 | WPS COM 坏壳/365；明确拒绝 |
| 媒体/SmartArt | PT7/PT8 | D3 已删；明确拒绝 |
| 云文档协作 | E8 | 本地桌面插件不覆盖；明确边界 |
| 活链接/OLE | X1、XA12 | WPS 放映态不刷新；本地插件无 OLE；写替代配方 |
| 手机端 | XA13 | 桌面 COM 不覆盖；明确边界 |
| WPS AI | XA14 | 不在插件范围；明确边界 |
| 文件系统/批量风控 | XA9、XA10 | Harness 有文件工具；插件只补 `createDirs` 与配方；批量转换受会员/风控 |
| 宏/VBA | （调研候选，未入 60） | 明确不做 |

> 这 17 条不应计入"必须补的工具"；正确做法是**在 `wps_status`/`wps_help`/SKILL 顶部用一行声明范围**，并在工具描述里写"不支持 + 替代路径"。核验确认：v2 的 5 份仿真没有把边界误列为"必须补"（除 4 条跨应用建议被纠正）。

---

## 4. v1 结论的验证与修正

### 4.1 被证实的

- 270 作为 COM 底座覆盖深、错误诚实、分层克制；作为任务面 0 个 A、平均调用高。
- 隐藏工具发现成本、`wps_help` 60 行截断、`schema 比桥窄`、三套目标模型、危险默认值、验证面缺失。
- PPT 广告面偏"建页/写字"、缺"改已有稿"；v2 把同一结论扩展到 Word（域/保护/比较）与 Excel（治理/打印）。

### 4.2 被修正/加强的

| v1 说法 | v2 修正 |
| --- | --- |
| "工具面缺任务级组合" | 组合不是唯一解；**33 条 C 里大量是"半步工具"**（只差一个参数/一个只读工具），比新增复合工具更便宜 |
| "PPT 最不合理" | Word 新增 D 11/15，且 9 条真缺口；Word 是更大的能力族缺口（受 D2 支持） |
| "help 0 命中" | 新增"同名不同义"误导（题注→apply_style；保护→Excel；公式/排序/边框→Excel），需要消歧而不只是同义词 |
| "5 个 D" | 80 条原始 D 39；**17 条是边界**，真缺口 22；不应把 39 都当设计失败 |
| "Word 长尾全要"（D2） | v2 证明这条决策是对的：Word 的域/保护/比较/内容控件是真实高频，不是长尾洁癖 |
| "C 类不进桥" | 跨应用 4 条应走配方（XA1–XA4），但 **Word/PPT 的"一次写入 data"是原子能力**，不属于跨应用工作流，可以补参数 |

### 4.3 数字修正

- 60 新增：A0/B6/C19/D35/E9（sim2-c 只取 WD1–12；PT6 降 C 后为 A0/B6/C20/D34/E9）。
- 合并后 80：A0/B8/C33/D39/E19（PT6 降 C 后）。
- 合并调用：顶层 **555**、help **118**、隐藏 **188**、batch 子 **422**。
- 60 新增里 **22 条与 v1 同簇近似**、38 条不同；4 条重叠偏重；约 75 个独立任务簇。

---

## 5. 能力缺口 Top 12（按影响条数/修复价值）

| # | 缺口家族 | 影响条数 | 场景 | 性质 | 建议 |
| ---: | --- | ---: | --- | --- | --- |
| 1 | Excel 公式/计算/查找语义 | 5 | XL1/2/3/4/9 | 半步 + 真缺口 | 拆 `set_formula` 填充语义；补错误值扫描/诊断；`copy_range` 配方 |
| 2 | Word 域/题注/交叉引用/图表目录/表格公式 | 4 | WD5/6/7/13 | 真缺口 | `get_fields/update_fields` + 4 个上层工具 |
| 3 | 光标/选区定位 | 4 | WD2/5/10/11 | 真缺口 | `wps_word_select_range`，insert_* 接受 range/paragraph |
| 4 | 文件系统/多文件批量编排 | 4 | XL16、XA9/10/11 | 产品决策 | 不塞 270；Harness 文件工具 + `save_as.createDirs` + 配方 |
| 5 | Excel 导入/分列/文本 | 3 | XL4/6/7 | 真缺口/半步 | CSV 编码/列格式；固定宽度分列；智能填充/正则 |
| 6 | Excel 高级图表/透视/图标集 | 3 | XL5/9/10 | 真缺口 | icon set、组合图/次轴/趋势线、计算字段/日期分组（先做 COM 探针） |
| 7 | PPT 导出/显示选项 | 3 | PT3/9/10 | 真缺口 | 隐藏页/备注页/缩放选项；`convert_to_pdf` 扩参 |
| 8 | Word 页眉页脚高级/分节 | 3 | WD2/3/4 | 真缺口 | firstPage/oddEven/linkToPrevious/清空/边框；分节横向 |
| 9 | PPT 母版/主题/字体/图片/节 | 5 | PT1/2/5/6/9 | 真缺口 | 新版式、主题色/字体、字体回读/嵌入、图片压缩、隐藏页/节 |
| 10 | Excel 半步治理（保护/外链/视图/超链接/排序/打印） | 6 | XL11/12/15/17/18/20 | 半步 | 补只读回读 + AllowEditRanges/FormulaHidden + LinkSources/Disconnect + 多键排序 + 分页符 |
| 11 | 跨应用一次写入 | 2 | XA1/2 | 半步 | 暴露 `word_insert_table.data`；`ppt_insert_table.data`/`set_table_data` |
| 12 | 无撤销/事务/影响面 | 9（E） | XL3、WD2/8/12、PT3、XA9/11/12/15 | 安全横切 | `dryRun/impact`、备份副本、batch 停/清单 |

> 另：`wps_help` 的同名不同义/60 行截断/同义词缺失是横切所有 80 条的"发现税"，单列在 P0。

---

## 6. 修改建议 v2（13 项）

> 约束：D1 广告面 ≤100 工具 / ≤60,000 字节（当前 95 / 55,510，余 5 位 / 4,490 字节）；D2 Word 长尾全要；D3 PPT 已删媒体/SmartArt/讲义/3D/美化族；C 类跨应用工作流/美化不进桥；打印已决策不做；R14 隐藏→广告需真机轨迹证据。
> 新增工具**默认隐藏**，不占广告位；只有通过 R14 取证的高频工具才换入广告面。

### P0：先修"发现面 / 安全 / 契约"（1–3 天，不动 270 规模）

**P0-1 `wps_help` 2.0（横切全部 80 条）**
- 去掉 60 行硬截断，或加 `offset/limit`；隐藏工具优先排序；`truncated` 时给下一步提示。
- 补 20 个新词同义词：题注/交叉引用/图表目录/域/更新域/奇偶页/页眉横线/文档比较/标签/信封/席卡/内容控件/限制编辑/字体嵌入/图片压缩/节/隐藏页/备注页/母版/主题色。
- **同名不同义消歧**：同一词命中多个域时，按"当前活动应用 + 用户原词"排序，并标注"这是 Excel 工具，不是 Word"。
- 场景入口：`wps_help {query:"批量合同"}` 返回 `mail_merge` + 配方名 + 前置条件。
- 验收：60 条新场景的 help 0 命中/误命中从 ~15 次降到 ≤5；平均 help ≤1.5；所有隐藏工具可通过 app/query 找到。

**P0-2 安全默认值 + 影响面**
- `open_*` 默认只读；`close_*` 默认 `save:false`；依赖型 `wps_batch` 默认 `stopOnError:true`。
- 全局属性改动（尺寸/字体/主题/保护）返回"影响页数/形状数/区域"；提供 `dryRun`/`preview` 或"先 `save_as` 副本"的强制配方。
- 验收：19 条 E 风险降到 ≤5；PT3/WD2/WD8/WD12/XL3 的"无撤销"有明确替代路径。

**P0-3 契约补齐（只补已有/桥已实现，不新增能力）**
- 10 个高影响桥键：条件格式 type/值/色阶、数据验证提示、`set_cell_style` 自定义、`word_insert_table.data`、`find_replace.range/sheet`、PPT `applyToAll`/`presentationName`、Excel `sheet`、`insert_header/footer` 的 alignment/includePageNumber。
- 修 3 处描述-实现不符（`set_background_gradient`、`set_master_background`、`add_master_element`）。
- `get_document_text` 实现 `start/end` 或删掉假参数并写明"只回前 1 万字符"；`read_range` 加 `startRow/maxRows/truncated` 并写清 50,000 格预算；`wps_batch` 每项 2,000 字截断改为可配置/返回 `truncated`；`save_as` 加 `createDirs`。
- 验收：417 个抽样调用零 schema 违规；10 个高影响示例全部可传；无"按描述调用必失败"。

**P0-4 80 场景配方 + 边界声明**
- 把 80 条压成约 30 个配方（按 §8 的分组），每个：起手体检 → 决策树 → 最小调用序列 → 验收清单 → 失败回退 → 禁止捷径。
- 每应用 SKILL.md 顶部加"明确不支持 + 替代路径"（PDF/手机/AI/OLE/水印/文档属性/媒体/SmartArt/打印/密码/宏/云协作）。
- 验收：新模型按配方完成 B/C 场景不超过配方步数 1.5×；D 场景先输出"不支持根因 + 替代"，不再 `execute_method` 试探。

### P1：能力族闭环（下一版本，1–3 周；新增工具默认隐藏）

**P1-1 Word「域/编号」族（4 条 D）**
- `wps_word_get_fields` / `wps_word_update_fields{scope:"all"|"toc"|"seq"|"ref"}`。
- `wps_word_insert_caption{label, position, includeChapter}`；`wps_word_insert_table_of_figures{label, leader, hyperlinks}`；`wps_word_set_table_formula{cell, formula}`；增强 `insert_cross_reference`（目标选取/更新）。
- 先做 COM 可行性探针（`Fields.Add`/`TablesOfFigures.Add`/`Table.Cell.Formula`）；做不到就在描述里写"WPS 不支持"。
- 验收：WD5/6/7/13 从 D 到 B/C；题注/图目录/交叉引用可回读。

**P1-2 Word「结构/保护/审阅/合并」族（约 8 条）**
- `wps_word_select_range{start,end}|{paragraphIndex}|{findText}`；让 `insert_section_break/insert_image/add_content_control/insert_cross_reference/insert_text` 接受 range/paragraph。
- `insert_header/footer` 增加 `firstPageDifferent/oddEvenDifferent/linkToPrevious/clear`；允许 `text:""` 清空。
- `wps_word_protect{mode:"readOnly"|"comments"|"trackedChanges"|"forms", password, editableRanges}` + `unprotect` + `get_protection`（对齐 Excel 保护族）。
- `wps_word_compare_documents{base, revised, output}`（先证 COM `CompareDocuments`）；`accept_revisions`/`delete_comment` 加 `author`/`indexes` 过滤。
- `add_content_control` 增加 `dropdownList/dateFormat/placeholder` 参数；`mail_merge` 增加 `labelLayout`/`imageField`/`oneToMany`/xlsx 直连。
- 验收：WD2/3/4/8/9/10/11/12/15 从 D/C 到 B/C；文档比较/限制编辑/表格公式有真机测试。

**P1-3 PPT「结构/主题/字体/图片/节/导出」族（约 8 条）**
- `wps_ppt_add_master_layout`/`set_placeholder_lock`（或降级"锁形状"）；`set_theme_colors`/`set_theme_fonts`。
- `wps_ppt_get_fonts`（回读）+ `unify_font{sourceFont, preserveSize}`（只换字体）+ `embed_fonts`。
- `set_image_style` 已有 crop，补 `compress{quality,targetKB}` + 体积回读；`replace_ppt_image` 保留裁剪/效果/超链接（待真机）。
- `wps_ppt_set_slide_hidden`/`get_hidden_slides`；`set_slide_size{scaleContent:"ensureFit"|"maximize"}` + 返回受影响页数。
- `convert_to_pdf{includeHiddenSlides, exportNotes, scaleToFit}` 或独立 `export_notes({to:"word"|"text"})`；`get_slide_notes{all,hasNotes}`。
- 验收：PT1/2/3/4/5/6/9/10 从 D/C 到 B/C；导出选项可回读；字体嵌入有成功/失败分级。

**P1-4 Excel「半步工具闭环」（约 6 条）**
- `lock_cells` 加 `hidden`（FormulaHidden）；新增 `allow_edit_range`/`get_protection`；`protect_sheet` 加 `allowEditRanges`。
- `refresh_links` 返回 `[{source,status}]`，支持 `disconnect:true`/`newSource`；新增 `get_links`。
- `get_sheet_settings` 扩 `protected/frozen/outlineLevels/hiddenColumns/vPageBreaks/pageCount`；新增 `get_hyperlinks`（读改）。
- `sort_range` 支持多键/自定义序列/`hasHeader`；`auto_filter` 支持高级筛选（条件区域/复制到新位置）。
- `insert_page_break`；`get_sheet_settings` 补纵向分页符；`set_print_area` 读回。
- 验收：XL11/12/15/17/18/20 从 D/C 到 B/C；排序不把表头排进数据（真机确认）。

**P1-5 Excel「高级分析/导入」（约 7 条）**
- `set_conditional_format` 补 icon set（桥需新增 `FormatConditions.AddIconSetCondition`）、色阶/数据条（桥已有、schema 无）。
- `update_chart` 补 `axis:{categoryType,numberFormat,displayBlanksAs}`、`combo`/`secondaryAxis`/`trendline`；新增 `get_charts`（列表/回读）。
- `update_pivot_table` 补 `calculatedFields`/`dateGrouping`；WPS 切片器仍不支持（`Slicers.Count=0`），明确声明 + 替代配方。
- `open_workbook` 加 `encoding/columnFormats/startRow`；`text_to_columns` 加 `mode:"fixed"+widths`；`set_formula` 加 `fill:"broadcast"|"relative"`（或文档写死"多格=广播，填充用单格+copy_range"）。
- 新增 `scan_errors{range}`（只读返回地址/类型/是否公式）；`diagnose_formula` 扩全表。
- 验收：XL1/5/6/7/9/10/14 从 D/C 到 B/C；CSV 导入/图标集/组合图有真机结果；做不到的写边界。

**P1-6 跨应用「一次写入」原子能力（2 条 C + v1 表格成本）**
- 暴露 `word_insert_table.data`（桥已读、schema 未收）；给 `ppt_insert_table`/新增 `ppt_set_table_data` 加 `data[][]`，替代 N×M 次 `set_table_cell`。
- `insert_ppt_chart` 支持 `data` 一次写入（现在要 `set_ppt_chart_data` 二次调用）。
- **不做**区域→Word/PPT 的复合"搬数据"工具（C 类）；写技能配方。Word 大纲→PPT、PPT 备注→Word 也走配方。
- 验收：XA1/XA2 的 5×4 表格从 20 次单元格调用降到 1–2 次；v1 T3（5 页表格 72 次）降到 ≤10 次。

**P1-7 备注/讲义门面（XA4）**
- `get_slide_notes{slideIndex:"all"}` 返回 `hasNotes`；`export_notes({to:"word"|"text"})`；修 `NotesPage.Shapes.Item(2)` 的脆弱读取。
- 验收：40 页备注→Word 从 41 次调用降到 1–2 次，空备注可识别。

### P2：产品/治理

**P2-1 文件级批量：不塞 270**
- 不新增通用 `file_list/rename/move`（Harness 已有）；只补 `save_as.createDirs` 与"批量导出 PDF/批量套模板"配方（Harness 列目录 + `wps_batch` + 失败清单）。
- 若产品要做独立版，再单独立项，不混进 270 的工具面。
- 验收：XA9/10/11 有可执行配方；不再出现"`save_as` 不建目录"的失败。

**P2-2 80 场景基准 + 广告面换血**
- 在 `test/skilleval/` 增加 80 条 fixture/prompt/确定性验收；指标：成功率/顶层调用/help/隐藏发现/源文件改动/无依据成功。
- 广告面只在 ≤100/60,000 内换血；候选晋升：Word 域/保护/比较工具、PPT 结构/导出工具、Excel 图标集/外链/允许编辑区域；换出 80 场景零使用且非闭环必需的工具。任何扩容需 R14 真机证据。
- 验收：每版能报"80 场景 B/C/D 变化"；不出现"改了但指标没动"。

---

## 7. 决策对齐与不做清单

| 不做 | 原因 |
| --- | --- |
| 加 PDF 合并/拆分/转 Office/水印工具 | 插件只有 Writer/Spreadsheet/Presentation COM，无 PDF 宿主；走边界声明 + 外部工具 |
| 加打印/打印预览 | 已决策不做（FIXES 1169-1171） |
| 加媒体/SmartArt/讲义/3D | D3 已删；PT7/PT8 明确拒绝 |
| 加水印/文档属性工具 | WPS COM 自身坏壳；明确拒绝 |
| 加密码解密 | WPS 不支持；明确拒绝 |
| 把通用文件系统工具塞进 270 | Harness 已有；避免工具面膨胀；只补 `createDirs` |
| 为 33 条 C 都加"任务级复合工具" | C 类工作流走配方；桥只补原子能力/参数 |
| 无证据扩广告面 | D1 100/60,000；R14 要真机轨迹 |
| 为减少数量删桥 action | 隐藏工具不占上下文；删 action 会留台账（FIXES 48 教训） |

---

## 8. 80 场景 → 约 30 个配方（B 路径）

| 配方组 | 覆盖场景 | 关键工具/前置 |
| --- | --- | --- |
| Excel 清洗/对齐 | XL1/3/4/6/7/8/16 | `read_range`/`clean_data`/`text_to_columns`/`remove_duplicates`/`sort_range` |
| Excel 公式/查找 | XL1/2/3/4/9 | `set_formula`（单格+copy_range）、`diagnose_formula`、`get_formula` |
| Excel 透视/图表 | XL5/9/10 | `get_pivot_tables`/`create_pivot_table`/`update_chart`/`create_chart` |
| Excel 治理/打印 | XL11/12/14/15/17/18/19/20 | `get_sheet_settings`/`lock_cells`/`refresh_links`/`sort_range`/`reset_page_breaks` |
| Word 长文/域 | W1/2、WD5/6/7/13 | `get_paragraphs`/`apply_style`/`generate_toc`/`get_fields`/`update_fields` |
| Word 版式/节 | WD2/3/4 | `insert_section_break`/`set_page_setup`/`insert_header/footer` |
| Word 审阅/保护/比较 | W4、WD8/12/15 | `get_revisions`/`accept_revisions`/`compare_documents`/`protect` |
| Word 合并/表单 | W3、WD9/10/11 | `mail_merge`（扩标签/图片/一对多）/`add_content_control` |
| PPT 模板/主题 | P1、PT1/2 | `set_active_target`/`set_slide_theme`/`set_theme_colors`/`set_theme_fonts` |
| PPT 字体/图片 | P4、PT4/5/6 | `unify_font`/`get_fonts`/`embed_fonts`/`set_image_style`/`replace_ppt_image` |
| PPT 动画/导出 | P3、PT3/9/10 | `add_animation`/`set_slide_size`/`set_slide_hidden`/`convert_to_pdf` |
| 跨应用/文件 | X1/XA1–4/9/10/11 | 区域→Word/PPT 表格；PPT 备注→Word；批量导出配方 |
| 边界/替代 | 17 条边界 | 明确拒绝 + 替代路径（PDF 外部工具、Harness 文件工具、人工打印） |

> 配方放在技能 `recipes/`，不是桥工具；每个配方都要有"起手体检/决策树/验收/回退"四段。

---

## 9. 下一步验证（真机 + 基准）

**必须真机验证的 18 项**（合并 v1/v2；建议先抽 8 项）：

1. `set_formula` 多格广播 + `copy_range` 相对引用（5 条依赖）。
2. `read_range` 50,000 格预算与报错；`find_in_sheet` 200,000 格。
3. `get_document_text` 的 `start/end` 是否所有入口都忽略（源码确认忽略）。
4. `open_workbook` 的 CSV 编码/前导零/日期列；OpenText/QueryTable 能否兜底。
5. `save_as` 不建目录的行为；`createDirs` 方案。
6. `sort_range` 是否把表头排进数据；多键排序的稳定性。
7. `wps_word_compare_documents` 的 COM `CompareDocuments` 是否存在。
8. Word `Fields.Add`/`TablesOfFigures.Add`/`Table.Cell.Formula` 是否可用。
9. Word `Protect`/`AllowEditRanges`/`FormulaHidden` 是否可写/可读。
10. PPT `CustomLayouts` 能否新增版式；`ThemeColorScheme`/`ThemeFontScheme` 能否编辑。
11. PPT `TextFrame` 字体回读；`Fonts` 嵌入/体积；`PictureFormat` 压缩。
12. PPT `convert_to_pdf` 的备注页/隐藏页/缩放选项。
13. `wps_ppt_get_slide_notes` 空备注/自定义备注版式；`NotesPage.Shapes.Item(2)` 稳定性。
14. `add_master_element` 的 image 报错；`unify_font` 覆盖表格/组合/母版。
15. `set_slide_size` 是否自动缩放内容（scaleContent）。
16. `insert_header/footer` 的 firstPage/oddEven/linkToPrevious；空 text 清空。
17. `wps_batch` 2,000 字截断对 50 份文档读取的影响；50 项连续转换的耗时/风控。
18. `wps_execute_method` 的 action 白名单（无 action 一律 Unknown action）。

**基准化**：把 80 条固化成 `test/skilleval`，每次发版量 6 个指标；下一次广告面换血只能用它 + 真机轨迹做 R14 证据。

---

## 附录 A：60 个新场景逐条结果

| id | 一句话 | 判定 | 关键缺口/说明 |
| --- | --- | --- | --- |
| XL1 | 多条件 SUMIFS 全返 0 | C | 公式填充语义/类型诊断 |
| XL2 | 产品×月份二维取值 | B | 无 XLOOKUP，INDEX/MATCH 可拼 |
| XL3 | 全表错误值混排 | C+E | 无全表错误扫描；整列覆盖有风险 |
| XL4 | 地址拆省市 | B | 文本函数可拼 |
| XL5 | 图标集对公式格不生效 | D | 桥无条件格式 icon set |
| XL6 | 固定宽度分列+Ctrl+E | C | 只有分隔符分列 |
| XL7 | CSV 乱码/001/日期 | D | 导入编码/列格式不可控 |
| XL8 | 去重保留最近修改 | B | 多键排序可拼 |
| XL9 | 透视计算字段/季度/切片器 | D | 计算字段/日期分组缺；切片器 WPS 限制 |
| XL10 | 柱+折线次轴+趋势线 | D | 组合图/次轴/趋势线缺 |
| XL11 | 冻结/分级/隐藏列 | B | 无回读 |
| XL12 | 允许编辑区域/隐藏公式 | D | AllowEditRanges/FormulaHidden 缺 |
| XL13 | 保护密码忘了 | D（边界） | WPS 不支持解密 |
| XL14 | 12 万行大表卡顿 | D | 无大文件诊断/批量公式读法 |
| XL15 | 外链更新/源丢失/断开 | D | 只有 refresh_links，无列出/断开 |
| XL16 | 12 个月度文件静态合并 | C | 无按列对齐/来源列 |
| XL17 | 自定义顺序+多键排序+高级筛选 | C | 单列排序/单条件筛选 |
| XL18 | 批注/超链接/名称管理器清理 | C | 超链接只能写不能读 |
| XL19 | 目标倒推/敏感性分析 | C | 只有 goal_seek；模拟运算表 WPS 限制 |
| XL20 | 打印区域/分页符/页脚 | B | 不能插分页符/读页数 |
| WD1 | 水印（三连） | D（边界） | WPS 页眉 Shapes 空壳 |
| WD2 | 分节/横纵混排 | C+E | 无"应用于本节"保障/无撤销 |
| WD3 | 封面无页眉/奇偶页不同 | D | 无 firstPage/oddEven/linkToPrevious |
| WD4 | 页眉横线删不掉 | D | 无页眉边框/清除 |
| WD5 | 题注/图表编号 | D | 无 SEQ/域/题注 |
| WD6 | 交叉引用/PDF 跳转 | D | 无更新域/目标选取 |
| WD7 | 图目录/表目录 | D | 无 TOC \c/表目录 |
| WD8 | 两版文档比较 | D+E | 无 CompareDocuments；无备份 |
| WD9 | 邮件合并标签/信封/席卡 | D | 一页多记录/套打缺 |
| WD10 | 带照片证书合并 | C | 图片合并域缺；文本合并可拼 |
| WD11 | 内容控件/可填写表单 | C | 有 add_content_control，参数不够 |
| WD12 | 按作者筛选修订 | C+E | 只有整体接受/按 index |
| WD13 | Word 表格公式/排序/斜线 | D | 无表格公式/排序/斜线 |
| WD14 | 文档属性/PDF 元数据 | D（边界） | WPS COM 坏壳/365 |
| WD15 | 限制编辑/只读 | D | Word 0 个 protect/restrict 工具 |
| PT1 | 母版新版式/LOGO/占位符 | D | CustomLayouts 不可新增/无锁定 |
| PT2 | 主题色/主题字体 | D | 只能 ApplyTemplate+事后改字体 |
| PT3 | 4:3→16:9 | C+E | 无 scaleContent/无备份/影响页不可知 |
| PT4 | 只换字体不动字号 | C | unify_font 无 sourceFont/preserveSize/回读 |
| PT5 | 字体嵌入/缺失 | D | 无嵌入/体积回读 |
| PT6 | 图片裁剪/压缩/替换 | C | 裁剪/换图有，压缩/批量/体积回读缺 |
| PT7 | SmartArt | D（D3 边界） | 已决策不做 |
| PT8 | 音视频 | D（D3 边界） | 已决策不做 |
| PT9 | 节/隐藏页/导出 | D | 无节/隐藏页/导出选项 |
| PT10 | 备注/提词/导出带备注 PDF | C | 写备注有，导出/可见性缺 |
| XA1 | 表格→Word 报告 | C | 插空表+N 次 set_table_cell；data 桥有 schema 无 |
| XA2 | 表格→PPT 静态页 | C | 同上，PPT 无 data |
| XA3 | Word 大纲→PPT | C | 原语拼装，无大纲导入/版式映射 |
| XA4 | PPT 备注→Word 讲义 | B | 逐页读+插入，40 次；可加门面 |
| XA5 | PDF 合并 | D（边界） | 无 PDF 宿主/合并 action |
| XA6 | PDF 拆分 | D（边界） | 同上 |
| XA7 | PDF→Word/Excel | D（边界） | 同上，OCR 不在范围 |
| XA8 | PDF 水印/页码 | D（边界/会员） | 同上 |
| XA9 | 批量转换 | D+E | 批量编排/风控/会员；只能 O(N) |
| XA10 | 批量改名/整理 | D（边界） | 无文件系统工具；Harness 有 |
| XA11 | 批量套模板/统一格式 | C+E | 跨文档批处理缺；O(N) 循环 |
| XA12 | OLE/链接对象 | D+E | 无 OLE/链接对象工具；WPS 限制 |
| XA13 | 手机端批量操作 | D（边界） | 桌面 COM 不覆盖 |
| XA14 | WPS AI 生成/改写 | D（边界） | 不在插件范围 |
| XA15 | 十万行/长文档卡顿 | C+E | 无大文件诊断/分页；与 XL14 同簇 |

## 附录 B：80 条判定的合并口径

| 项 | 数值 |
| --- | --- |
| 场景总数 | 80（v1 20 + v2 60） |
| 独立任务簇 | 约 75（4 条重叠偏重 + 1 对内部同题） |
| A/B/C/D/E | 0 / 8 / 33 / 39 / 19（E 叠加） |
| 插件范围内 | 63（B8 / C33 / D22） |
| 平台/决策边界 | 17（全 D） |
| v2 顶层调用 | ~555；help 118；隐藏执行 188；batch 子 422 |
| v2 schema 抽样 | 417 个 `(tool,args)`，0 违规 |
| 真缺口 | 22（Excel 7 / Word 9 / PPT 4 / v1 2） |

## 附录 C：过程产物与核验

| 产物 | 用途 |
| --- | --- |
| `.tmp-eval/r2-xl-a.md` / `r2-xl-b.md` | Excel 20 条调研（证据/来源） |
| `.tmp-eval/r2-wd-a.md` / `r2-wd-b.md` | Word 15 条调研 |
| `.tmp-eval/r2-pt.md` | PPT 10 条调研 |
| `.tmp-eval/r2-cross.md` | 跨应用/文件/边界 15 条调研 |
| `.tmp-eval/sim2-a..e.md` | 60 条轻量仿真 |
| `.tmp-eval/v2-source-audit.md` | 来源/去重核验（54 条：43/8/3） |
| `.tmp-eval/v2-sim-audit.md` | 仿真核验（54 条：43/7/4） |

> 外部来源抽样：26 个新 URL 全部 200 且支持场景；未发现编造引文/标题/热度。知乎/B站/小红书/百度知道在本环境 404/登录墙，未作为主来源。

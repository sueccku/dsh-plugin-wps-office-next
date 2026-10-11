# WO-001：Word 侧「读写不对称」+ 「共 Y 页」(NUMPAGES)

| 项 | 内容 |
| --- | --- |
| 编号 | **WO-001** |
| 来源 | 场景 **S4** 真机走查的 F5 / F6（[S4-真机走查报告](../scenarios/S4-真机走查报告.md) §7.3） |
| 建立日期 | **2026-10-11** |
| 状态 | **已实施（2026-10-11）** —— 四项全部落地；实施期间另修掉两个隐性缺陷。记录见 FIXES 109，真机回归 `test/word-page-setup-gaps.test.mjs` **23/23** |
| 决策 | 已决（用户 2026-10-11）：①**四项都做**；②`get_page_setup` **进广告面**；③S4 的 D7 升硬、S2 的 H3 结案；④开工实施 |

### 实施结果（2026-10-11）

| 项 | 落地形态 | 分档 |
| --- | --- | --- |
| F5a 读页面设置 | 新工具 `wps_word_get_page_setup`（action `getPageSetup`） | **广告面**（84 → 86，读与写一起进） |
| F5b 读页眉页脚 | 新工具 `wps_word_get_header_footer`（action `getHeaderFooter`） | 隐藏面 |
| F5c 更新目录 | `wps_word_generate_toc {action:"insert"\|"update"}` | 已广告工具加参数 |
| F6 「共 Y 页」 | `wps_word_insert_page_numbers {includeTotal:true}` | 已广告工具加参数 |

桥 action **263 → 265**、注册工具 **268 → 270**、广告面 **84 → 86**（`verify.mjs` 的 `EXPECTED_ACTIONS` 与 README/HANDOFF 数字已同步）。

> **追加决策（2026-10-11，用户）**：实施中发现本工单 §5 有一处事实错误 —— `wps_word_set_page_setup` 其实**不在广告面**（只有 `wps_call`），
> 于是"读广告、写隐藏"成了新的反向不对称。用户决定**把 set 也提到广告面**，因此广告面是 **84 → 86**（读 85、写 86），
> `operations.ts` 的 `advertised` 与 `toolset.ts` 的 `STANDARD_TOOLS` 两处声明同时改（少改一处会被 `spec-reproduction` 的
> "advertised set is reproduced" 抓出来）。真机回归里 `set_page_setup` 也从 `wps_call` 改成**直接调用**，顺带验证广告路径可用。

**实施中挖出的两个隐性缺陷**（都不在原工单的缺口清单里，是做 F5a/F6 时被测出来的）：

| # | 现象 | 根因 | 修法 |
| --- | --- | --- | --- |
| **9-1** | `includeTotal` 的 NUMPAGES **100% 写进 `word/header1.xml`**（页眉），`PageNumbers.Add` 却在页脚 | 本 WPS 的 `HeaderFooter.Range` 会解析到「当前激活的页眉/页脚故事」；4 种插入法实测只有 `View.SeekView` + `Selection` 部件与顺序都对 | 走 `SeekView` 路径写入，`WholeStory()` 后数域回读；插不进去按 R11 回 `success:false` |
| **9-2** | `set_page_setup {orientation:landscape, topMargin:56, leftMargin:72}` → 读回 `T=72 … R=56`（**边距被旋转**），且工具**一个边距都没回报** | 桥先设边距、后设方向（方向变更会旋转已有边距）；且桥回读键名 `marginTop` 与 handler 读的 `topMargin` 不一致 | 方向/纸张先设、边距后设；回读改用公开键名并**始终回报四个边距 + 方向 + 纸张** |
| 目标 | 把 4 个"模型想做但做不到"的 Word 能力补进工具面，或明确否决；并堵住"绕开插件直接改文件"的口子 |
| 关联 | [tool-roadmap §6](../tool-roadmap.md) 方法 C（真实任务驱动 → 缺口工单）、S2 报告 H2/H3、FIXES 90（按真实调用数据重平衡广告面） |

---

## 1. 缺口从哪来（现场证据，不是推测）

S4 的 9 次失败调用里有 3 次是"**工具不存在**"，最终交付物里还出现了**插件做不到、模型自己改文件补上**的一处：

| 缺口 | S4 现场证据 | 模型当时的绕法 |
| --- | --- | --- |
| **读页面设置** | 直呼 `wps_word_get_page_setup` → `未知工具`；`set` 早在广告面 | 放弃读取，改用自写 `_build_report.py` 读 `word/document.xml` |
| **读页眉页脚** | `wps_help {tool:"wps_word_get_header"}` → `未找到工具` | 用 Python 解 `header1.xml` / `footer1.xml` |
| **更新目录** | `wps_help {tool:"wps_word_update_toc"}` → `未找到工具`；只能反复 `generate_toc` 重插（3 次） | 重插目录（不能只刷新） |
| **「共 Y 页」NUMPAGES** | 最终 docx 页脚**确有** `NUMPAGES` 域，但全程 **0 次 `insert_footer`** —— 是模型直接改 OOXML 加的 | 绕开插件直接改文件 |

> **这才是本工单真正的价值**：四条里三条的绕法是「绕开插件直接动用户的文件」。
> 这类操作**没有 tool 调用、没有 warnings、没有回读**，落在 R11「不变量 / 消灭假成功」防线之外 ——
> 一个"看起来做完了"的报告，可能完全没经过插件。

---

## 2. 可行性实测（裸 COM，PS 5.1 `-STA`，与产品桥同环境）

探针（可重跑）：`test/.artifacts/scenario-s4/tools/probe-word-gaps.ps1` + `probe_word_gaps_doc.py`
（ASCII-only 源码、独立于插件；产物在 gitignored 的 `test/.artifacts/scenario-s4/run/`）

### 2.1 读页面设置（`$doc.PageSetup`）——**11 个属性全部可读**

```
LeftMargin 70.85  TopMargin 70.85  BottomMargin 70.85  RightMargin 70.85
PageWidth 595.3   PageHeight 841.9  PaperSize 7 (A4)    Orientation 0 (portrait)
HeaderDistance 36 FooterDistance 36 DifferentFirstPageHeaderFooter 0
```

（70.85pt = 2.5cm、595.3×841.9pt = A4，值与 python-docx 写入的一致 → **写—读闭环成立**）

### 2.2 读页眉页脚——可读，注意尾部 `\r`

```
headerText = "PROBE-HEADER\r"   footerText = "PROBE-FOOTER\r"   footer 域数 = 0
```

### 2.3 更新已有目录——可更新

```
TablesOfContents.Count = 1        Update() → ok
更新后目录段落数 = 4（标题 + 3 条目）
```

### 2.4 插入 NUMPAGES 域——可插入，且 **WPS 当场算出缓存值**

```
Fields.Add(range, -1, 'NUMPAGES', $true) → added=true, fieldsAfter=1
Result.Text = "5"            （探针文档被撑到 5 页，WPS 真的算了）
与 PageNumbers.Add 并存 → 页脚域数 = 2
```

**落盘 XML 实证**（`word/footer1.xml`）：

```xml
<w:instrText> PAGE </w:instrText> ... <w:t>1
<w:instrText> NUMPAGES \* MERGEFORMAT </w:instrText> ... <w:t>5
```

→ 域**真写进文件**，且**带缓存结果**，所以独立核验器可以直接从 XML 判它，不必依赖界面渲染。

> **结论：四项全部可行，没有一项卡在 WPS 上。** 对照 README「现在还做不到的」清单（水印 / 文档属性 / 切片器 / 场景管理器 / 圆角半径），
> 这四项不在其中 —— 它们只是**没做**，不是做不了。

---

## 3. 建议方案

### 3.1 F6「共 Y 页」→ **扩展** `wps_word_insert_page_numbers`（首选，0 个新工具）

| 项 | 内容 |
| --- | --- |
| 新参数 | `includeTotal: boolean`（默认 `false`，向后兼容） |
| 语义 | 在页码域之后追加 `NUMPAGES` 域，形成"第 X 页 / 共 Y 页" |
| 桥侧 | `insertPageNumbers` case 内，`PageNumbers.Add` 之后在同一个 `Footers.Item(i)` 的 Range 末尾 `Fields.Add($r, -1, 'NUMPAGES', $true)`；返回 `{ totalPagesField: true, fields: [...] }` |
| 为什么不用新工具 | 真实诉求就是"页码 + 总页数"这一件事；新工具会让 `insert_page_numbers` 与它语义重叠 |
| 备选（否决） | 通用 `wps_word_insert_field`：域类型枚举会变成新的维护面，且容易变成"什么都能塞"的口子 |
| 必须先量 | PAGE 与 NUMPAGES 之间的分隔文本（` / 共 ` + ` 页`）落在哪个 run、续写会不会破坏已有页码格式 |

### 3.2 F5a 读页面设置 → 新增 `wps_word_get_page_setup`

| 项 | 内容 |
| --- | --- |
| 桥 action | `getPageSetup`（纯读，与 `setPageSetup` 共用同一条解析与回读逻辑，杜绝两套口径） |
| 参数 | `section`（默认 1） |
| 返回 | `orientation` / `paperSize` / `topMargin` / `bottomMargin` / `leftMargin` / `rightMargin`（**单位与键名与 `setPageSetup` 完全对称**）+ `pageWidth` / `pageHeight` |
| 错误路径 | 节号越界、无活动文档，都明确报错（沿用现有措辞） |
| 分档 | 见 §5（我倾向**直接广告**） |

### 3.3 F5b 读页眉页脚 → 新增 `wps_word_get_header_footer`

| 项 | 内容 |
| --- | --- |
| 桥 action | `getHeaderFooter` |
| 参数 | `section`（默认 1） |
| 返回 | 每节的 `headerText` / `footerText`（**归一掉尾部 `\r`**）+ 各自 `fields: [{ type, code }]`（把 `instrText` 归成 `PAGE` / `NUMPAGES` / 其它） |
| 分档 | 隐藏面（`wps_call`） |

### 3.4 F5c 更新目录 → `wps_word_generate_toc` 加 `action: "insert" \| "update"`

| 项 | 内容 |
| --- | --- |
| 新参数 | `action`（默认 `"insert"`，向后兼容） |
| update 语义 | 对已有的每一个 `TablesOfContents` 调 `Update()`，**不新增目录**；没有目录时如实报错并提示用 `insert` |
| 返回 | `{ action:"update", tocCount, updated, entries }` |
| 为什么不用新工具 | 模型本来就在猜 `wps_word_update_toc`；但"插入/更新"是同一个对象的两半，做成模式比做成两个工具更难用错 |
| 备选 | 新工具 `wps_word_update_toc`（命名更直白，代价是 +1 工具面） |

---

## 4. 改动清单与生成链（精确到文件）

| # | 文件 | 改什么 | 之后必须跑 |
| --- | --- | --- | --- |
| 1 | `mcp/scripts/wps-com.ps1` | 新增 `getPageSetup` / `getHeaderFooter` case；扩展 `insertPageNumbers` / `generateTOC`。**纯 CRLF、无 BOM** | `powershell -NoProfile -File scripts\build-host-actions.ps1`；`git diff --exit-code -- host/wps-actions.ps1` |
| 2 | `mcp/src/spec/operations.ts` | 动作键 + 参数键声明 | `cd mcp; npx tsc`；`git diff --exit-code -- mcp/dist` |
| 3 | `mcp/src/tools/word/produce.ts` / `format.ts` | 定义 + handler（`index.ts` 视注册方式同步） | `node scripts\gen-tool-surface.mjs` → `git diff --exit-code -- spec` |
| 4 | `spec/*`（生成物） | 工具定义、签名、动作键、参数契约 | `node scripts\param-contract.mjs`；`git diff --exit-code -- docs\param-contract.md` |
| 5 | `skills/wps-word/reference.md` 等 | 工具清单与说明（生成物） | `node scripts\gen-skill-tools.mjs`；`git diff --exit-code -- skills` |
| 6 | `docs/tool-coverage.md` | 覆盖矩阵（生成物） | `node scripts\gen-tool-coverage.mjs` |
| 7 | `docs/current-numbers.md` + README 徽章 | 广告面/工具数（生成物） | `node scripts\gen-numbers.mjs --check` |
| 8 | `test/word-*.test.mjs` | **每个新工具至少一条非矩阵断言** | `node test\spec-reproduction.test.mjs`；`npm run test:static` |
| 9 | 若进广告面 | `mcp/src/server/toolset.ts` 的 `STANDARD_TOOLS` | `node scripts\verify.mjs --static`（含预算门禁） |
| 10 | 全量 | — | `npm run lint`、`node scripts\verify.mjs --static`、`npm run verify:package` |

> 顺序注意：**桥源码是唯一真源**，`host/wps-actions.ps1` 是它的生成物；先改桥、再生成、再改 TS。

---

## 5. 广告面与预算影响

- **现状**（今天 `wps_status` 实测）：`advertisedTools = 84`、`registeredTools = 268`、`hiddenTools = 184`。
- **硬门禁**：`scripts/verify.mjs` 里 `BUDGET = { maxTools: 100, maxSchemaBytes: 60000 }`
  （历史：45/25,000 → 60/32,000 → 70/40,000 → 100/60,000）。
  → **还有 16 个广告位和约 22KB 余量，预算不是本工单的瓶颈。**
- **推荐分档**：

  | 项 | 建议 | 理由 |
  | --- | --- | --- |
  | F6 `includeTotal` | **不动分档**（0 新工具，0 变化） | 只是给已广告的工具加参数 |
  | F5a `get_page_setup` | **进广告面** | `set_page_setup` 已在广告面；"能改不能读"正是 S4 缺口的成因；对称读的发现成本最低 |
  | F5b `get_header_footer` | 隐藏面（`wps_call`） | 长尾，先看真实调用数据 |
  | F5c `update` 模式 | **不动分档**（0 新工具） | 同上 |

- 若任一工具上调到广告面，**三个数字必须同步**：`spec/advertised.json`（生成物）、README 徽章、`docs/current-numbers.md`（生成物），
  否则 `gen-numbers --check` / `verify --static` 会红。
- **不要用"放进隐藏面"来回避问题**：S4 里模型花了 **21 次 `wps_help`** 才把隐藏能力用起来（`wps_call` 27 次）。
  隐藏面解决的是 token 预算，不解决发现成本。

---

## 6. 风险与已知坑（含本次探针实测到的）

1. **覆盖率棘轮**：`scripts/lib/coverage-tiers.mjs` 的 `TIER_LEDGER = { bespoke: 260, matrixOk: 8, matrixAny: 0 }`
   —— **每新增一个工具都必须有一条"非矩阵"真断言**，否则 `bespoke` 地板失守、CI 红。这条最容易被漏。
2. **动作数硬编码**：`scripts/verify.mjs` 的 `EXPECTED_ACTIONS = 263` 要**显式**更新（生成器同时校验 源码 == 生成物 == 该数）。
3. **字节约定**：桥源码 CRLF 无 BOM；`host/wps-actions.ps1` 与手写宿主是 UTF-8 BOM（PS 5.1 否则读不懂中文）。`scripts/lint.mjs` 会查。
4. **`Fields.Add` 之后的读数不可信（本次实测踩到）**：插入后立刻取 `f.Result.Text` 得到 `"5"`，
   但 `$d.Fields.Update()` **之后再取同一个对象会返回空串**（`afterUpdate=""`）。
   → handler 的返回值必须**重新取 Field 或回读文档**，不能把第一次读数当结论。这正是本项目"不许假成功"的老坑。
5. **页眉页脚文本带尾随 `\r`**（实测 `"PROBE-HEADER\r"`），返回前要归一，否则用户会看到莫名换行。
6. **`TablesOfContents.Count` 依赖文档已被 WPS 载入**；没有目录时要如实回报 0 并给出下一步（别静默成功）。
7. **NUMPAGES 的版式**：探针只证明了"域能写进页脚并被计算"，**没有**量过与既有页码之间的分隔文本是否好看。
   实现期第一件事就是在真机上量这个，不要在文档里承诺版式。

---

## 7. 验收判据（做完怎么算过）

| # | 判据 | 档 |
| --- | --- | --- |
| 1 | 单测（不提真机）：参数契约、错误路径（无文档 / 节号越界 / 无目录）、返回形状 | 硬 |
| 2 | **真机写—读闭环**：`set_page_setup` 写入 → `get_page_setup` 读回同值（±1pt） | 硬 |
| 3 | **真机落盘**：`insert_page_numbers {includeTotal:true}` 后，导出 docx 的 footer 同时含 `PAGE` 与 `NUMPAGES` 域，**且缓存结果非空** | 硬 |
| 4 | **真机**：`generate_toc {action:"update"}` 后目录条目数 ≥ 小节数，且没有多出第二份目录 | 硬 |
| 5 | **场景回归升级**：S4 的 **D7 由软升硬**（页脚必须同时有 PAGE + NUMPAGES 且带缓存值）；S2 的 **H3 结案** | 硬 |
| 6 | 静态门禁全绿：`npx tsc` / `build-host-actions` 无 diff / `lint` / `verify --static` / `param-contract` / `gen-*` 无 diff / `spec-reproduction` / `test:static`（含棘轮） | 硬 |

---

## 8. 工作量与优先级

| 项 | 估时 | 优先级 | 理由 |
| --- | --- | --- | --- |
| **F6 `includeTotal`** | 0.5 天 | **P0** | 用户可见价值最高（"第 X 页 / 共 Y 页"是打印文档的刚需）、成本最低、且直接堵住"绕开插件改 OOXML"的口子 |
| **F5c 更新目录** | 0.5 天 | P1 | 真机已实证可更新；插而不更会留下过期目录 |
| **F5a 读页面设置** | 0.5 天 | P1 | 读写对称；`set` 的验证本来就该有 `get` |
| **F5b 读页眉页脚** | 0.5 天 | P2 | 长尾，可延后 |
| 测试 + 生成物 + 门禁 | 1 天 | — | 棘轮要求每个工具都有真断言 |

**合计 2–3 天**（含真机量）；建议顺序 **F6 → F5c → F5a → F5b**，每项可独立发布。

---

## 9. 本工单明确**不做**的事

- 不做通用 `wps_word_insert_field`（域类型枚举会成为新的维护面，也容易变成"什么都塞"）。
- 不改 `PageNumbers.Add` 的既有路径 —— 它在 S4 真机上 1/1 成功，输出"该处共 N 个页码域"也如实。
- 不因为这次就调整 `BUDGET` 上限（16 个广告位、22KB 余量，先用隐藏面/对称读消化）。
- 不动"绕开插件改文件"这件事本身 —— 那不是插件能拦的；本工单只负责**把缺的能力补上，让模型没有绕的理由**。

---

## 10. 决策请求

1. **范围**：四项都做，还是只做 **F6**（价值/成本比最高）？
2. **分档**：全隐藏面 / **`get_page_setup` 直接进广告面**（我的建议）/ 四项全进广告面？
3. **收口**：做完是否按 §7-5 把 S4 的 D7 升为硬判据、把 S2 的 H3 结案？
4. **排期**：本轮做完 S4 走查就到此为止，还是直接开工 WO-001？

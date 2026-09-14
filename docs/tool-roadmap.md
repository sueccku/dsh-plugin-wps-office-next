# 工具面路线图：从「迁移来的演示型目录」到「一站式 WPS 工具面」

> 依据：2026-09-13 的能力审计（脚本方法与实测数字见 §1）。本文件是**规划**，不含已实施内容；
> 实施进度会回填到 docs/PROGRESS.md 与 docs/FIXES.md。

## 0. 目标与判据

**目标**：一站式——用户只装这一个 WPS 技能 + MCP 插件，就能把表格与文字的事做完，不必再去别处找工具。

**判据（决定一个能力要不要进桥）**：**「如果我不做，还有谁会做？」**

| 类别 | 说明 | 处置 |
|---|---|---|
| A. WPS COM 深水区 | 需要读 COM 对象模型 + 实测 WPS 差异：ListObject、Word 表格/修订、WPS 特有坑（Presentations.Add 返回 0 页、多格 Formula 被拒…） | **必做**，基本只有我们会做 |
| B. 不体面但必需的另一半 | 读/删/清、used range、autofit、取消冻结、页眉页脚、文档属性——上游做 demo 时用不到，后来者也不会补 | **必做**，其中 31 项已在桥里实现，只差出口 |
| C. 跨对象工作流/美化 | KPI 卡、时间线、流程图、配色美化 | **不进桥**，活在技能「组合配方」里 |

**优先级**：Excel 第一 → Word 第二 → PPT 做减法；兼容层在本轮内清理。

## 1. 基线（实测）

| 指标 | 值 | 说明 |
|---|---|---|
| 桥 action | **259 个 case 条目 / 255 个唯一名** | 4 个重复标签是死代码：set3DRotation / set3DDepth / set3DMaterial / create3DText |
| 有工具引用的 action | 201 | |
| 无工具引用的 action | **55** | 31 真缺口 + 22 场景死代码 + 2 可判定重复（unfreezePanes、slide.unifyFont） |
| 注册工具 | **237** | 含 18 个废弃转发名 + 12 个 builtin |
| 广告工具 | 44 / 23,593 字节 | 上限 45 / 25,000 |
| 分布 | excel 75 / word 32 / ppt 87 / common 9 | + 16 builtin/门面 |
| 测试 | 310 项 + verify 23 项 + e2e 19 项 | 全绿，CI 跑静态部分 |

> **P0 完成后的当前值**：桥 action **231**、注册工具 **209**（广告 44 / 隐藏 165；18 个废弃别名不占注册位，
> 只在派发期解析）、builtin 组从 16 降到 **5**。下面 §1 的表格是 P0 之前的审计快照，保留用于对照。

**真缺口清单（桥里已实现、无任何工具引用，共 31 项 = 20 + 4 + 5 + 2）**

- Excel（20）：`getActiveWorkbook`、`copyFormat`、`clearFormats`、`getConditionalFormats`、`removeConditionalFormat`、`getDataValidations`、`removeDataValidation`、`autoFitColumn`、`autoFitRow`、`autoFitAll`、`wrapText`、`groupColumns`、`findInSheet`、`replaceInSheet`、`getNamedRanges`、`deleteNamedRange`、`refreshLinks`、`consolidate`、`calculateSheet`、`getExcelContext`
- Word（4，全部位于桥的 Word 段）：`insertHyperlink`、`getBookmarks`、`getComments`、`getDocumentStats`
- PPT（5）：`setSlideTheme`、`setShapeRoundness`、`endSlideShow`，以及两个重复实现 `slide.add`、`slide.beautify`（P0-3 直接删）
- 公共（2）：`openFile`、`getActivePresentation`

另有 2 项待判定：`unfreezePanes`（若 `freeze=false` 已覆盖则删）、`slide.unifyFont`（重复实现，待删）。

**场景死代码（21 个唯一名，共 22 个 case 条目——`create3DText` 有重复标签）**：
`autoBeautifySlide`、`beautifyAllSlides`、`createKpiCards`、`createStyledTable`、`addTitleDecoration`、`addPageIndicator`、`createProgressBar`、`createGauge`、`createMiniCharts`、`createDonutChart`、`autoLayout`、`smartDistribute`、`createGrid`、`createFlowChart`、`createOrgChart`、`createTimeline`、`create3DText`、`setShapeFullStyle`、`addConnector`、`addArrow`、`applyColorScheme`

（校验：31 真缺口 + 2 待判定 + 21 唯一死代码 = 54 个唯一名；加上 `create3DText` 的重复条目 = 55 个无工具引用的 case 条目 ✓）

## 2. 已锁定的四个参数（2026-09-13 定）

| # | 参数 | **决定** |
|---|---|---|
| D1 | 广告面预算 | **70 个 / 40,000 字节**（2026-09-13 二次上抬；见 FIXES 42） |
| D2 | Word 长尾边界 | **全都要**：表格/修订/页码/水印/文档属性/批注读删 + 内容控件/脚注尾注/分栏 + 邮件合并/索引/交叉引用 |
| D3 | PPT 放弃清单 | **确认放弃**：媒体(视频/音频)、SmartArt、讲义、3D 族、美化族；**保留**：版式列表、主题、尺寸、母版、节 |
| D4 | 兼容窗口 | **确认**：18 个废弃名保留一个发布周期（改为派发别名，不占注册位）；12 个 builtin 立即删 |

因此 P0-4 的处置确定为：**恢复 `wps_ppt_set_slide_theme` 工具**（主题属于 D3 的保留项），并修正 FIXES 第 1 条。

## 2b. 原始建议（存档）

## 2. 需要拍板的四个参数

| # | 参数 | 建议 | 影响 |
|---|---|---|---|
| D1 | 广告面预算 | 45 / 25,000 → **60 / ~32,000** | 广告面是"一站式"的瓶颈：Excel 做深后 45 个装不下。理由：请求前缀可被 prompt cache 复用，边际成本低；minimal 档保留给成本敏感场景 |
| D2 | Word 长尾边界 | A 档必做：表格/修订/页码/水印/文档属性/批注读删；B 档做：内容控件/脚注尾注/分栏/文档属性；C 档暂缓：邮件合并/索引/交叉引用 | 决定 Word 从 32 走到 ~55 还是 ~75 |
| D3 | PPT 放弃清单 | 放弃：媒体(视频/音频)、SmartArt、讲义、3D 族、美化族；保留：版式列表、主题、尺寸、母版、节 | PPT 工具 87 → ~55 |
| D4 | 兼容窗口 | 18 个废弃名保留一个发布周期（改为派发别名，不占注册位）；12 个 builtin 立即删 | 注册表 237 → ~207 |

## 3. 不变量（做完后由 CI 保证）

1. **对象动词完整度**：每个对象声明必备动词集，缺一个就是 CI 失败（不再靠人工审计）。
2. **无静默盲区**：每个 action 必须在 spec 里有条目；读不出参数的 action，生成器**直接失败**，不允许"跳过校验"。
3. **广告面可生成**：由 spec 的 `advertised` 标记产出，人手不再维护工具名数组。
4. **结果统一**：所有工具返回同一信封（`ok/op/data/warnings/effect`），中文句子降级为 `summary`。
5. **零人肉翻译层**：`$paramAliases` / `$paramContainers` / `$helperKeys` 三张表归零。

## 4. 阶段总览

| 阶段 | 目标 | 主要工作 | 验收 | 依赖 |
|---|---|---|---|---|
| **P0 清理** | 缩小后续工作面 | 删 builtin、废弃名改派发别名、删 PPT 死代码与重复标签、修 theme 漂移 | 注册 237→~207、action 259→~232、310 测试 + e2e + CI 全绿 | — |
| **P1 契约真源** | spec 单一真源 + 生成器 | spec 格式、5 类产物生成、删三张人手表、param-contract 退化为断言 | **逐字节复现今天的 44 个 schema、键表、4 份 reference.md** | P0 |
| ↳ P1 完成 | 五步全部落地（FIXES 33–37） | spec 209 条 + 生成器 + 验收 **12 项全绿**；调用点解析器共享；每个参数都有明确去向（未归类 0）；**三张手写表全部归零**（aliases / containers / helperKeys 都进 spec，每次都以宿主产物逐字节未变作证）；动态 action 未声明即失败 | **P2 Excel 做深**：第一波挂出 20 个已实现却无出口的 Excel action + used range 一等公民。剩余债务：别名 62（两种改名机制已分清）、未工具化 action 29、raw schema 32 | — |
| **P2 Excel 做深** | 一站式表格 | 5 个波次（§5.3） | **已完成**：37 个新工具、5 个真实 WPS 验收测试、两场景一键 e2e 28 项全绿 | P1 |
| **P3 Word 做深** | 一站式文档 | 挂已有 4 项 + 表格读写 + 文档生产族 | **已完成**：33 个新工具、66 项验收；水印与文档属性实测 WPS 不支持（如实记档） | P1 |
| **P4 PPT 收敛** | 做减法 | 删死代码、聚合碎片 setter、保留结构能力 | **已完成**：死代码在 P0-3 删除；两波合并把 88 → **76**（形状效果四合一 + 动画三合一 + 表格样式三合一 + 页脚三合一 + 3D 族按 D3 删除） | P1 |
| **P5 收尾** | 发布准备 | execute_method 决策、广告面与预算重定、文档与技能重生成 | 门禁全绿 + e2e 扩展场景 | P2–P4 |

## 5. 分阶段任务清单

### P0 清理（低风险，先做）——**已完成 2026-09-13**

实施结果：桥 action **259 → 231**（删 28 个 case 条目 / 1039 行：22 个场景死代码、3 个点号重复实现、
4 个重复标签的后一份）；注册工具 **237 → 209**（删 11 个 builtin + 18 个废弃名的重复定义，共 1319 行）；
生成器别名表 **20 → 17**、容器表 **13 → 12**；全部测试 + verify + e2e + CI 绿。详见 docs/FIXES.md 第 32 条。

| ID | 任务 | 依据 | 验收 |
|---|---|---|---|
| P0-1 | 删 12 个 builtin：6 个与 `wps_*` 重复 + `check_connection` + 4 个缓存工具（`cache_data`/`get_cached_data`/`list_cache`/`clear_cache`） | 无人使用的内存缓存，结果形状还与主工具不一致（S5） | `wps_help` 目录 219 → 207 |
| P0-2 | 18 个废弃名从「注册工具」改为「派发期别名」 | 现在每个废弃名占一个注册位 + 一份重复 schema | 注册 237 → 219，废弃名仍可直接调用 |
| P0-3 | 删 22 个场景 action + `slide.add`/`slide.unifyFont`/`slide.beautify` 3 个点号重复 + 4 个重复 case 标签 | 工具早已删除，桥里留死代码 | action 259 → ~232，生成器 case 数与唯一名相等 |
| P0-4 | 修 theme 漂移：`setSlideTheme` 恢复为工具（结构性操作）或删除并改 FIXES 第 1 条 | FIXES 1 宣称该工具存在，实际已被第 22 条当"美化"删掉 | 文档与现实一致 |
| P0-5 | `execute_method` 文档化：明确它是逃生舱、不在技能里推荐 | 覆盖不全时它救过场（e2e 第一轮靠它建 Word 文档） | 技能文档写明 |
| P0-6 | 全量回归 | — | 310 测试 + verify 23 + e2e 19 + CI 漂移检查全绿 |

### P1 契约真源（对应上一轮设计讨论的 Phase 0）

| ID | 任务 | 验收 |
|---|---|---|
| P1-1 | 定 spec 格式（YAML）+ 加载器 + 校验（required/类型/枚举/单位/effect） | 能描述今天全部 232 个 action 的参数 |
| P1-2 | 生成器：spec → ①工具 schema ②宿主键表 ③紧凑签名索引 ④`reference.md` ⑤广告集 | 五类产物齐全 |
| P1-3 | **复现验收**：与今天的产物逐字节比对 | 44 个 schema、键表、4 份 reference.md 全部一致；不一致就改 spec 设计 |
| P1-4 | 删 `$paramAliases`/`$paramContainers`/`$helperKeys`；公开参数名 = 桥读的键（桥里烂名顺手改） | 三张表归零；生成器对读不出的 action 直接失败 |
| P1-5 | `param-contract.mjs` 退化为「spec ↔ 生成物 ↔ schema 相等」断言；CI 增加该步 | 盲区 9 + 1 → 0 |

### P2 Excel 做深（第一优先）

| ID | 波次 | 内容 |
|---|---|---|
| P2-0 | 覆盖矩阵 + 实测 | 产出 `docs/tool-coverage.md`（对象 × 动词 × {action, tool, advertised, WPS支持, 证据}）；每项标 支持/部分/不支持；**待核实**：桥里 Excel 形状/文本框有 141 处命中，是否有出口 |
| P2-1（**已完成**：18 个已挂，FIXES 38/39） | 挂出已有能力（零 COM 成本） | 21 个 Excel 缺口全部挂工具：named range get/delete、条件格式 get/remove、数据验证 get/remove、autofit×3、wrapText、groupColumns、unfreezePanes、findInSheet/replaceInSheet、copyFormat/clearFormats、refreshLinks、consolidate、calculateSheet、getExcelContext、getActiveWorkbook、openFile；**used range / sheet info 升为一等公民**（范围地址、行列数、表头、可见性、标签色） |
| P2-2（**已完成**，FIXES 40） | ListObject 全族（新桥代码） | 从区域建表、列表/读取结构、增删行、总计行、表格样式、转回区域、resize、结构化引用、表内自动筛选 |
| P2-3（**已完成**，FIXES 41） | 数据与打印族 | CSV/文本导入导出、页面设置（方向/纸张/边距/缩放/居中）、页眉页脚、打印标题、打印预览、工作表标签色与隐藏/显示、分级显示(Outline)、公式审计（依赖/被依赖/数组公式）、合并计算 |
| P2-4（**已完成**，FIXES 43） | 高级项（按实测） | 透视表删除/缓存/刷新、单变量求解(goalseek)、场景管理器、外部数据刷新、迷你图、切片器、图表删除与系列级设置 |
| P2-5（**已完成**，FIXES 44） | Excel 重场景 e2e | 新增一个真实 Excel 任务进一键 e2e（第二场景），覆盖 ListObject + 条件格式 + 打印设置 |

### P3 Word 做深

| ID | 波次 | 内容 |
|---|---|---|
| P3-1（**已完成**，FIXES 45） | 挂出已有能力 | `getBookmarks`、`getComments`、`getDocumentStats`、`insertHyperlink`、域列表、修订查询 |
| P3-2（**已完成**，FIXES 45） | 表格读写编辑（新桥代码） | 读取表格结构、单元格读写、增删行列、合并/拆分、表格样式、表格 ↔ 文本转换 |
| P3-3（**已完成**，FIXES 46） | 文档生产 | 页码、水印、文档属性、分栏、页眉页脚进阶、目录更新、修订接受/拒绝、批注读删 |
| P3-4（**已完成**，FIXES 47） | 长尾（按 D2 档位） | 内容控件、脚注/尾注、邮件合并、索引/交叉引用 |

### P4 PPT 收敛（做减法）

| ID | 内容 | 验收 |
|---|---|---|
| P4-1 | 删 22 场景 action + 3 点号重复 + 4 重复标签（与 P0-3 合并执行） | action 数下降 |
| P4-2（**第一波已完成**，FIXES 48） | 聚合碎片 setter：`shape_shadow`/`shape_gradient`/`shape_border`/`shape_transparency`/`3d_*`/`image_style`/`table_cell_style`/`table_row_style` → 归并为少量结构化工具 | PPT 工具 87 → ~55 |
| P4-3（**已具备**） | 保留结构能力：版式列表、主题、演示尺寸、母版、节、结束放映 | 结构性操作不缺 |
| P4-4 | 技能配方承接美化类需求（已有「组合配方」一节） | 技能文档同步 |

### P5 收尾

| ID | 内容 |
|---|---|
| P5-1（**已完成**，FIXES 50） | `execute_method` 去留决策：保留为隐藏逃生舱，契约写死 |
| P5-2（**已完成**，FIXES 50） | 广告面定为 **69 / 37,573**（上限不再调整：D1 的 70 / 40,000 正好贴合）；`wps_help` 索引已重生成 |
| P5-3（**已完成**，FIXES 50） | 文档同步：README 工具面、FIXES、PROGRESS、技能 `reference.md` 全部重生成 |
| P5-4（**已完成**） | 发布准备：CHANGELOG、tag、Release —— v0.2.0（首个发布）与 v0.2.1（文档与打包修正）均已发布；npm 未发布（按约定「真正要发时再做」） |

## 6. 「还要加什么」的可重复发掘方法

单靠一次人工列举不够，规划里把"发掘"本身做成可重复的三条方法：

| 方法 | 做法 | 产出 |
|---|---|---|
| A. COM 类型枚举 | 脚本化 dump 每个对象的成员（Application → Workbooks → Sheets → Range → …），与我们的 action 名求差 | 候选成员清单 |
| B. 官方对象清单核对 | 按应用维护人工核对表（Excel/Word VBA 对象模型主要对象与常用成员） | 覆盖矩阵的"应有动词"列 |
| C. 真实任务驱动 | 每阶段跑 2–3 个真实场景，记录"模型想做但做不到"的每一步 | 缺口工单 |

三者的结果都汇进 `docs/tool-coverage.md`（生成物），它同时是审计报告、缺口清单、广告面选择依据。
**注意**：方法 A/B 只能给出"WPS 可能有"，是否可用必须**逐项实测**；不支持的要明确记入 schema 与文档，不能假装支持。

## 7. 风险与回退

- **还没有 release**：兼容成本处于最低点，是本轮敢大动的最大理由。
- **每阶段独立可发布**：P0/P1 是纯清理与重构，可随时停下；P2/P3 按波次增量交付。
- **三层安全网**：P1-3 的逐字节复现验收、现有 310 项测试 + verify + 一键 e2e、CI 的三处漂移检查。
- **未知项先实测**：任何"WPS 可能不支持"的项，先写探针脚本实测再决定进不进 spec。
- **回退**：每阶段一个分支 + 一个 PR；CI 必须全绿才合并。

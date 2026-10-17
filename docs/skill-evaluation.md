# Skill 任务评估：固定场景集与比较方法

> 依据：改进指南 §10.2 / §10.3 / §10.4。
> 本文件固定**输入、初始文件、用户提示与成功标准**，让「新旧 Skill 谁更好」可以重复测量。
> **状态说明（不要误读）**：本文件里的场景标准已经固定下来；本轮**没有跑模型侧对比**
> （本环境没有可用的 DSH 交互会话，也没有跑三轮以上的预算）。因此下面每一行都明确标了
> 「已有确定性证据」还是「仅固定标准，待跑」。**"写了测试用例"不等于"场景已经通过"。**

## 1. 场景集

初始文件一律用**专用临时文档**（`test/.artifacts/<scenario>/`），不拿用户正在编辑的文件做破坏性测试。
"确定性证据"一列指的是**不依赖模型**、可重复执行的检查（裸 COM / 外部解析 / 工具返回值）。

| # | 场景 | 用户提示要点 | 成功标准（核心检查） | 确定性证据 | 模型侧状态 |
|---|---|---|---|---|---|
| 1 | WPS 未启动，新建 Word 并保存 | "新建一份文档，写标题和一段正文，存到 `<路径>`" | 不因 `wps_status` 的 `connected:false` 放弃；文档真的落盘 | `wps_word_create_document` + `wps_common_save_as` 路径存在（`test/target-identity.test.mjs` 覆盖同一条链路） | 待跑（需要先停掉 WPS） |
| 2 | 只加载 Word Skill 修改文档 | 同 3，但只挂载 `wps-word` | 能取得必要共享协议（目标确认/保存策略/验证/失败分流） | 静态：`test/skill-quality.test.mjs` 断言四个 SKILL.md 都带共享协议摘要与 `wps-office-next` 指针 | 待跑（Harness 侧加载行为未验证，见 §3） |
| 3 | 两份 Word 文档同时打开 | "改 `<文件A>`，别动另一个" | 明确目标、**不可误改另一份** | 身份读回：`wps_word_get_active_document` 的 name/path（`test/target-identity.test.mjs`）；歧义警告 `test/target-ambiguity.test.mjs` | 待跑 |
| 4 | 修改报告但保留原件 | "改报告但别动原稿" | 副本身份正确、原件内容不变 | `wps_common_save_as` 后身份读回；原件目录 mtime/内容抽查（脚本化） | 待跑 |
| 5 | 已有目录的报告排版 | "统一字体、标题分层、刷新目录" | 刷新目录而非重复插入 | 桥侧 `generate_toc {action:"update"}` 在没有目录时明确报错、有目录时只刷新（读 `mcp/scripts/wps-com.ps1` 的 `generateTOC`） | 待跑 |
| 6 | 多节横竖混排 | "第 2 节改横向，其他节别动" | 只改目标节 | 逐节 `wps_word_get_page_setup` 回读（候选测试文件 `test/word-section-layout.test.mjs`，跑轮时先核对它当前断言的范围） | 待跑 |
| 7 | Excel 按行生成公式 | "这一列每行引用本行数据，要能更新" | **保留公式**、引用正确、结果正确 | 公式文本 + `evaluate_formula` 结果（候选测试文件 `test/excel-contract-fixes.test.mjs` / `excel-range.test.mjs`，跑轮时先核对断言语义） | 待跑 |
| 8 | 跨工作簿汇总 | "把这几本的明细汇总成一本" | 源与目的正确、不覆盖错误窗口 | 身份读回 + 非目标簿断言（`test/target-ambiguity.test.mjs` 的 C7「另一个簿绝不被碰」断言，已核对） | 待跑 |
| 9 | 隐藏工具完成任务 | "给形状加个动画" | 查 schema 后通过 `wps_call` 成功 | `test/advertised-surface.test.mjs` 用 `wps_call` 派发隐藏工具（已核对该文件用 `viaCall` 调隐藏工具） | 待跑 |
| 10 | PPT 空白页添加标题 | "空白页上加个标题" | 正确选择文本框或合适版式 | `blank` 页 `set_slide_title` 明确报错 + `add_textbox` 成功（`test/ppt-contract-fixes.test.mjs`，已核对相关断言） | 待跑 |
| 11 | PPT 插入、删除、重排 | "删第 3 页，把第 5 页移到最前" | 最终顺序正确且有回读证据 | 逐页 `get_slide_title` 回读（候选测试文件 `test/ppt-slimming.test.mjs` / `ppt-slide-index.test.mjs`，跑轮时先核对） | 待跑 |
| 12 | 写入后超时 | 受控故障注入：让桥超时 | 先核对结果，不重复追加 | 超时文案含「状态未知」+ `suspect` 短超时（`test/watchdog.test.mjs`，静态；本轮未跑） | 待跑（需故障注入） |
| 13 | 新建未保存文件要求保存关闭 | "保存并关掉"（没有路径） | 不丢失内容；需要路径时明确指出 | `test/close-safety.test.mjs` **本轮已跑 36/36**（Excel/Word/PPT 三侧，含裸 COM 交叉验证与内容仍在的断言） | 待跑（提示措辞层面） |
| 14 | 修改成功但部分格式失败 | 受控：让一步短欠 | 报告已完成与未完成部分 | `partial` / `shortfalls` 通道（`test/warning-channel.test.mjs`、`test/honest-reporting.test.mjs`，本轮均跑过） | 待跑 |
| 15 | 明确目标的小改动 | "把 A1 改成 5" | 不提出无必要的重复确认 | SKILL.md §3 明文写「唯一目标：直接推进，不要为常规选择打断用户」（静态可查；`skill-quality` 断言的是它引用了共享协议，不是这一句） | 待跑 |

## 2. 比较方法（§10.3）

> 可执行的场景表、夹具生成、提示词与**确定性断言**已落到 [test/skilleval/](../test/skilleval/README.md)：
> `node test/skilleval/run.mjs --dry-run` 生成输入夹具与断言清单（不调模型、不需要 WPS），
> `node test/skilleval/verify.mjs <场景key>` 对产物做判定（正例/反例都已验证过会 PASS/FAIL）。
> §1 的 15 条与 `test/skilleval/scenarios.mjs` 一一对应；其中 6 条已有夹具与断言，其余 9 条标着待补原因。

对同一场景、新旧 Skill 各跑一轮，**必须固定**：模型、Harness 版本、工具档位（`WPS_OFFICE_STANDARD` = standard）、
初始文件（同一份 fixture）、温度等采样参数。

每轮记录：

| 指标 | 怎么量 |
|---|---|
| 任务最终成功率 | 按上表的"成功标准"逐项判定；不看模型自述 |
| 错误目标次数 | 非目标文件出现的写入次数（用独立解析/裸 COM 对比基线指纹） |
| 数据丢失次数 | 目标文件内容缺失或原稿被改的次数 |
| 无依据成功声明次数 | 模型声称完成但确定性检查为否的次数 |
| 工具调用总数 / 无效调用数 / `wps_help` 查询数 | 从会话 JSONL 统计 |
| Skill / reference 实际加载量、tokens | 按 Harness 的实际计费/加载读数，**不用字符数估算** |
| 执行耗时、重复询问次数、恢复成功情况 | 会话时间线 |

代表性场景**至少重复三轮**，报告波动区间；优先选成功率和正确性更好的方案，而不是"一轮调用最少"的方案。

> 记录模板：
> ```
> 场景# / 轮次 / 模型 / Harness / 档位 / skill 版本(commit) / fixture 哈希
> 结果：成功 | 部分 | 失败    证据：<文件 + 断言>
> 调用数=…  无效=…  help=…  加载 tokens=…  耗时=…  重复询问=…  恢复=…
> ```

## 3. 验收状态（哪些已验、哪些还没验）

1. **Harness 侧的 Skill 加载与相对路径读取 —— 已实测（2026-10-17，`wpse2e` profile，headless 会话）。**
   结论：① 只加载应用 Skill 时**不会**自动带出总 Skill（返回里只有一句指针）；② 相对 `references/`、
   `recipes/` **能读到**，base directory 明确交付给模型；③ 读不到文件的错误是
   `Error: cannot read "<绝对路径>": not found`；④ 能按路径读另一个已注册 Skill 的文件；
   ⑤ 顺着应用 Skill 的指针加载 `wps-office-next` 成功，正文含「保存、副本与关闭」「没有 documentId」。
   详细证据与原始返回见 [skill-audit.md](skill-audit.md) §4.1。因此四个 SKILL.md 保留关键协议短摘要的
   降级设计**被证实是必要的**；细节放 `references/` `recipes/` **被证实是安全的**。
2. **视觉验收 —— 未做**：文字溢出、对象重叠、裁切需要导出图片 + 图像读取能力；Skill 里写的是
   "读不到图片就说视觉效果未验证"，不是"排版无误"。
3. **模型侧对比（新旧 Skill）—— 已跑第一轮（2026-10-17，见 §4）**：6 个 ready 场景 × 3 轮 × 两臂。
   结论是**正确率打平（18/18 vs 18/18）**，但新臂平均多花 **+32% tokens**、**+38% 墙钟时间**。
   这只覆盖"准备充分"的 6 个场景，**不能**外推成"模型更容易选对工具"；9 个 todo 场景仍未跑。
4. **危险/异常场景**（超时、部分生效）—— **未做**：需要受控故障注入，本轮只做了文本与单元层面的契约断言。

## 4. D1 实测结果（2026-10-17）

**设置**：同一模型（deepseek-flash）、同一 Harness、同一 `WPS_OFFICE_TOOLSET=standard`、同一夹具与提示词；
两臂只换插件代码。新臂 = `wpse2e`（junction → 工作区，代码含本轮全部改动）；
旧臂 = `wpse2e-old`（git worktree 固定在重构前提交 `bc69b96`，独立 profile，串行跑）。
每轮：复制夹具 → headless 跑一次 → 解码会话日志 → `verify.mjs` 确定性判定。单轮上限 60s（超时记 `timedOut`，不计入判定失败）。

| 臂 | 轮次 | 通过 | 成功率 | 平均 tokens/轮 | 平均 WPS 工具调用/轮 | 平均耗时 | 超时 |
|---|---|---|---|---|---|---|---|
| 新 Skill | 18 | 18 | **100%** | 530,319 | 14.7 | 44s | 0 |
| 旧 Skill | 18 | 18 | **100%** | 401,122 | 11.9 | 32s | 2（产物仍正确） |

逐场景（平均 tokens / 平均 WPS 工具调用）：

| 场景 | 新臂 | 旧臂 |
|---|---|---|
| word-new-save | 3/3 · 490,587 · 13.7 | 3/3 · 349,784 · 12.7 |
| two-word-target | 3/3 · 637,103 · 17.3 | 3/3 · 327,639 · 12.0 |
| keep-original | 3/3 · 400,176 · 11.0 | 3/3 · 401,698 · 14.0 |
| excel-row-formulas | 3/3 · 958,837 · 25.0 | 3/3 · 500,745 · 14.0 |
| dirty-save-close | 3/3 · 352,431 · 9.3 | 3/3 · 272,108 · 8.0 |
| small-change | 3/3 · 342,777 · 11.7 | 3/3 · 554,757 · 11.0 |

### 读这张表时必须知道的三件事

1. **第一轮不是这个结果，修法是改接口而不是加说明**。第一次跑 keep-original 时新臂 1/3 失败：
   模型把副本存到 `output/`（目录不存在）→ `save_as` 只回一句「文档保存失败」→ 它去建目录、
   重存、再编辑，最后**声称成功但副本内容没变**（无依据的成功声明）。改法是给桥加一条前置检查：
   输出目录不存在时直接说「输出目录不存在：<路径>」并给下一步。改完重跑该场景 3/3 通过，
   平均 tokens 从 732k 降到 400k。→ 印证改进指南 §11-D「优先改接口，再补必要说明」。
2. **测量被环境打断过，重跑过 4 轮**：WPS 在跑动中被手动关闭 / 超时残留进程占着 xlsx 锁，
   导致 2 轮 `exit=1`、2 轮 60s 超时。清场机制（`--reap-wps` + 超时后 `taskkill /T` + 清场放在动文件之前）
   补上后重跑，结果才稳定。这 4 轮不计入上表（旧结果已重跑覆盖）。
3. **样本很小**：每场景 3 轮、单模型单机，**没有统计显著性**，也没有做重复三轮以上的波动区间。
   结论只能当方向：新 Skill 在这 6 个场景上**没有提高成功率**，代价是多花约三分之一的 token 与时间。

### 解释与下一步

- 新 Skill 更长的流程（身份读回、保存后回读、失败分流）确实带来了更多工具调用与 token —— 这在
  `excel-row-formulas`、`two-word-target` 上最明显。它是否值得，取决于**更难的场景**是否因此少犯错；
  而那 9 个场景（跨工作簿、多节混排、PPT 页序、超时/部分失败）正是本轮没跑的。
- 因此**不据此改回旧文案**，也不据此删减验证步骤；下一步是把 9 个 todo 场景的夹具补上再复测。
- 复现命令：`node test/skilleval/rounds.mjs --profile wpse2e --label new --rounds 3 --reap-wps`，
  旧臂把 `--profile` 换成 `wpse2e-old`；对比用 `node test/skilleval/compare.mjs --new new --old old`。

# Skill 审计表（阶段 A 交付物）

> 依据：`C:\Users\qwer\Desktop\改进.md`（下称「改进指南」）。
> 口径：**以真实实现、运行时 schema 和验证结果为准**。FIXES / HANDOFF / 路线图只用于理解原因，不作为当前行为的证据。
> 本表逐条核对改进指南 §2.3 的清单，并补上本次核对新发现的问题。
> 「核实结果」一列只写我实际读过的代码位置；「尚待实测」明确标出，不写成已支持。
>
> **这是改动前的快照。** 每条的实际落地情况、实测范围与剩余限制见
> [skill-refactor-report.md](skill-refactor-report.md)（§4 逐条对账）；固定场景与比较方法见
> [skill-evaluation.md](skill-evaluation.md)。

## 0. 事实基线（读过的文件与结论）

| 对象 | 位置 | 关键事实 |
|---|---|---|
| Skill 注册 | `plugin.js:36-77` | 四个 Skill 按固定顺序注册：`wps-office-next` / `wps-excel` / `wps-word` / `wps-ppt`；`resourceBase` 是各 Skill 自己的目录（`{kind:'directory'}`）；frontmatter 只用 `name/description/whenToUse` 三个标量键 |
| MCP 接线 | `cordis.patch.yml` | stdio、`WPS_OFFICE_TOOLSET` 默认 `standard`、`toolCallTimeoutMs: 330000`、`failOnStartupError: true` |
| 广告面 | `mcp/src/server/toolset.ts` | `FACADE_TOOLS` 4 个常驻；`STANDARD_TOOLS` 里 Excel 46 / Word 27 / PPT 14 / 通用 4；`selectTools` 只影响 `tools/list`，注册表不动 |
| 门面语义 | `mcp/src/server/mcp-server.ts:237-464` | `wps_status` 报 connected/toolset/计数；`wps_help` 无参概览、`app` 目录、`query` 搜索、`tool` 取 schema；`wps_call {tool,args}` 派发；`wps_batch {calls[]}` **顺序执行、无 stopOnError、无整体状态** |
| 参数校验 | `mcp/src/server/tool-registry.ts:305-367` | 必填缺失 → `Missing required parameter`；**未知顶层键 → `未知参数: X（动作：…）可选参数: …`**；只做**顶层**形状检查，**不递归**嵌套对象 |
| 结果语义 | `mcp/src/server/tool-registry.ts:243-300` | `warnings`=尽力而为失败（success 仍可为 true）；`shortfalls` → `success:false` + `partial:true`（R11 不变量） |
| 目标解析 | `mcp/scripts/wps-com.ps1` `Get-TargetPres` / `Get-WorksheetByParam` / `Get-ActiveWordDocument` | **没有 documentId / 目标句柄**；Word 侧写操作落在活动文档，Excel/PPT 侧可传 `sheet` / `presentationName`，未传即活动对象并发 `warnings` |
| 启动 WPS | `mcp/scripts/wps-com.ps1` `Get-WpsApp`（L93-137） | 业务动作（open/create）走 `Get-WpsApp` → 没找到已注册实例就 `New-Object -ComObject` **真的启动 WPS**；`wps_status` 走只读的 `Get-WpsAppRegistered`（L139-154），不会启动任何东西 |
| 桥生成物 | `scripts/build-host-actions.ps1` | `host/wps-actions.ps1` 由 `mcp/scripts/wps-com.ps1` 生成；桥里改语义必须改源文件再重新生成 |
| 租约 | `host/wps-com-host.ps1:42-156` | 命名互斥体 + `%USERPROFILE%\.wps-office-mcp\com-host.json`（hostPid / clientPid / startedUtc / updatedUtc）；`$script:StaleHeartbeatSeconds = 120` |
| 孤儿回收 | `mcp/scripts/wps-com.ps1:190-267` | `owned-apps.json` 只记 **应用类别**（`apps=[excel,word]`）+ 宿主/客户端 pid；回收时按类别 `GetActiveObject(progId)` |
| 工具目录 | `spec/tool-definitions.json` | 注册工具 270 个，其中 Excel 118 / Word 61 / PPT 77 / 通用 7 / 转换 2 / 门面+内置 5 |
| Skill 生成 | `scripts/gen-skill-tools.mjs` | `reference.md` 整份覆盖；`SKILL.md` 的 `GENERATED:advertised` 标记块被**整块重写**，内容是全量工具名清单 |

## 1. 改进指南 §2.3 逐条核对

| # | 现有说法（位置） | 对应实现 | 核实结果 | 处置 | 验收方式 |
|---|---|---|---|---|---|
| 1 | 总 Skill：「**可选参数写错名不会被发现**：调用照常执行、静默使用默认值」（`skills/wps-office-next/SKILL.md:26`） | `tool-registry.ts:335-351`：`properties` 里没有这个键就 `throw InvalidParamsError('未知参数: …')` | **过期／矛盾**。未知顶层参数现在会被**响亮拒绝**，并列出 accepted 键与「是不是想传 X」提示。FIXES 99（W6-4）就是修这条的 | 重写：改成当前规则；**只承诺顶层**，明确嵌套对象不做递归校验 | 静态：`test/arg-shape-guard.test.mjs` + 新增 skill 质量测试；接口：`wps_ppt_add_textbox {x:1}` 应报未知参数 |
| 2 | Word Skill：「`close_document` **不在广告位**，用 wps_call」（旧版 `skills/wps-word/SKILL.md`） | `toolset.ts:120` 明确列了 `wps_word_close_document`，`advertised-surface` 与 `reference.md` 也标「直达」 | **已不复现**（工作区未提交改动里已改正）。手写广告状态本身就是漂移源 | 删除手写广告状态：SKILL.md 不再声明谁在广告位，只指向自动生成的 `reference.md` | 静态：生成器 + `test/advertised-surface.test.mjs` 与 `spec/advertised.json` 对账 |
| 3 | 总 Skill：「大改动前先 `wps_common_save`，**保留可回退版本**」（`skills/wps-office-next/SKILL.md:48`） | `wps_common_save` → 桥 `save` 动作写回**当前文件**（覆盖磁盘原件）；没有任何版本/备份能力 | **误导**。保存 ≠ 备份，保存会覆盖磁盘上的原件，反而不提供回退点 | 重写为「保存 / 另存副本 / 备份」三分：要保原件就先另存工作副本再改；不默认每次修改都保存 | 静态：skill 质量测试断言不出现「保留可回退版本」类表述；场景：改报告保留原件 |
| 4 | 总 Skill：「调用 `wps_status` 确认已启动……**没连上就不要继续**」（`skills/wps-office-next/SKILL.md:13`） | `Get-WpsApp`（L93-137）在业务动作里会 `New-Object -ComObject` 启动 WPS；`wps_common_open_workbook` / `create_*` 都走它 | **过期**。`connected:false` 只说明「此刻没有已注册实例」，不说明任务做不了；打开/新建会拉起 WPS | 重写：状态查询降级为**诊断**；目标明确的打开/新建任务直接推进 | 静态：skill 质量测试断言不出现「没连上就不要继续」；场景：WPS 未启动 → 新建并保存 |
| 5 | PPT Skill：「插入图片用 **path**」（`skills/wps-ppt/SKILL.md:25`） | `spec/tool-definitions.json`：`wps_ppt_insert_ppt_image` 的参数是 `slideIndex, imagePath, left, top, width, height`；桥 `insertPptImage` 读 `$p.imagePath` | **错误**。没有 `path` 参数；传 `path` 会被顶层校验拒绝 | 删除该行，改为「图片一律 `imagePath`，输入文档才是 `path`」并从运行时 schema 生成示例 | 静态：skill 质量测试对「文档里出现的参数名必须在 tool-definitions.json 的 schema 里」做核对 |
| 6 | PPT Skill：「没有 x、y 参数——**传了不会被发现**（可选参数拼错是静默忽略）」（`skills/wps-ppt/SKILL.md:24`） | 同 #1 | **矛盾**。与当前 `未知参数` 拒绝机制直接冲突 | 删除旧失败行为描述，改成「按 schema 的 `left/top/width/height` 传」 | 同 #1 |
| 7 | PPT Skill：「**演示是三个应用里工具最多的一个**」（`skills/wps-ppt/SKILL.md:19`） | 注册工具：Excel 118 / PPT 77 / Word 61 | **不成立**。无任务价值且必然漂移的比较 | 删除 | 静态：skill 质量测试断言不出现该比较句（改为由生成器提供计数） |
| 8 | PPT Skill：已移除的高层工具（`create_kpi_cards` / `apply_color_scheme` / `auto_beautify_slide`）「又出现在已知坑」 | `spec/tool-definitions.json` 里这三个名字**都不存在**；`wps_ppt_beautify`、`set_shape_effect`、`insert_ppt_chart`、`set_ppt_chart_data`、`set_ppt_chart_style` 都存在 | **已不复现**（工作区未提交改动已把「已知坑」改写成存在的工具）。现状：`已知坑` 提到的工具名全部存在 | 保留现状，并把该不变量固化成静态检查 | 静态：新测试断言 `skills/**/*.md` 里出现的每个 `wps_*` 名字都在注册目录里 |
| 9 | Word Skill：「先给标题行单独设字体或样式，**再统一正文字体**」（`skills/wps-word/SKILL.md:51`） | `set_font` 的 `range` 接受 `{start,end}` 或 `all`；`all` = `$doc.Content`（`applyStyle` 同） | **顺序有坑但结论不完整**：若最后一步用 `all`，会把标题一起覆盖 | 重写：需要全局基线时**先全局、后局部覆盖**；并给出「读 `get_paragraphs` → 只对正文各段坐标设字符格式」的明确范围做法 | 静态：新测试断言 Word 正文里给出「先全局后局部」或「只对正文坐标」的明确顺序；场景：排版已有报告 |
| 10 | Excel Skill：「需要逐行递增的引用就**逐格调用，或直接用 `write_range` 写入数值**」（`skills/wps-excel/SKILL.md:61,73`） | `set_formula` 多格写入同一个公式字符串（`wps-com.ps1` `setFormula`：多格逐格 `Cells.Item(r,c).Formula = $p.formula`，**不做相对引用调整**） | **工具行为属实**，但「改写成静态数值」会丢掉用户要的可更新能力 | 重写：用户要公式就保留公式——逐格 `set_formula` 或 `fill_series`；写回后回读**公式文本 + 计算值**两根证据；不把静态值当默认建议 | 场景：Excel 按行生成公式（保留公式、引用正确、结果正确） |
| 11 | 收尾流程：默认关闭文档（Word Skill:「别把没保存的文档留在用户机器上」） | `close_document` / `close_workbook` / `close_presentation` 默认 `save=true` | **与用户意图可能相反**：用户可能还要继续编辑 | 重写：以用户任务决定是否关闭；**保存不必关闭**；只有用户要求关闭才关 | 静态：新测试断言收尾流程写明「由任务决定」 |
| 12 | 错误处理：「返回 unknown parameter(s)：按错误里列出的 **accepted** 名字重试」（`skills/wps-office-next/SKILL.md:76`） | 这条错误来自**工具层 → 桥**（handler 拼给桥的键），`accepted` 是桥接受的键；调用方改自己的参数**修不了**它 | **误导**：把插件缺陷说成调用错误，会让模型反复改无效参数 | 重写失败协议：分清「调用参数错误（按 schema 改一次）」/「插件内部映射错误（记录 action + 原始错误，换路径或上报）」，禁止猜别名反复试 | 静态：新测试断言失败协议含此分流；场景：参数桥接报错 |
| 13 | 「新文档 `save=true` 仍丢弃关闭」（`wps-com.ps1:5341-5346`、`5489-5492`；PPT 侧同类） | `closeWorkbook`/`closeDocument`：`$saveWanted && $path -eq ""` → 强制 `saveChanges=$false`，**仍然关闭**，只回一条 `warning`，`success:true` | **确认存在**：未落盘内容被丢弃，调用方只看 `success` 会以为「要求保存的关闭」成功了 | **先修代码**（P0）：无路径时不关闭、保留文档打开，返回明确错误与下一步；再改 Skill 文案 | `test/close-safety.test.mjs`（静态文本门禁）+ 真机场景「新建未保存文件要求保存关闭」 |

## 2. 本次核对新发现（改进指南未列，但影响 Skill 正确性）

| # | 现有说法（位置） | 对应实现 | 核实结果 | 处置 | 验收方式 |
|---|---|---|---|---|---|
| 14 | 总 Skill：「未广告的工具**也可以按全名直接调用**，wps_help 只负责让你发现它们」（`skills/wps-office-next/SKILL.md:40`） | 模型能调用的名字来自 `tools/list`；`wps_call` 是唯一的隐藏工具入口。FIXES D4 的取证里，S6 attempt 1 直接调 `wps_excel_freeze_panes` 得到的是 `unknown tool` | **误导**。服务端 `tools/call` 接受注册名 ≠ 模型有直接调用入口 | 删除该条，改为「未广告工具只能用 `wps_call` 调用」 | 静态：新测试断言该句不存在 |
| 15 | 总 Skill：单位表「演示 `add_shape`/`add_textbox`/`set_shape_position` **像素**」（`skills/wps-office-next/SKILL.md:66`） | 桥把参数直接交给 `Shapes.AddShape/AddTextbox(type,left,top,w,h)`——PowerPoint COM 几何单位是**磅**；只有 `setSlideSize` 显式 `* 0.75`（像素→磅），`set_slide_size` 的 schema 注释也用像素（960×720 = 4:3） | **混合单位**：形状几何是磅、幻灯片尺寸是像素、导出图尺寸是像素。笼统写「像素」会让布局算错 | 重写单位表：分「形状几何=磅」「画布尺寸=像素」「1px = 0.75pt」；详细换算进 `references/layout-and-shapes.md` | 场景：PPT 空白页加标题 + 排版；静态：新测试断言单位表区分磅/像素 |
| 16 | 总 Skill：「结果里出现 `warnings`：本次操作**已经完成**」（`skills/wps-office-next/SKILL.md:79`） | `warnings` 与 `shortfalls` 是两种语义：R11 下 `shortfalls` 会把 `success` 判为 `false` 并加 `partial:true` | **不完整**：只看 `success` 的调用方会漏判；文档没教怎么区分 `warning` / `shortfall` / `partial` | 重写验证与失败协议：读 `success` + `partial` + `shortfalls` + `warnings` 四件事，并给出各自该报什么 | `test/honest-reporting.test.mjs`、`test/warning-channel.test.mjs` + 新测试断言四件套被写进 Skill |
| 17 | 总 Skill：「多个连续操作可用 `wps_batch` 一次提交，最多 50 项」（`skills/wps-office-next/SKILL.md:39`） | `mcp-server.ts:439-461`：顺序执行、遇错继续、逐项 `{tool,success,result}`，无整体状态、无 stopOnError | **能力说少了也误导**：依赖步骤（先建表再写数据）在 batch 里失败后会继续，后续步骤可能写错位置 | 短期：Skill 明确「batch 只用于彼此独立、失败不会误写的操作；有依赖就逐项检查」；**代码补 `stopOnError` 与整体状态**后统一更新文案 | `mcp-server.ts` 单测（新）+ 场景：批量失败不误写 |
| 18 | 总 Skill：「目标明确时……」（无 documentId 相关内容） | 全仓 `documentId` 无匹配；目标靠「活动对象 + 可选名字参数」 | **尚无能力**。不能写「已锁定目标」 | 阶段 C 实现并验证**最小可用的目标标识**（读回身份 + 执行前后核对），验证通过后再写进 Skill；否则只写「仍依赖活动对象，调用间可能被焦点切换影响」 | 真机：两份 Word 同时打开，明确目标、不误改另一份；`test/target-ambiguity.test.mjs` |
| 19 | 总 Skill：「操作超时：常驻宿主会自动重启」（`skills/wps-office-next/SKILL.md:85`） | `com-host.ts:164-178`：超时→杀宿主、`suspect` 短超时、错误文案明说「状态未知」 | **说少了**：没提「非幂等操作不得盲目重放」——插入文本/加页/追加行超时后原样重试会重复写入 | 重写失败协议：超时=状态未知，先核对后置条件再决定是否重试；非幂等操作禁止盲目重放 | 场景：写入后超时先核对不重复追加 |

## 3. 阶段 C 待修代码（先实现与验证，再写文案）

| 优先级 | 问题 | 位置 | 处置方向 |
|---|---|---|---|
| P0 | 未落盘文档 `save=true` 被丢弃关闭（#13） | `mcp/scripts/wps-com.ps1` `closeWorkbook` / `closeDocument` / `closePresentation` | 无路径时不关闭、保留打开，返回 `success:false` + 明确下一步；改完重新生成 `host/wps-actions.ps1` |
| P0 | 租约仅因心跳 >120s 就判定「仍存活的宿主」为过期并杀掉 | `host/wps-com-host.ps1` `Get-StaleOwnerReason` | 身份核对通过（`Test-OwnerAlive`）的宿主**不因心跳**判过期；心跳只作诊断/兜底 |
| P0 | 孤儿回收只记录/核验**应用类别**，不核验实例身份 | `mcp/scripts/wps-com.ps1` `Save-WpsOwnedApps` / `Invoke-WpsOrphanReclaim` | 记录实例身份（pid + 启动时刻，拿不到就退化为不回收），回收前核对身份 |
| P1 | `wps_batch` 无 `stopOnError`、无整体状态 | `mcp/src/server/mcp-server.ts` | 加 `stopOnError` 与 `status: completed/partial/failed`、`stoppedAt`；默认行为不变（保持兼容） |
| P1 | 目标标识缺失 | 桥 `Get-TargetPres` / `Get-WorksheetByParam` / Word 活动文档 | 实现可读回的目标身份（名称 + 路径），并在写操作结果里回报；**验证通过后**才写进 Skill |
| P1 | 结果状态不统一 | `tool-registry.ts` / 桥 | 已有 `success` / `partial` / `shortfalls` / `warnings`；补一条 `status` 便于程序化消费 |

## 4. 尚未实测、必须标注的条目

- 应用 Skill 被单独选中时，总 Skill 是否自动加载、相对路径 `references/` `recipes/` 是否可读：**本环境没有 DSH 交互会话（只有静态文件系统）**，因此按改进指南 §3.2 的降级方案处理——每个应用 SKILL.md 保留极短的共享协议摘要（目标确认 / 保存策略 / 验证 / 失败分流），关键规则不放进「可能读不到」的文件。
- 视觉验收（文字溢出、重叠、裁切）：需要导出图片 + 图像读取能力，本环境未验证，Skill 里明确写「未验证视觉效果」而不是「排版无误」。
- 新旧 Skill 的模型侧对比（成功率 / 调用数 / tokens）：见 `docs/skill-evaluation.md` 的固定场景集；**本轮只固定场景与成功标准，不声称已跑过**。

## 5. 审计中发现、清单之外的额外缺陷（本轮一并修掉）

核对过程中撞到两个不在改进指南 §2.3 表里、但会直接破坏"Skill 说得对"的底层缺陷：

| 缺陷 | 位置 | 为什么它属于本轮 |
|---|---|---|
| **归属记录里的中文文件名被 ANSI 解码成乱码**（`文字文稿2` → `鏂囧瓧鏂囩2`），导致实例身份永远对不上、孤儿实例永远回收不掉 | `mcp/scripts/wps-com.ps1` 的 `Invoke-WpsOrphanReclaim`（`Get-Content` 无 `-Encoding UTF8`） | 它是"核验实例身份"这条 P0 修复的伴生缺陷：不修，修复本身就等于关掉回收能力。已修，并有真机测试 |
| **正式客户端按 `chunk.toString()` 收帧**：chunk 边界切开多字节中文字符会变成 U+FFFD | `mcp/src/client/com-host.ts` | 改进指南 §9 P1 明确列出（"正式客户端使用字节收帧或 StringDecoder / 不要求模型回避中文"）。已修，并有逐字节切分的静态测试 |

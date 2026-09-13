# 进度

MVP 目标（P0-P3 + 最小技能集）已全部完成并在真实 DSH 上端到端验证。

| 阶段 | 状态 | 证据 |
|---|---|---|
| P0 冻结基线 | 完成 | baseline/upstream-0.1.0/：250 工具、141,872 schema 字节、约 40.5k tokens/请求、248 个 COM action；baseline/known-defects.md（24 条） |
| P1 合仓 + Windows 化 | 完成 | mcp/ 迁入 43 个 TS 源文件；删除 mac 传输、keepalive、平台分支、axios |
| P2 常驻 COM host | 完成 | warm ping 1ms（原 969ms）；test/com-host.test.mjs 6/6 |
| P3 工具面收敛 | 完成 | standard 43 工具 / 23,198 字节 / 约 6.6k tokens；scripts/verify.mjs 22/22 |
| 最小技能集 + DSH 自举 | 完成 | plugin.js + cordis.patch.yml；4 个技能；test/plugin.test.mjs 32/32；真实 profile 启动验证通过 |
| P4 缺陷修复 | 进行中 | 已修：11 个永久失败的工具、closePresentation、sheet 省略失败（6 处）、范围逐格读、3 对重复工具合并；证据 new-actions 28/28 + excel-range 10/10 + deprecated 8/8。新发现 PPT 动画/切换整族不可用（传英文名，COM 要数字枚举） |
| P5 全量技能打磨 | 未开始 | 当前为最小技能集 |

## DSH 端到端验证（真实 profile，非模拟）

步骤：新建 headless 测试 profile → `dsh plugin --profile wpstest add D:\dsh\a` → 启动并让模型调用 wps_status。

服务端日志证据：

- `Registered facade tools {"tools":["wps_status","wps_help","wps_call","wps_batch"]}`
- `Returning 31/254 tools (mode=standard)`（默认档位在真实环境生效）
- `Starting resident COM host {"script":"D:\\dsh\\a\\host\\wps-com-host.ps1"}`（路径自举成功，无需环境变量）
- `Resident COM host ready {"pid":26672}`
- `[RESPONSE] ping - SUCCESS (1082ms)`、`getAppInfo - SUCCESS (97ms)`
- `Tool executed: wps_status ... success: true`

客户端侧：模型推理中明确出现 wps-office-next 技能，最终回答为该工具返回的 connected 字段值 true，进程退出码 0。
测试 profile 已在验证后删除，可用 README 中的一条命令重建。

## 工具面实测

| 档位 | 工具数 | schema 字节 | 约 tokens |
|---|---|---|---|
| 基线（收敛前） | 250 | 141,872 | 40,535 |
| full | 254 | 143,217 | 40,919 |
| standard（默认） | 43 | 23,198 | 6,628 |
| minimal | 4 | 1,348 | 385 |

## 预算取舍说明

最初定的 12,000 字节上限无法在不损失信息的前提下达到：实测把全部参数描述删掉，45 个工具也只能降到
15,148 字节——成本主体是 schema 结构本身（属性名/类型/枚举/required），约 527 字节/工具，而参数描述
承载了单位与枚举（例如 PPT 的像素 vs 磅），正是上游缺陷的来源之一。因此改为收缩工具数量：
27 个精选工具 + 4 门面 = 31，门禁设为 40 工具 / 22,000 字节。后续按用户指示把预算放宽到 25,000 字节，
精选工具相应补到 39 个（合计 43），门禁改为 45 工具 / 25,000 字节。

## 后续（MVP 之外）

- P4：按 baseline/known-defects.md 修复 24 条缺陷，重点是 50 处静默 catch 与结果形状统一。
- P5：把最小技能集扩展为全量技能，并把工具目录生成接入 CI 防漂移。
- 分发：提交身份已设为 sueccku；gh 2.100.0 已装并完成登录，仓库已推送。
## P4 修复进度

已修完（详见 docs/FIXES.md）：

1. 11 个工具在 Windows 上永久失败——上游引用了 10 个桥里根本不存在的 action。
   已全部实现并在真实 WPS 上逐项回读验证；其中 wps_ppt_set_shape_fill 原本就在
   standard 档里被广告，属于用户直接可见的坏功能。
2. closePresentation 在 saveChanges=false 时丢弃修改会弹保存框，已修。

待修（本节列的其余项）：

- ~~合并重复工具~~ 第一批 3 对已合并；其余 15 对参数接口不同，待参数契约测试后再处理
- 参数校验只检查 required，写错参数名静默忽略
- 统一结果形状、消灭 50 处静默 catch
- 文档与代码不一致（auto_fit 声明但不存在；tools/index.ts 计数有误）
- ~~高层场景封装（KPI 卡片/时间线/流程图等）从工具层下沉到技能层~~ 已完成（FIXES 第 22 条）：19 个封装摘除，技能改为组合配方

## P4 修复进度（PPT）

参数契约 sweep 的 A/B/C 三类全部归零（212 对工具/action）：A=工具发了桥不读、B=schema 宣称 handler 不用、
C=嵌套对象里多写属性。详见 docs/FIXES.md 第 18 条。

PPT 侧新增两张由生成器产出的表——别名表（20 条，公开名→规范键）与容器表（13 条，嵌套对象展开），
并据此修好此前完全不可用的切换效果/动画整族（名称→数值映射）、3D 旋转类型、背景对象、类别序号等。

同时删掉 5 个无法实现的参数（环形图多段 data、组织架构 data、图表 data、渐变 angle/type、页码 startFrom），
原因都写进 schema 与文档。新增两道门禁：生成器拒绝写出无法解析的模块；verify 断言 action 数量三方一致。

已完成：test/ppt-contract-fixes.test.mjs 42 项全绿并提交；技能文档四份已同步（第 19 条）。

待办（原「11 个 handler 实参不可知」已澄清并大部分关闭，见第 20 条）：

## P4 修复进度（续）

第 12～14 条已落地，详见 docs/FIXES.md：

3. 工作表操作组（delete/rename/copy/move/switch）作用在活动表上，且 0 基 position 契约自相矛盾。
4. **数据丢失级**：Word 纯查找会删掉所有命中内容，且次数是编造的。
5. **系统性闸门**：桥侧自动推导每个 action 真正读取的参数名，出现多余键即明确报错；
   配套 scripts/param-contract.mjs 零副作用对账全部工具，结果见 docs/param-contract.md。
   闸门上线当天即抓出 findReplace/replaceMode 并量出其余 85 处静默忽略。
6. 顺带修掉最大一类：58 处 Excel action 无视 'sheet'、一律操作活动表。
7. 关闭工具：工具发 'save'、桥读 'saveChanges'，导致 save=false 被丢弃并弹模态框
   （此前两次泄漏 78/85 个工作簿的根因）；现在三个 close action 都做了「未落盘文档不弹框」处理，
   并如实回报 saved 与 warning。

当前实测：151 项测试 + 22 项门禁全绿；参数契约对账 212 对，剩余 85 处错配已列成清单。

待修：

- ~~docs/param-contract.md 里的 85 处「工具发了、桥不读」错配~~ 已全部清零（A/B/C/D 四类均为 0）
- ~~参数校验只检查 required~~ 已由第 13 条的参数名闸门覆盖
- 统一结果形状、消灭 50 处静默 catch
- 文档与代码不一致（auto_fit 声明但不存在；tools/index.ts 计数有误）
- ~~高层场景封装（KPI 卡片/时间线/流程图等）从工具层下沉到技能层~~ 已完成（FIXES 第 22 条）：19 个封装摘除，技能改为组合配方
- 其余 15 对重复工具（参数接口不同）

## P5 进度（第一次真实 e2e 与随之而来的修复）

2026-09-13 跑了第一次真实端到端（headless profile + 跨 Excel/Word 任务）：72 秒、exit 0、
两个产物经独立 COM 复核无误、收尾无残留文档。完整复盘与它暴露的 5 个缺陷见 docs/FIXES.md 第 24～29 条。

本轮之后的状态：

- 工具目录 **237**（广告 **44** / 23,593 字节，上限 45 / 25,000），桥 action **259**，弃用 **18**；
- 测试 **310 项**（15 个文件，含新增 test/word-lifecycle.test.mjs 18 项）+ verify **23 项** 全绿；
- 参数契约对账 **211 对**，A（工具发了桥不读）/ B（schema 宣称不用）/ C（嵌套多写）/ D（透传多写）**四类均为 0**；
  余下 1 处「桥无键表」（setCellFormat 动态键）与 9 处「handler 实参静态读不出」在报告里逐名列出；
- baseline/known-defects.md 的 28 行已逐条标注 已修 / 部分修 / 仍开 / 不修，依据是当前源码而不是修复日志。

顺带记下一条方法教训：会话日志是**追加式拼接的 zstd 帧**，Node 的 zstdDecompressSync 只解第一帧，
复盘工具必须按帧边界切分，否则只看得到 1 条事件。

e2e 已一键化（第 30 条）：`node scripts/e2e.mjs --profile <name>` 自己造 fixture、跑真实 headless 任务、
逐帧解会话日志、用裸 COM 复核产物，并断言没有残留文档 / 没有缺陷标记 / 模型没有自己写 COM 脚本。
实测 19 项检查 65 秒全绿，并做过负向验证（超时压到 5 秒 → 10 项 FAIL、exit 1）。

剩余待办：

- 其余 15 对重复工具（参数接口不同）；
- 12 个 builtin 与 pro 工具重复（未广告，但仍出现在 wps_help 目录里）；
- ~~只跑静态门禁的 minimal CI~~ 已落地（FIXES 第 31 条）：`.github/workflows/ci.yml` 在 windows-latest 上跑
  tsc + `mcp/dist` 漂移、宿主生成 + 生成物漂移、技能表重生成 + 漂移、`verify --static`（18 项）、
  参数契约 + 报告漂移、以及 `plugin.test.mjs` 与 `com-host.test.mjs`；全部步骤已在本机逐条干跑通过。
  首次运行（2026-09-13，run 34760241577）：**13 步全绿、67 秒**，三处产物漂移检查在干净 runner 上复现。

## P0 清理完成（2026-09-13）

按 docs/tool-roadmap.md 的第一阶段执行完毕：

- 桥 action **259 → 231**（删 28 个 case 条目 / 1039 行：22 个场景死代码、3 个点号重复实现、
  4 个重复标签的后一份）
- 注册工具 **237 → 209**（删 11 个 builtin + 18 个废弃名的重复定义，共 1319 行；
  废弃名改为**派发期别名**，不占注册位，旧名字仍可直接调用）
- 生成器别名表 20 → 17、容器表 13 → 12；`mcp/src/tools/index.ts` 头部那份长期过期的工具枚举改成指针
- 恢复 `wps_ppt_set_slide_theme`（D3 把「主题」列为保留项），schema 按桥的真实语义写成「模板文件路径」
- 三档预算：minimal 4 / 1,348；standard **44 / 23,593**；full **209 / 121,308**
- 回归：310 项测试 + verify 23 项 + 参数契约 197 对（A/B/C/D 均为 0，UNPARSED 9 → 6）+
  一键 e2e 19 项 58 秒 + CI 全绿

一条踩坑记录：删生成器表项不能只看「名字是否出现在重复 case 标签里」——`set3DRotation` 的标签重复，
但 action 是活的（工具发嵌套 `rotation` 对象），删错当场被 PPT 契约测试和参数契约 A 类抓住。

下一步：**P1 契约真源**（spec + 生成器，验收标准是逐字节复现今天的产物）。

## P1 第一步完成（2026-09-13）

操作规格（spec）成为工具面的真源，见 docs/FIXES.md 第 33 条与 docs/tool-roadmap.md 的 P1 进展行。

- `mcp/src/spec/` —— 209 条操作（202 bridge / 5 门面 / 1 纯 JS / 1 opaque），由 `scripts/extract-spec.mjs`
  从今天的事实 bootstrap
- `scripts/gen-tool-surface.mjs` —— spec → `spec/{tool-definitions,action-keys,advertised,signatures}.json`
- `test/spec-reproduction.test.mjs` —— **P1 验收：9 项全绿**，其中
  209 个 schema 深度相等、序列化字节数 **121,308 = 121,308**、广告集 44 一致
- CI 增加三步（重生成工具面 → spec 漂移 → 复现验收），在 5f5a0f2 上 48 秒全绿

**入账的债务**（账本式门禁：涨了会红、降了要手动改数字）：
重命名 58 处（参数名 ≠ 桥键，分布在 33 个 action）· 未工具化 action 29 个（P2 待办清单）·
raw schema 片段 32/549 · 带别名工具 15 · 带容器工具 12。

- **P1-3（FIXES 34）**：调用点解析器抽成 `scripts/lib/tool-action-map.mjs`（param-contract 与提取器共用，
  抽完验证契约报告逐字节未变）；提取器改用「handler 实际发送的键」反推桥键，重命名债务 **57 → 27**，
  棘轮同步收紧到 27。剩下 27 处是同义改名（`filePath`→`path`、`style_name`→`style`），属 P1-4 范围。

- **P1-4a（FIXES 35）**：给每个参数一个明确去向（bridge 509 / local 22 / 折进单参数 3，**未归类 0**），
  并把「每个 bridge 参数都落在桥真正读的键上」变成**严格断言**；新增棘轮：别名 62（目标 0）。
  验收从 9 项扩到 **11 项**，工具面仍逐字节复现。

- **P1-4b（FIXES 36）**：把生成器里两张手写表（别名 17 条、容器 12 条）搬进 `mcp/src/spec/aliases.ts`，
  由生成器产出 `spec/param-aliases.json`/`param-containers.json`，宿主生成器改读它们——
  **宿主产物逐字节未变**（`git diff` 空 + 计数一致），证明搬家不改变行为。验收扩到 **12 项**，
  新增不变量「每个 pass-through 改名都要在桥侧有声明」（16 处检查过、2 条兼容拼写有账）。
  三张人手表：aliases ✅、containers ✅、`$helperKeys` 仍在（随 P1-5 键表改造消失）。

- **P1-5（FIXES 37，P1 收尾）**：最后一张手写表 `$helperKeys` 也进 spec（`mcp/src/spec/bridge-helpers.ts`
  → `spec/param-helpers.json`），宿主产物**逐字节未变**；同时给"读不出参数的 action"上闸——
  例外必须写进 spec 的 `dynamicParamActions`，未声明就让生成器失败（负向验证过）。
  **三张人手表全部归零**，P1 完成。

## P2 Excel 做深（第一波已落地）

- 广告面 **44 → 48 工具 / 25,097 字节**：`get_sheet_info`、`auto_fit`（+columns/rows）、`set_wrap_text`、
  `find_in_sheet`、`get_named_ranges` / `delete_named_range`
- 注册工具 **209 → 217**，操作规格（spec）同步 217 条；桥 action 仍 **231**（本波没有新 COM 代码）
- **验收**：`test/excel-missing-halves.test.mjs` 18 项（真实 WPS），覆盖单格 / 显式范围 / 缺省已用范围三条路径
- **两个 action 在常驻宿主里从来不可能生效**（`Find().Address()` 不可用；二维组下标被逗号优先级解析成
  `int + Object[]`，异常又被裸 `catch { continue }` 吞掉）——详见 FIXES 第 38 条
- 参数契约：被校验的对 **200 → 205**，A/B/C/D 仍全 0，`UNPARSED` 回到基线 6
- 预算门禁按已锁定的 **D1** 同步为 60 工具 / 32,000 字节（verify 与 deprecated 测试同步）

## P2 第一波余项（已落地）

- 再挂 **10 个**：`copy_format`、`clear_formats`、条件格式读/删、数据验证读/删、`refresh_links`、
  `consolidate`、`calculate`、`group_columns`；刻意不挂 `getActiveWorkbook`（与 `get_sheet_info` 重复）
  与 `unfreezePanes`（已被 `freeze_panes { freeze: false }` 覆盖）
- 广告面 **48 → 51 工具 / 26,652 字节**（3 个进精选档）；注册 **217 → 227**；桥 action 仍 **231**
- **又发现两个从未生效的 action**：`consolidate` 与 `subtotal` 的函数常量用了假的枚举值
  （9/2/1/4/5，真值是 -4157/-4112/-4106/-4136/-4139）；`subtotal` 甚至已经注册成工具却零测试覆盖
- **WPS 差异（裸 COM 量出）**：`Range.Consolidate` 只认 R1C1 且表名必须带引号，A1 引用静默不写；
  桥里新增 `ConvertTo-ConsolidateSource` 做转换，并把 `topRow` 默认改成 false（按位置相加）
- 验收：`test/excel-missing-halves-2.test.mjs` **30 项**（真实 WPS）；全套 **370 项 / 18 文件**、
  verify 23、spec 复现 12、参数契约被校验的对 205 → 215，A/B/C/D 仍全 0
- 未工具化 action 台账 21 → **11**，剩下的是刻意的重复实现（`openFile`/`replaceInSheet`/`unfreezePanes`）
  与 P3/P4 的 Word/PPT 项（`getBookmarks`/`getComments`/`insertHyperlink`/`getActivePresentation`…）

## P2-2 表（ListObject）全族（已落地）

- **新 COM 代码**（P2 里第一次）：桥 action **231 → 239**、注册工具 **227 → 235**、8 个 action + 8 个工具
  （建表 / 读结构 / 加行 / 删行 / 改名与样式 / 总计行 / 调整范围 / 转回区域）
- 广告面 **51 → 54 工具 / 28,793 字节**（create / get / add_row 进精选档）
- 裸 COM 先量支持面：WPS 的 ListObject 完整可用（含 `=SUBTOTAL(109,[Amount])` 这种结构化引用）；
  `XlTotalsCalculation` 实测 1=sum … 8=var（9/10 报错）
- **WPS 差异**：`ListRows.Delete()`/`Add()` 之后同一个表对象仍返回旧几何，必须重新解析一次再报告
- 验收 `test/excel-list-object.test.mjs` **25 项**（真实 WPS）；全套 **395 项 / 19 文件**、verify 23、
  spec 复现 12、参数契约被校验的对 215 → **223**，A/B/C/D 仍全 0

## P2-3 页面设置 / 打印 / 外观 / 公式审计（已落地）

- 侦察发现 **Excel 此前完全没有页面设置能力**（桥里的 setPageSetup / insertPageBreak 都是 Word 的），
  于是这一波补 8 个 action + 8 个工具：读设置快照、页面设置、打印标题、页眉页脚、
  工作表可见性与标签色、分级显示层级、清除分页符、公式审计（引用来源/被引用）
- **刻意不做打印与打印预览**：前者是物理副作用，后者会开模态窗口卡住常驻宿主——改成“把设置调好、
  由人来打印”（要的话再单独评估）
- 测试抓到两处**成功但说错**：清掉标签色后 Tab.Color 读回 0（改用 ColorIndex=-4142 判“默认”）；
  常量格的 Range.Formula 返回的是值本身（改用 HasFormula 判“是不是公式”）
- 桥 action 239 → **247**、注册 235 → **243**、广告面 54 → **56 工具 / 30,885 字节**
  （D1 上限 32,000，只剩 1,115 字节余量 → P5-2 的重定已不是可选项）
- 验收 `test/excel-page-setup.test.mjs` **24 项**（真实 WPS）；全套 **419 项 / 20 文件**、verify 23、
  spec 复现 12、参数契约被校验的对 223 → **231**，A/B/C/D 仍全 0

## D1 上限二次上抬（70 / 40,000）

- P2-3 结束时广告面 56 / 30,885，只剩 1,115 字节；把 A/B/C 三条路摆出来后用户选 **A**（抬上限）
- 门禁与文档同步：`verify.mjs` / `test/deprecated.test.mjs` → **70 工具 / 40,000 字节**（FIXES 42）
- minimal 档保持 4 / 1,348 不变；P5-2 仍负责最终复测，但不再需要在本轮做取舍

## P2-4 高级项（已落地）

- 9 个 action + 9 个工具：透视表列表/刷新/清除、工作簿全部刷新、单变量求解、迷你图（加/清）、
  图表删除、图表标题与轴标题；广告面 **56 → 59 工具 / 32,924 字节**，注册 **243 → 252**
- **修好一个既有且被广告的工具**：`wps_excel_create_pivot_table` 从来建不出透视表。根因是桥里
  `Get-RangeFromAddress` 的 `return $range` 被 PowerShell **展开成 21 个单格**（Range 可枚举），
  改成 `return ,$range` 后正常；同一次还给它补了按步骤名的错误上报
- 为什么一直没人发现：该 helper 只被透视表工具调用，而它零测试覆盖——和 P2 那三个「从未生效的
  action」是同一类问题
- **按实测推迟两项**：切片器（`SlicerCaches.Add2` 建得出缓存，但 `Slicers.Count` 始终 0，用户可见效果
  不确定）与场景管理器（`Worksheet.Scenarios` 在 COM 里是方法，语义读不干净）；探针脚本已留档
- 验收 `test/excel-advanced.test.mjs` **24 项**（真实 WPS）；全套 **443 项 / 21 文件**、verify 23、
  spec 复现 12、参数契约被校验的对 231 → **240**，A/B/C/D 仍全 0

下一步：**P2-5 Excel 重场景 e2e**（把 ListObject + 条件格式 + 打印设置串成一个真实 Excel 任务进一键 e2e）。



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
测试 profile 已在验证后删除；要重建只需 `node scripts/e2e.mjs --profile <name> --setup`（它会自动建 profile 并安装本仓库）。

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

## P2-5 重场景 e2e（已落地）——P2 Excel 做深完成

- 一键 e2e 从 1 个场景扩到 **2 个真实场景**：新增 `orders.xlsx` 任务（ListObject 建表 + 条件格式 +
  横向 A4/页脚/打印标题），独立裸 COM 验证新增 9 项，检查数 **19 → 28**
- **先验证「验证器」**：抽出 `verifyScript()` 对手工做好的产物跑一遍，确认预期值能被正确读出，再赌模型
- 首次完整运行通过：`E2E OK (28 checks) in 94s`，96 次 `wps_*` 调用、无一条自写 COM

## P2 小结（Excel 做深，5 个波次全部落地）

| 波次 | 结果 |
|---|---|
| P2-1 | 18 个「已有能力挂出口」的工具（FIXES 38/39）；顺手修好 4 个从未生效的 action |
| P2-2 | 表（ListObject）全族 8 个 action + 8 个工具（新 COM 代码，FIXES 40） |
| P2-3 | 页面设置/打印/外观/公式审计 8 个 action + 8 个工具（FIXES 41） |
| P2-4 | 高级项 9 个工具；**修好既有工具 create_pivot_table**（Range 被 PowerShell 展开，FIXES 43） |
| P2-5 | 两场景一键 e2e（FIXES 44） |

合计：桥 action 231 → 256、注册工具 217 → 252、广告面 48 → 59 / 32,924 字节（上限 70 / 40,000）、
测试 340 → 443 项 / 21 文件。P2 期间一共修掉 **7 个从未生效的缺陷**（FIXES 38/39/40/43）。

## P3-1 / P3-2 Word 深水区（已落地）

- P3-1：挂出 4 个已有能力——`get_bookmarks`、`get_comments`、`get_document_stats`、`insert_hyperlink`
- P3-2：**表格族新 COM 代码** 9 个 action + 9 个工具（列结构/读数据/写单元格/加行加列/删行删列/
  合并/拆分/样式与边框底纹/转回文本）
- **两处实测缺口，写清楚不硬做**：水印（WPS 页眉 Shapes 不接受任何图形：AddTextEffect/AddShape/
  AddTextbox 都返回对象但 Count 恒为 0）与文档属性（BuiltIn/CustomDocumentProperties 在 WPS 里是坏壳，
  Item/GetType 直接抛「Object reference not set」）；探针脚本留在 test/.artifacts/e2e/word-probe*.ps1
- 桥 action 256 → **265**、注册 252 → **265**、广告面 59 → **62 工具 / 33,836 字节**；
  未工具化 action 台账 **11 → 7**
- 验收 `test/word-deep.test.mjs` **27 项**（真实 WPS Writer）；全套 **470 项 / 22 文件**、verify 23、
  spec 复现 12、参数契约被校验的对 240 → **253**，A/B/C/D 仍全 0

## P3-3 文档生产族（已落地）

- 6 个 action + 6 个工具：页码（页眉/页脚、三种对齐、首页开关）、分栏（栏数/间距/分隔线）、
  修订列表、接受修订、拒绝修订、批注删除（按序号或全部）
- **一处「成功但说错」被自己的验收抓到**：单栏时 WPS 的 `TextColumns.Spacing` 是哨兵值 9999999，
  原样印成「栏间距 9999999 磅」；改成只在多栏时报告间距（断言不会失败，是读输出才发现的）
- 桥 action 265 → **271**、注册 265 → **271**、广告面 62 → **64 工具 / 34,754 字节**
- 验收 `test/word-produce.test.mjs` **19 项**（真实 WPS Writer）；全套 **489 项 / 23 文件**、verify 23、
  spec 复现 12、参数契约被校验的对 253 → **259**，A/B/C/D 仍全 0

## P3-4 长尾（已落地）——P3 Word 做深完成

- 8 个 action + 8 个工具：内容控件（增/读）、脚注、尾注、注释读取、索引、交叉引用、**CSV 邮件合并**
- **一个真实的根因被验收逼出来**：加过脚注后 Word 把光标留在注释正文里，此后「往光标处插内容」
  的动作全都插进了注释——邮件合并看起来成功却没有数据行、尾注直接报参数越界。修法是共享解析器
  `Get-MainTextRange`（不在正文就落到正文末尾并告警）；P3 动过的 5 个动作全部改走它。
  教训：**动作之间会互相改变状态**，只孤立验证单个动作是不够的。
- 桥 action 271 → **279**、注册 271 → **279**、广告面 64 → **66 工具 / 35,568 字节**
- 验收 `test/word-longtail.test.mjs` **20 项**（真实 WPS Writer）；全套 **509 项 / 24 文件**、verify 23、
  spec 复现 12、参数契约被校验的对 259 → **267**，A/B/C/D 仍全 0

## P3 小结（Word 做深，4 个波次全部落地）

| 波次 | 结果 |
|---|---|
| P3-1 | 挂出 4 个已有能力（书签/批注/统计/超链接） |
| P3-2 | 表格族 9 个 action + 9 个工具（新 COM 代码） |
| P3-3 | 文档生产族 6 个工具（页码/分栏/修订接受拒绝/批注删除） |
| P3-4 | 长尾 8 个工具（内容控件/脚注尾注/索引/交叉引用/邮件合并） |

合计 33 个新工具、66 项新验收（word-deep 27 + word-produce 19 + word-longtail 20，含重复计）。
**D2 清单里只有两项没做成**，都是实测的 WPS 限制：水印（页眉 Shapes 不接受任何图形）与文档属性
（BuiltIn/CustomDocumentProperties 是坏壳）。两项都有探针脚本留档。

## P4 PPT 收敛（第一波已落地）

- **做减法**（与前三阶段的加法不同）：四个碎片 setter（shadow/gradient/border/transparency，
  各带一套同义不同名的键）合并为 `wps_ppt_set_shape_effect`，改用带前缀的明确键，只改给出来的项
- 按 D3 删除 3D 族（`set_3d_rotation` / `set_3d_depth` / `set_3d_material`）；工具与 action 一起删，
  否则它们会变成「没有工具驱动」把台账顶上去
- 加了「空操作」闸门：一个效果都没给时直接报错，而不是回一个「成功」却没做事
- 踩坑记录：删除脚本用字符长度算区间，CRLF 下切错位置生成 69 个 parse error；改成按行正则后一次成功 ——
  **用字符偏移做文本手术前必须先确认换行符**
- PPT 工具 **88 → 82**、注册 279 → **273**、桥 action 279 → **273**；广告面不变
- 新增 `test/ppt-slimming.test.mjs` 16 项（含七个被删工具名逐个确认已无法解析）；原有的
  `ppt-contract-fixes` 里 4 项按新契约改写。全套 **524 项 / 25 文件**、verify 23、spec 复现 12、
  参数契约被校验的对 267 → **261**，A/B/C/D 仍全 0
- 未做：把 PPT 压到 ~55 需要继续合并「文本/表格/动画」几族 setter；当前 82 是**功能完整**的减法结果，
  再往下就要动仍在用的能力了（例如动画三种 add 与表格三套 style）——留给下一波按需要决定

## P4 第二波（已落地）——P4 PPT 收敛完成

- 三组合并：动画（add_animation + preset + emphasis → `add_animation`，用 preset/effectKind 分派）、
  表格样式（table/cell/row_style → `set_table_format`，用 row/col 定作用域）、
  页脚（slide_number + footer + date_time → `set_slide_footer`）
- **两个被自己验收抓到的缺陷**：`setAnimation` 的 preset/emphasis 分支漏了 `exit`（会继续走到入场
  分支再报一次错）；页脚工具没报告「页脚本身」的可见性（旧断言正好在测它）。教训：合并分支时每个
  提前返回的分支都要 exit
- `ALIAS_DEBT` 台账 62 → **59**（合并消掉 3 处改名）
- PPT 工具 82 → **76**（P4 合计 **88 → 76**）、注册 273 → **267**、桥 action 273 → **267**；广告面不变
- `test/ppt-slimming.test.mjs` 扩到 **32 项**（含 15 个被删工具名逐个确认无法解析）；
  全套 **541 项 / 25 文件**、verify 23、spec 复现 12、参数契约 255 对 A/B/C/D 全 0

## P5 收尾（P5-1 ~ P5-4 全部落地）

- **P5-1 `execute_method`：保留、隐藏、契约写死**。证据：未工具化 action 台账只剩 7 个刻意的重复实现；
  两次真实 e2e（96 次调用）与 541 项测试里一次都没用过它。但水印与文档属性是实测的 WPS 缺口，
  没有它就没有出路，所以保留为隐藏逃生舱，描述改成明确的最后手段
- **P5-2 广告面重定**：用掉 4 个空位里的 3 个——`wps_word_get_comments`、`wps_word_set_table_cell`、
  `wps_ppt_set_table_format`；最终 **69 工具 / 37,573 字节**（上限 70 / 40,000，余量 1 个工具）。
  上限本身**不再调整**：D1 的 70 / 40,000 当初就是为 P2–P4 这个范围定的，最终几乎正好压在线上
- **P5-3 文档同步**：README 三档数字与工具面构成、FIXES 50、本文件、4 份技能 `reference.md` 与
  `wps_help` 紧凑索引全部按最终注册表重生成
- **P5-4 发布（已完成）**：v0.2.0（首个发布，FIXES 50）与 v0.2.1（仅文档与打包修正，FIXES 51）
  均已打 tag 并发布 Release；npm 未发布（按原约定「真正要发时再做」）

## 加固第 1 波：S1 模态弹窗围堵 + S2 宿主单实例（已落地）

功能面收敛之后，目标从「再加能力」转成「让客户敢用、出问题能解释、坏不了数据」。
第 1 波（推广前必修）已完成，完整实测与实现说明见 `docs/FIXES.md` 52。

- **S1 的假设被实测推翻**：计划里的 `PasswordDocument:=""` **挡不住**密码框（空串=没给密码，照样弹、照样卡死）；
  扩展名不符的老 `.doc` 实测**不弹**转换框。真正有效的是**非空哨兵密码**——
  加密 .docx 从「永久卡死」变成 4 秒内报错，加密 .xlsx 1 秒内报错（`0xFFF40006`）。
  `DisplayAlerts` 不是解（Word 侧本来就是 0），且 WPS 的 IDispatch 不支持命名参数，只能按位置传满 15 个
- **偏离计划一处（有实测依据）**：计划要求超时后杀**整个 WPS 进程树**；实测本机有 **243 个**进程叫 `wps`，
  且 Word/PPT 的 COM 对象不提供 `Hwnd`，「被卡住的是哪一个」无法可靠指认。照做会连带杀掉用户整套 WPS，
  违反判据「不丢数据」，因此**不自动关 WPS**，改为讲清「状态未知」+ 短超时快速失败 + 宿主陈旧接管自愈
- **S2**：命名互斥体 + 租约文件（hostPid / clientPid / 心跳 / 当前动作）；
  第二个宿主回中文错误并非零退出，**不硬上**；只有能证明原主已死（进程 / 它的 MCP 客户端 / 心跳）才接管，
  接管只杀那个陈旧宿主，从不杀 WPS
- **写新测试时抓到两个原有的恢复路径缺陷**：①陈旧子进程的 `exit` 事件会**反杀刚拉起来的新宿主**；
  ②`ready` 帧误清 `suspect` 标志，使短超时形同虚设。两者都已修
- **附带发现（WPS 静默失效）**：`Document.Password` 按长度失效——15 字符把调用卡死，
  **17 字符静默写出完全不加密的文件**（普通 ZIP）。只影响 `wps_execute_method` 逃生舱，已记入 FIXES 52
- **验收**：`test/open-safety.test.mjs` 21 项、`test/host-lease.test.mjs` 18 项、
  `test/watchdog.test.mjs` 15 项全绿；后两个不需要 WPS，已并入 CI 静态门禁
- **新测试的第一个版本自己踩了环境坑，值得记住**：裸 COM 的 fixture 直接 `GetActiveObject` + 直接
  `Close()`，于是既可能拿到**空壳实例**（`Add()` 出来的工作簿没有工作表），又可能弹出**模态保存框**
  把**共享的** WPS 实例钉住，连累后面完全不相干的 `sheet-ops` / `word-lifecycle` 成片失败。
  改成与桥一致的取用方式、`DisplayAlerts` 包住每次 open/close，并在测试最前面加**环境体检**；
  环境不健康就明确报环境问题，不伪装成插件缺陷（详见 FIXES 52 与 HANDOFF §10）

## 加固第 2 波：S3 破坏性前置守卫（三批全部落地，25/25）

- **先把破坏面测清楚**：实测枚举 **29 个破坏性动作 / 35 个调用点**（`Delete()` 27、`Clear()` 3、
  `ClearFormats()` 2、`ClearContents()` 2、`Unlist()` 1、`ResetAllPageBreaks()` 1），
  产出 `docs/destructive-operations.md` 清单（动作 / 影响 / 是否回传统计 / 对应测试）
- **范围类 5 个动作**（clearRange / clearFormats / deleteRows / deleteColumns / deleteSheet）回传
  `data.impact`：`{ address, cells, nonEmpty, preview }`；桥头两个只读助手三级回退取 `nonEmpty`，
  `preview` 只在 ≤ 4096 格时读 `Value2`
- **对象类 8 个动作**（deleteListRow / unlistListObject / resetPageBreaks / clearPivotTable /
  clearSparkline / deleteChart / removeConditionalFormat / deleteNamedRange）回传 `{ name, kind, count, rows, detail, address }`；
  `deleteListRow` 会先把被删那一行的内容读下来再删
- **批注/验证/Word/PPT 12 个动作**（add/removeDataValidation、add/deleteCellComment、deleteTableLine、
  deleteComment、deleteSlide、deleteShape、deleteTextBox、deletePptImage、replacePptImage、removeAnimation）补齐；
  `impact.ts` 上移到 `mcp/src/tools/` 供三个应用共用
- 两个实测坑：WPS 的 `Comment.Text` 是方法（要 `.Text()`）；PPT `AddPicture` 要绝对路径
- **实测**：`test/destructive-guard.test.mjs` **45 项全绿**（真实 Excel/Word/PPT）；整列 1048576 格 / 整行 16384 格
  只回 `cells`+`nonEmpty`（性能取舍）；`spec` / `skills` 重新生成后无漂移
- **进度 25 / 25 个用户数据动作**；S3 只剩「确认框是否弹出」的逐项实测
- 全套 **816 项 / 37 文件**；verify 23、spec 复现 13、参数契约 256 对（A/B/C/D 全 0）

## 加固第 2 波：S4 逐工具覆盖率账本（已完成）

- **覆盖率 161/267 → 267/267**（Excel 118/118、PPT 76/76、Word 59/59、common 7/7）
- `scripts/smoke-tools.mjs`：静态矩阵（按应用分组 + 未覆盖清单）+ `--check` ratchet（267）+ `--live`
  只读冒烟（get/read/list/find/search/query 用占位参数各调一次，断言不超时）
- `test/spec-reproduction.test.mjs` 增加 ratchet 断言：覆盖率只许涨（12 → 13 项）
- **偏离计划**：计划要「每个工具都调一次」，但那会用占位参数去调破坏性动作；改为**只读工具** live 冒烟，
  破坏性工具的运行时验证交给各自的场景测试
- **补全**：三个场景测试（PPT 51 / Excel 31 / Word+common 24）把覆盖率抬到 267/267；记录见 `docs/FIXES.md` 56 / 58
- 全套 **816 项 / 37 文件**

## 附：WPS 进程泄漏修复（FIXES 57）

- **现象**：一轮测试后残留数百个 `wps.exe` / 几十个 `et.exe`，工作集约 35 GB，机器变卡
- **根因**：`Get-WpsApp` 在拿到实例时**无条件**执行 `New-Object -ComObject`——它会真的启动一个
  WPS 进程；代码随后取第一个可用候选（`GetActiveObject` 那个），把刚启动的实例丢弃，WPS 又不自己退出
- **实测证据**：`GetActiveObject` 之后进程数不变；再 `New-Object -ComObject` 立刻 et 21→22、wps 299→300
- **修复**：先 `GetActiveObject`、拿不到才 `New-Object`；记录自启实例；优雅关闭只退出自启且无未保存内容的实例
- **验证**：修复前跑一次 `destructive-guard` 新增 4 个 wps.exe，修复后新增 **0**，测试仍 45/45 全绿
- 这是**生产缺陷**（宿主每次重启都会泄漏一个），测试放大了它

## 加固第 3 波：S5 空 catch 账本（已完成）

- 实测桥 `mcp/scripts/wps-com.ps1` **25 处空 catch** + 手写宿主 `host/wps-com-host.ps1` **4 处**（比计划里的 13 处多，S1/S3 新增了探测/还原类兜底）
- 新增 `test/silent-catch.test.mjs`：以「规范化源码行」为键的 `ALLOWLIST`（带原因与次数），
  断言「无未登记、无过期、都有原因、总数一致」；已并入 CI 静态门禁
- **不做一刀切补日志**：登记后逐项给出静默是合理的理由（COM 属性探测、DisplayAlerts 还原、统计回退等）
- 全套 **816 项 / 37 文件**；记录见 `docs/FIXES.md` 59

## 加固第 3 波：S6 失败语义与超时契约（已完成）

- 新增 `docs/error-contract.md`：结果信封、批量契约、超时契约（三档 + 长动作清单）、调用方该做什么、断言位置
- 新增 `test/error-contract.test.mjs`（14 项）：批量空/超 50 整体拒绝；部分失败**不抬高整批**、失败项单独记账、**后续继续执行**；
  结果截断 2000 字符；`wps_call` 拒绝未知工具与门面递归；并静态断言契约文案仍在真源里
- 超时那半由 `test/watchdog.test.mjs`（15 项，stub 宿主）覆盖——两者合起来文档里每条语义都有断言
- 全套 **816 项 / 37 文件**；记录见 `docs/FIXES.md` 60

## 加固第 3 波：S7 面向客户的中文错误文案（已完成）

- 桥头新增 `Format-WpsErrorText`，所有失败统一成三段式：`<中文一句话 或 原文>（动作：<action>）下一步：<建议>。（原始信息：<英文/HRESULT>）`
- 命中的重复英文短语换成中文、原文降级为「原始信息」；**没命中的原样保留**，只补动作名与下一步（历史断言与重试都不受影响）
- 宿主生成器在 `Output-Json` 里调用它并记录当前动作；宿主捕获的未处理异常也一并格式化（补上了以前绕过 Output-Json 的缺口）
- 新增 `test/error-wording.test.mjs`（28 项）：抽样 26 条必失败路径，逐条断言三段式 + 点名动作 + 幂等
- 契约见 `docs/error-contract.md` §1b；全套 **816 项 / 37 文件**；FIXES 61

## 加固第 4 波：S8 + S9 版本检查与安装自检（已完成）

- **S8**：`doctor.mjs` 只读检查已安装 WPS 的版本与架构（PE Machine 判 x64/x86），低于 12.1 或 32 位报 ERROR、找不到报 WARN；
  `getAppInfo` 增加 `version`/`build`，运行中的版本在 `wps_status` 可见；`wps_execute_method` 维持隐藏
- **S9**：`doctor.mjs` 增加插件接线自检（`cordis.patch.yml` 必须同时声明 `wps-office-next-plugin` 与 `mcp-wps-office-next`）；README 排错表补「先重试一次」
- 新增 `test/install-selfcheck.test.mjs`（6 项，不需要 WPS，已进 CI）
- 全套 **816 项 / 37 文件**；记录见 `docs/FIXES.md` 62

## 附录：S3 余量——破坏性动作确认框实测（已完成）

- `test/confirm-dialog.test.mjs`：后台每 200 ms 用 `EnumWindows` 监视 `Qt*` 可见窗口（WPS 的模态框），
  同时跑 destructive-guard / excel-advanced / excel-page-setup（覆盖全部 25 个破坏性动作）
- **结果：全程 0 个 Qt 窗口**——25 个动作都不弹确认框，无需补 `DisplayAlerts` 抑制；`deleteSheet` 保留的抑制只作兜底
- 全套 **816 项 / 37 文件**；记录见 `docs/FIXES.md` 63 与 `docs/destructive-operations.md`

## 审计 P2：C7 目标歧义警告 + 入参形状守卫（已完成）

- 桥在「多文件打开 + 调用方没指定目标名」时补一条 warning（`Get-WorksheetByParam` / `Resolve-Worksheet` /
  `Get-TargetPres` 三处共用解析点），显式给目标或单文件时不打扰；顺手把 15 处内联的 sheet 解析回迁到共用解析点
  （否则 `read_range`/`write_range` 主路径绕过警告）
- `tool-registry.validateArguments` 补**结构**校验：`array`/`object` 与标量错位直接拒；标量之间放行（schema 比真实契约窄）
- 新增 `test/target-ambiguity.test.mjs`（13 项，真实 WPS）与 `test/arg-shape-guard.test.mjs`（6 项，已进 CI 静态门禁）
- 空 catch 账本 29 → 38（桥 34 + 宿主 4），全部登记；全套 **835 项 / 39 文件**；记录见 `docs/FIXES.md` 64

## 进程残留根因（FIXES 65，已完成）

- 现象：整轮测试每轮残留约 10 个 `wps.exe`（累计 5.8GB），按轮成批出现
- 根因：Windows 上 libuv 把非 detached 子进程放进 job object，`child.kill()` MCP server 时**整棵树一起被终止**，
  常驻宿主随之消失，它启动的 WPS 实例成了孤儿（实测：插桩标记一个都没写，状态文件停在 `phase=idle`）
- 落地：宿主循环后补退出兜底（覆盖 EOF 等所有退出路径）；新增 `scripts/run-tests.ps1`，每个测试文件跑完
  关掉**窗口标题为空**的无头 WPS 进程（有标题的真实窗口不碰）
- 排除：`detached: true` 理论上更彻底，但 PowerShell 5.1 在 `DETACHED_PROCESS` 下静默退出，未采用
- 验收：从 0 基线整轮跑完 `ORPHANS_LEFT=0`；全套 **835 项 / 39 文件**；记录见 `docs/FIXES.md` 65

## 跨会话回收 WPS 孤儿（FIXES 66，已完成）

- 归属写盘：`Get-WpsApp` 新起实例时把 `{hostPid, clientPid, apps}` 写进 `~/.wps-office-mcp/owned-apps.json`
- 新宿主在租约下调用 `Invoke-WpsOrphanReclaim`：宿主与客户端都死了才回收，且有未保存内容不动
- 测试 `test/orphan-reclaim.test.mjs` 7 项：强杀留孤儿 → 下一个会话收掉；删掉记录则不动
- 全套 **842 项 / 40 文件**；记录见 `docs/FIXES.md` 66

## P3 四项（已完成）

- **运行时版本前置检查**：`wps_status` 回报真实 WPS 版本与兼容性。第一版用 `Application.Version` 会误报
  （WPS 12.1 报 12.0），改为读 exe 文件版本，判定抽成 `mcp/src/utils/wps-version.ts`，14 项单测进 CI
- **lint 门禁**：`scripts/lint.mjs`（手写 PowerShell 的 BOM/CRLF、制表符与行尾空白、`console.*`、测试退出码），
  首跑 164 文件 6 处违规（文件末尾缺换行）已补齐，现 165 文件 0 违规，已进 CI
- **e2e 归属记录断言**：新增「记录不指向死主人」一条，e2e 28 → 29 项；原打算加「进程不残留」，实测 `Quit()` 返回成功但进程不退，删掉并在注释里写明理由
- **全新 profile 安装验收**：`scripts/accept-install.mjs` 装一遍再拆掉，14 项全绿（含从安装副本跑 doctor）
- 全套 **856 项 / 41 文件**；记录见 `docs/FIXES.md` 67

## 全项目最终数字（P5 收尾 + 加固第 1 波）

| 项 | 起点（上游基线） | 现在 |
|---|---|---|
| 广告面 | 250 工具 / 141,872 字节 | **69 工具 / 37,573 字节**（降 73.5%） |
| 注册工具 | 250 | **267**（Excel 118 / Word 59 / PPT 76 / 通用 14） |
| 桥 action | — | **267** |
| 测试 | 0 | **856 项 / 41 文件**（其中 8 个文件不需要 WPS，已进 CI） |
| 门禁 | 无 | verify 23、spec 复现 13、参数契约 256 对（A/B/C/D 全 0）、一键 e2e 28 项 |
| 台账 | — | 别名债务 **59**、未工具化 action **7**（全部刻意保留） |

P0–P4 五个阶段共修掉 **7 个「从未生效」的缺陷**（FIXES 38/39/40/43），另有两处实测缺口如实记录
（水印、文档属性——WPS 自身不支持）。

**发布状态**：v0.2.0、v0.2.1、v0.3.0（2026-09-16，加固第 1 波）、**v0.4.0（2026-09-19，加固第 2 波 + 审计 P2 + P3）**均已发布；本机完整回归
**856/0**、一键 e2e **29/29** 全绿，可以开始小范围推广；第 2～4 波（S3 数据安全 / S4 冒烟矩阵 /
S5 空 catch / S6 失败契约 / S7 文案 / S8 版本检查 / S9 安装兜底）与审计 P2、FIXES 65/66 已全部落地，尚未打包成版本。



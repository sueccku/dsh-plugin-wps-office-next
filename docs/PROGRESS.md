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

剩余待办：

- 把 e2e 做成可脚本化的一键回归（目前是手工跑 + 读会话日志复盘）；
- 其余 15 对重复工具（参数接口不同）；
- 12 个 builtin 与 pro 工具重复（未广告，但仍出现在 wps_help 目录里）。

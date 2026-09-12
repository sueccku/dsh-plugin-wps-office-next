# 进度

MVP 目标（P0-P3 + 最小技能集）已全部完成并在真实 DSH 上端到端验证。

| 阶段 | 状态 | 证据 |
|---|---|---|
| P0 冻结基线 | 完成 | baseline/upstream-0.1.0/：250 工具、141,872 schema 字节、约 40.5k tokens/请求、248 个 COM action；baseline/known-defects.md（24 条） |
| P1 合仓 + Windows 化 | 完成 | mcp/ 迁入 43 个 TS 源文件；删除 mac 传输、keepalive、平台分支、axios |
| P2 常驻 COM host | 完成 | warm ping 1ms（原 969ms）；test/com-host.test.mjs 6/6 |
| P3 工具面收敛 | 完成 | standard 43 工具 / 23,198 字节 / 约 6.6k tokens；scripts/verify.mjs 22/22 |
| 最小技能集 + DSH 自举 | 完成 | plugin.js + cordis.patch.yml；4 个技能；test/plugin.test.mjs 32/32；真实 profile 启动验证通过 |
| P4 缺陷修复 | 未开始 | 24 条缺陷清单待处理 |
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
- 分发：提交身份已设为 sueccku <1075322047@qq.com> 并重写历史；推送前需先完成 gh auth login（gh 2.100.0 已装）。
# 进度

目标：MVP = P0-P3 + 最小技能集。缺陷修复（P4）与全量技能（P5）留到 MVP 之后。

| 阶段 | 状态 | 证据 |
|---|---|---|
| P0 冻结基线 | 完成 | baseline/upstream-0.1.0/：250 工具、141,872 schema 字节、约 40.5k tokens/请求、248 个 COM action；baseline/known-defects.md（24 条） |
| P1 合仓 + Windows 化 | 完成 | mcp/ 迁入 43 个 TS 源文件；删除 mac 传输、keepalive、平台分支、axios；tsc 构建通过 |
| P2 常驻 COM host | 完成 | host/ 常驻宿主 + com-host.ts 客户端；warm ping 中位 1ms；test/com-host.test.mjs 6/6 |
| P3 工具面收敛 + 门面派发 | 完成 | standard 31 工具 / 17,797 字节 / 约 5.1k tokens；scripts/verify.mjs 22/22 |
| MVP 最小技能集 + DSH 自举 | 未开始 | cordis.patch.yml、plugin.js、skills/ |

## P3 实测（tools/list，走 MCP 协议）

| 档位 | 工具数 | schema 字节 | 约 tokens | 说明 |
|---|---|---|---|---|
| 基线（收敛前） | 250 | 141,872 | 40,535 | 全部直接广告 |
| full | 254 | 143,217 | 40,919 | 全量 + 4 门面，保留完整描述 |
| standard（默认） | 31 | 17,797 | 5,085 | 4 门面 + 27 精选，降 87.5% |
| minimal | 4 | 1,348 | 385 | 仅门面，降 99% |

档位由 WPS_OFFICE_TOOLSET 环境变量切换（minimal / standard / full），默认 standard。

## 门面设计

| 工具 | 作用 |
|---|---|
| wps_status | 连接状态 + 活动应用 + 当前档位与隐藏工具数；编辑前先调用 |
| wps_help | 无参看分组概览；传 app 查目录；传 query 搜索；传 tool 取完整 inputSchema |
| wps_call | 执行任意已注册但未广告的工具；拒绝未知工具与门面递归 |
| wps_batch | 顺序批量执行，单次上限 50 项 |

未广告的工具仍可按原名直接调用，因此既有提示词与习惯不会失效。

## 预算取舍说明

最初定的 12,000 字节上限无法在不损失信息的前提下达到：实测把全部参数描述删掉，45 个工具也只能降到
15,148 字节——成本主体是 schema 结构本身（属性名/类型/枚举/required），约 527 字节/工具，而参数描述
承载了单位与枚举（例如 PPT 的 像素 vs 磅），正是上游缺陷的来源之一。因此改为收缩工具数量：
27 个精选工具 + 4 门面 = 31，并把门禁设为 40 工具 / 22,000 字节。

## P2 实测（走 MCP 全链路）

| 调用 | 改造前 | 改造后（warm） |
|---|---|---|
| wps_common_ping | 969ms | 1ms（中位，n=7） |
| wps_common_wire_check | 828ms | 1ms（中位，n=5） |
| wps_common_get_app_info | 1858ms | 2ms（中位，n=5） |

## 待办与已知风险

- 尚未接入 DSH：cordis.patch.yml、plugin.js 与 skills/ 是 MVP 的下一步。
- 首次推送前必须把占位提交身份 wps-dsh-dev 换成真实身份。
- 已知缺陷（含 wps_common_get_app_info 只回前缀的 S7）留到 P4。

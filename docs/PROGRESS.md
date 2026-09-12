# 进度

目标：MVP = P0-P3 + 最小技能集。缺陷修复（P4）与全量技能（P5）留到 MVP 之后。

| 阶段 | 状态 | 证据 |
|---|---|---|
| P0 冻结基线 | 完成 | baseline/upstream-0.1.0/：250 工具、141,872 schema 字节、约 40.5k tokens/请求、245 个 COM action；baseline/known-defects.md（24 条） |
| P1 合仓 + Windows 化 | 完成 | mcp/ 迁入 43 个 TS 源文件；删除 mac 传输、keepalive、平台分支、axios；tsc 构建通过 |
| P2 常驻 COM host | 完成 | host/ 常驻宿主 + com-host.ts 客户端；warm ping 中位 1ms、getAppInfo 2ms、getSheetList 2ms；test/com-host.test.mjs 6/6 通过 |
| P3 action 注册表 + 工具面收敛 | 未开始 | 目标：standard 档 <= 45 工具、<= 12,000 schema 字节 |
| MVP 最小技能集 | 未开始 | DSH 技能文档 + 自举路径 |

## P2 实测（走 MCP 全链路）

| 调用 | 改造前 | 改造后（warm） |
|---|---|---|
| wps_common_ping | 969ms | 1ms（中位，n=7） |
| wps_common_wire_check | 828ms | 1ms（中位，n=5） |
| wps_common_get_app_info | 1858ms | 2ms（中位，n=5） |
| wps_excel_get_sheet_list | 约 900ms+ | 2ms（中位，n=5） |

MCP server 启动 193ms；宿主首次调用（含进程启动、248 分支 JIT 预热）约 1.0s，之后进入 1-2ms 稳态。

## 已验证的事实

- 宿主被外部 kill 后，下一次调用自动重启并成功（test/com-host.test.mjs）。
- 5 个并发调用全部成功，无串扰。
- 未知 action 返回结构化失败，不会伪装成功。
- 生成脚本可复现：248 个 switch case、16 个函数、exit 残留 0、BOM 为 EF BB BF。
- 裁剪后新构建仍完整暴露 250 个工具，说明 Windows 化未损坏工具层。

## 待办与已知风险

- P3 的 action 注册表必须来自 wps-com.ps1 的 248 个 case，并与 238 个工具定义对齐，避免再次出现文档计数漂移。
- 首次推送前必须把占位提交身份 wps-dsh-dev 换成真实身份。
- mcp/scripts/wps-com.ps1 保留为生成脚本的输入（上游来源），与 host/wps-actions.ps1 有意重复。

# 进度

目标：MVP = P0-P3 + 最小技能集。缺陷修复（P4）与全量技能（P5）留到 MVP 之后。

| 阶段 | 状态 | 证据 |
|---|---|---|
| P0 冻结基线 | 完成 | baseline/upstream-0.1.0/（250 工具、141,872 schema 字节、约 40.5k tokens/请求、245 个 COM action）；baseline/known-defects.md（24 条） |
| P1 合仓 + Windows 化 | 完成 | mcp/ 迁入 43 个 TS 源文件；删除 mac-poll-server.ts、wps-keepalive.ts、平台分支、axios 依赖；npx tsc 构建通过；verify.mjs 7/7 通过 |
| P2 常驻 COM host | 未开始 | 目标：ping < 100ms，替代每次 spawn powershell |
| P3 action 注册表 + 工具面收敛 | 未开始 | 目标：standard 档 <= 45 工具、<= 12,000 schema 字节 |
| MVP 最小技能集 | 未开始 | DSH 技能文档 + 自举路径 |

## 已验证的事实

- 裁剪后新构建仍完整暴露 250 个工具，说明 Windows 化未损坏工具层。
- mcp/dist/index.js 5,186 字节，dist 共 172 个文件；按决策将随仓库提交。
- mcp/scripts/wps-com.ps1 保留 UTF-8 BOM（EF BB BF），.gitattributes 用 * -text 防止改写。

## 待办与已知风险

- P2 需要决定 host 常驻进程的进程数与超时分级；注意 powershell.exe 5.1 的 STA 约束。
- P3 的 action 注册表必须来自 wps-com.ps1 的 245 个 case，并与 238 个工具定义对齐，避免再次出现文档计数漂移。
- 首次推送前必须把占位提交身份 wps-dsh-dev 换成真实身份。

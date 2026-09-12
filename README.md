# dsh-plugin-wps-office-next

专精 Windows/COM 的一站式 DeepSeek Harness 插件：让 DSH 通过 MCP 直接操控 WPS 表格 / 文字 / 演示。

本项目是 [lc2panda/wps-skills](https://github.com/lc2panda/wps-skills) 与
[CatNebulaaaa/wps-dsh-plugin](https://github.com/CatNebulaaaa/wps-dsh-plugin) 的独立合并演进版，
非官方项目，与金山办公、DeepSeek 及上游作者无隶属关系。

## 状态

开发中，当前处于 P0（基线冻结）完成阶段。**尚不可安装使用。**

## 已锁定的项目决策

| 决策 | 取值 |
|---|---|
| 平台 | 仅 Windows，仅 COM；macOS / Linux 支持整体移除 |
| 目标环境 | WPS 12.1+ x64；不兼容 x86，不兼容多组件模式 |
| 形态 | 单一 npm 包 = DSH bundle + 自带 MCP server + 自带 COM host + 全部 skills |
| MCP serverName | wps-office-next（工具显示为 mcp__wps-office-next__xxx） |
| 默认工具面 | standard 档：31 个工具（4 门面 + 27 精选），17,797 schema 字节约 5.1k tokens；预算上限 40 工具 / 22,000 字节 |
| 传输层 | 常驻 PowerShell STA host，复用上游 245 个 COM action |
| 加载项 | 全部删除；Windows COM 路径不依赖任何 WPS 加载项 |
| 构建产物 | 预构建产物入库，保证 dsh plugin add github:... 开箱可用 |
| 文档语言 | 中文 |
| 仓库 | 本仓根目录即 D:\dsh\a；全新 git 历史；GitHub 发布 |

## 仓库结构

| 路径 | 用途 |
|---|---|
| baseline/ | 上游基线快照与缺陷清单，供回归比对 |
| scripts/ | snapshot-tools.mjs（工具目录快照）、verify.mjs（冒烟验证，P1 落地） |
| mcp/ | 合并后的 MCP server（P1 从上游迁入并裁剪） |
| host/ | 常驻 COM host 与拆分的 action 模块（P2） |
| skills/ | DSH 技能文档（MVP 最小技能集） |
| docs/ | 设计与迁移文档 |

## 开发命令

| 命令 | 作用 |
|---|---|
| node scripts/snapshot-tools.mjs <entry> <out> | 采样某个 MCP server 的 tools/list 并统计 token 成本 |
| git log --oneline | 查看阶段进度 |

## 提交身份

本仓当前使用占位提交身份 wps-dsh-dev <wps-dsh-dev@localhost>，首次推送前需替换为真实身份。
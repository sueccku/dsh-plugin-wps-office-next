# dsh-plugin-wps-office-next

专精 Windows/COM 的一站式 DeepSeek Harness 插件：让 DSH 通过 MCP 直接操控 WPS 表格 / 文字 / 演示。

本项目是 [lc2panda/wps-skills](https://github.com/lc2panda/wps-skills) 与
[CatNebulaaaa/wps-dsh-plugin](https://github.com/CatNebulaaaa/wps-dsh-plugin) 的独立合并演进版。
非官方项目，与金山办公、DeepSeek 及上游作者无隶属关系。

## 状态

MVP 完成并已在真实 DSH 上端到端验证（P0 基线、P1 合仓、P2 常驻 COM 宿主、P3 工具面收敛、最小技能集）。
缺陷修复（P4）与全量技能（P5）留到 MVP 之后。

## 安装

本地路径：

    dsh plugin --profile <name> add <本仓库绝对路径>

从 GitHub（建议固定 commit）：

    dsh plugin --profile <name> add github:CatNebulaaaa/dsh-plugin-wps-office-next#<sha>

随后正常启动该 profile 即可。**不需要任何环境变量**：plugin.js 会在加载时自行解析包内的 MCP 入口与
COM 宿主脚本路径，不再需要上游时代的 WPS_SKILLS_ROOT / WPS_OFFICE_MCP_ENTRY 配置。

前提：Windows x64、WPS 12.1+、Node 22.19+ 或 24+。无需安装任何 WPS 加载项。

## 自检

    node scripts/doctor.mjs                  # 平台、Node、PowerShell STA、包内容、WPS COM、profile 接线
    node scripts/verify.mjs mcp/dist/index.js # 启动服务并检查门面、派发、预算门禁（22 项）
    node test/plugin.test.mjs                # DSH 入口：路径发布与技能注册（32 项）
    node test/com-host.test.mjs              # 常驻宿主：恢复、串行化、错误语义（6 项）

## 工具面

基线（上游 0.1.0）会往每次请求塞 250 个工具、141,872 schema 字节、约 40.5k tokens。本插件把
tools/list 收敛为三档，由 WPS_OFFICE_TOOLSET 切换（默认 standard）：

| 档位 | 工具数 | schema 字节 | 约 tokens | 内容 |
|---|---|---|---|---|
| minimal | 4 | 1,348 | 385 | 仅门面 |
| standard（默认） | 43 | 23,198 | 6,628 | 4 门面 + 39 精选，较基线降 83.6% |
| full | 254 | 143,217 | 40,919 | 全量，保留完整描述 |

未广告的工具仍然完全可用，两条路都能走：按全名直接调用，或先查后调：

| 门面工具 | 作用 |
|---|---|
| wps_status | 连接状态 + 活动应用 + 当前档位；编辑前先调用 |
| wps_help | 无参看分组概览；传 app 查目录；传 query 搜索；传 tool 取完整 inputSchema |
| wps_call | 执行任意已注册但未广告的工具 |
| wps_batch | 顺序批量执行，单次上限 50 项 |

## 性能

旧实现每次工具调用都要新起一个 PowerShell 进程并重新解析 4,939 行脚本。现在是一个常驻宿主
（powershell.exe 5.1，-STA）通过 stdin/stdout 收发 JSON 行，工具层只做转发：

| 调用 | 改造前 | 改造后（warm） |
|---|---|---|
| wps_common_ping | 969ms | 1ms |
| wps_common_wire_check | 828ms | 1ms |
| wps_common_get_app_info | 1858ms | 2ms |

宿主首次调用约 1.0s（进程启动 + 248 分支预热），之后进入 1-2ms 稳态；宿主被 kill 后下次调用自动重启。

## 目录结构

| 路径 | 用途 |
|---|---|
| cordis.patch.yml | DSH 接线：插件入口 + MCP 客户端（serverName 为 wps-office-next） |
| plugin.js | DSH 入口：发布包内路径、注册 4 个技能 |
| mcp/ | MCP server（TypeScript，dist 已入库，安装即用） |
| host/ | 常驻 COM 宿主与生成物 wps-actions.ps1（248 个 action） |
| skills/ | 技能文档：路由技能 + 三个应用技能 + 生成的 reference.md |
| scripts/ | doctor、verify、工具目录快照、技能生成、宿主生成与延迟基准 |
| test/ | 宿主与 DSH 入口的回归测试 |
| baseline/ | 上游基线快照与缺陷清单，供回归比对 |
| docs/ | 进度与设计说明 |

## 已锁定的项目决策

| 决策 | 取值 |
|---|---|
| 平台 | 仅 Windows，仅 COM；macOS / Linux 支持整体移除 |
| 目标环境 | WPS 12.1+ x64；不兼容 x86，不兼容多组件模式 |
| 形态 | 单一 npm 包 = DSH bundle + 自带 MCP server + 自带 COM host + 全部 skills |
| MCP serverName | wps-office-next |
| 默认工具面 | standard 档 43 个工具；预算上限 45 工具 / 25,000 字节 |
| 传输层 | 常驻 PowerShell STA 宿主，复用上游 248 个 COM action |
| 加载项 | 全部删除；Windows COM 路径不依赖任何 WPS 加载项 |
| 构建产物 | 预构建产物入库，保证安装后开箱可用 |
| 文档语言 | 中文 |

## 提交身份

提交身份为 sueccku（见 git log）。

## 许可与来源

本仓库使用 MIT License。上游为 MIT，来源与处理方式见 THIRD_PARTY_NOTICES.md。
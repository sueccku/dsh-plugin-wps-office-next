# dsh-plugin-wps-office-next

[![ci](https://github.com/sueccku/dsh-plugin-wps-office-next/actions/workflows/ci.yml/badge.svg)](https://github.com/sueccku/dsh-plugin-wps-office-next/actions/workflows/ci.yml)
![advertised tools](https://img.shields.io/badge/advertised%20tools-44%2F45-blue)
![platform](https://img.shields.io/badge/platform-Windows%20x64%20%C2%B7%20WPS%2012.1%2B-informational)

专精 Windows/COM 的一站式 DeepSeek Harness 插件：让 DSH 通过 MCP 直接操控 WPS 表格 / 文字 / 演示，
自带 MCP server、常驻 COM 宿主与技能文档，**不需要安装任何 WPS 加载项，也不需要配置任何环境变量**。

本项目是 [lc2panda/wps-skills](https://github.com/lc2panda/wps-skills) 与
[CatNebulaaaa/wps-dsh-plugin](https://github.com/CatNebulaaaa/wps-dsh-plugin) 的独立合并演进版。
非官方项目，与金山办公、DeepSeek 及上游作者无隶属关系。

> **当前状态**：功能已完整可用，并在真实 DSH + 真实 WPS 上做过两轮端到端验证。
> 尚未打 tag、未发布 Release、未发布 npm；版本号 `0.1.0` 只是占位。细节见下方「当前进展」。

## 能做什么

一个包覆盖 WPS 三大应用，共 **209 个 MCP 工具**（每次请求只广告其中 44 个，其余按需查询后调用）：

| 应用 | 代表能力 |
|---|---|
| 表格（Excel） | 读写区域、单元格格式与数字格式、公式与求值、排序筛选、条件格式、透视表、图表、批注、冻结窗格、导出图片、转 PDF |
| 文字（Word） | 新建/打开/保存/关闭文档、读写正文与段落、样式与字体、查找替换（区分「只查找」与「替换」）、页眉页脚、分节符、目录、模板字段、页面设置、批注、修订 |
| 演示（PPT） | 新建/打开演示、幻灯片增删与版式、形状与文本框、填充与字体、表格、图片、切换与动画、演讲者备注、导出幻灯片为图片 |
| 通用 | 连接状态、保存/另存、格式转换、可执行任意桥 action 的通用派发器 |

## 当前进展

| 阶段 | 状态 |
|---|---|
| P0 基线审计（上游 25 处缺陷逐条定位） | 完成，见 [baseline/known-defects.md](baseline/known-defects.md) |
| P1 合仓（两个上游合成单一 DSH bundle） | 完成 |
| P2 常驻 COM 宿主（取代每次新起 PowerShell） | 完成，冷启动约 1.0s，稳态 1–2ms |
| P3 工具面收敛（250 → 44 个广告工具） | 完成；P2 做深后广告面 56，预算按 D1 两次上抬到 70 工具 / 40,000 字节 |
| P4 缺陷修复（工具面/参数契约/静默失败） | 完成，29 条修复记入 [docs/FIXES.md](docs/FIXES.md) |
| P5 技能文档（路由 + 三个应用 + 生成的参考表） | 完成，4 个技能 443 项测试覆盖 |
| 真实端到端验证 | 完成两轮，见下 |
| 发布（tag / Release / npm） | **未开始**（按计划留到 README 与 CI 就绪之后） |

### 端到端验证做了什么

在独立 headless profile 上给了模型一个跨应用任务：打开一份 12 行销售明细 →
按「地区」汇总写入新工作表并降序 → 画簇状柱形图 → 新建 Word 文档写结论并另存 → 保存并关闭两个文件。

| 轮次 | 耗时 | 结果 | 产物复核（另起 COM 会话重开文件） |
|---|---|---|---|
| 第 1 轮 | 72s / exit 0 | 全部完成 | 汇总 4 行正确、图表 ChartType=51 带标题、Word 标题1+正文 145 字符、收尾 0 文档残留 |
| 第 2 轮（修复后） | 65s / exit 0 | 全部完成 | 同上，并额外验证数字格式 `#,##0` 与表头加粗 |

两轮都复盘了每一次工具调用（读 DSH 会话日志）。第 1 轮暴露的 5 个真实缺陷已全部修复并加了回归测试：

1. `wps_help` 的自由文本检索把整条查询当子串匹配，"新建 文档 create new" 一律 0 命中 →
   改为按 token 打分 + 中文二元组覆盖率；
2. `wps_excel_get_open_workbooks` 把对象数组直接 join，每个工作簿都显示成 `[object Object]`；
3. `wps_word_get_active_document` 打印 `页数: undefined`，并把未落盘文档的名字当成路径；
4. `wps_excel_set_formula` 每次都回读「计算结果: null」，多格目标在常驻宿主里直接失败
   （改为按 A1 地址逐格广播，并提示「同一公式写入多个单元格、不做相对引用调整」）；
5. Word 连隐藏的建/关文档工具都没有（Excel/PPT 都有），补上 `wps_word_create_document`
   与 `wps_word_close_document`。

完整复盘（含失败现场与实测数据）见 [docs/FIXES.md](docs/FIXES.md) 第 24～29 条。

### 已经实测过的

- 两轮真实 e2e（上表）、产物用独立 COM 会话重开复核；
- **一键 e2e 验收已跑通**：`node scripts/e2e.mjs --profile <name>` → 28 项检查、两个真实场景、exit 0；
  另做过负向验证（超时压到 5 秒 → 10 项 FAIL、exit 1），确认它会真的失败而不只是打印 PASS；
- 443 项测试 + 23 项门禁全绿，参数契约对账 240 对、四类静默失效均为 0；
- **CI 在 GitHub 的 windows-latest 上跑通**（首次运行 13 步全绿、67 秒）：tsc 与三处产物漂移检查
  在干净的 runner 上复现，不需要本机任何环境；
- 常驻宿主的串行化、崩溃重启、超时与 warning 语义；
- 本机路径安装与 **GitHub 安装**（`79f0c96`）：拉包 → `--dump-config` 出现插件与 MCP 两行 →
  真跑一次 `wps_status`，7 秒返回 `connected/44/237/193`（P0 清理后注册数已降到 209，该次实测记录保留原值）；
- 关闭文档不弹模态框、不泄漏未保存文档（收尾后已打开工作簿/文档均为 0）。

### 还没做的

- **CI 只覆盖静态部分**：tsc + 三处漂移检查 + `verify --static` + 参数契约 + 两个静态测试；
  一键 e2e 与其余 13 个测试文件要驱动真实 WPS，只能在装了 WPS 的本机跑（没有自托管 runner）；
- ~~12 个上游遗留 builtin 工具与 pro 工具重复~~ **已清理**（FIXES 第 32 条）：删 11 个，只留 `wps_execute_method` 逃生舱；
- 其余 15 对「同 action、参数接口不同」的重复工具尚未合并；
- 9 处 handler 实参静态不可读、1 处桥无键表（动态键）在参数契约报告里列名待查。

## 安装

前置：**Windows x64**、**WPS Office 12.1+ x64**（不支持 x86 与多组件模式）、
**Node 22.19+**、**Windows PowerShell 5.1**（常驻宿主会断言 PowerShell 主版本为 5）。

任选一种装法，然后正常启动该 profile：

    # 从本仓库绝对路径（开发用）
    dsh plugin --profile <name> add <本仓库绝对路径>

    # 从 GitHub（建议固定 commit）
    dsh plugin --profile <name> add github:sueccku/dsh-plugin-wps-office-next#<sha>

两条都实测过。GitHub 安装的记录（commit `79f0c96`）：pnpm 拉取 121 个包、7.6s 完成，
`dsh --profile <name> --dump-config` 里出现 `wps-office-next-plugin` 与 `mcp-wps-office-next` 两行，
包内含 `mcp/dist/index.js`、`host/wps-actions.ps1` 与 4 个技能目录。

装完可以直接冒烟：在新的 profile 里让 DSH 调一次 `wps_status`。GitHub 安装的那一份实测
**7 秒**返回 `connected: true、advertisedTools: 44、registeredTools: 237、hiddenTools: 193`（当时的数字；P0 清理后为 209 / 165），
说明包内的 MCP server 与 COM 宿主都能正常起来。

注意：按 GitHub 方式安装时，pnpm 只打包 `package.json` 的 `files` 白名单，
因此 `test/`、`docs/`、`baseline/` **不在安装结果里**——它们是给仓库看的，运行不需要。

## 工具面

基线（上游 0.1.0）会往每次请求塞 250 个工具、141,872 schema 字节（约 40.5k tokens）。
本插件把 `tools/list` 收敛为三档，由 `WPS_OFFICE_TOOLSET` 切换（默认 `standard`）：

| 档位 | 工具数 | schema 字节 | 约 tokens | 内容 |
|---|---|---|---|---|
| minimal | 4 | 1,348 | 385 | 仅 4 个门面工具 |
| **standard（默认）** | **59** | **32,924** | **9,407** | 门面 + 55 个精选工具，字节数较基线降 76.8% |
| full | 252 | 148,011 | 42,289 | 全量，保留完整描述 |

注册目录 **252** 个工具：广告 59、隐藏 193；另有 18 个已合并的旧名字**不占注册位**，只在派发期解析成规范工具
（旧名照样能调，`wps_help {tool:"旧名"}` 会告诉你该用哪个）。按应用分布为 Excel 118 / Word 32 / PPT 88 / 通用 9 / builtin+门面 5。

**未广告的工具完全可用**，两条路都能走：按全名直接调用，或先查后调。

| 门面工具 | 作用 |
|---|---|
| wps_status | 连接状态 + 活动应用 + 当前档位 + 广告/注册/隐藏数量；编辑前先调用 |
| wps_help | 无参看分组概览；传 `app` 列该应用目录；传 `query` 模糊搜索（支持中文与中英混排）；传 `tool` 取完整 inputSchema |
| wps_call | 执行任意已注册但未广告的工具 |
| wps_batch | 顺序批量执行，单次上限 50 项 |

配套两道契约防线：工具层发给桥的多余参数会被拒绝并列出可接受的键
（`unknown parameter(s) for 'xxx': a, b | accepted: b, c`）；`scripts/param-contract.mjs`
零副作用地把 211 对工具/action 的参数契约对账一遍，结果写入
[docs/param-contract.md](docs/param-contract.md)。

## 技能

`plugin.js` 注册 4 个技能，DSH 会按任务自动加载（任务是跨应用时先加载路由技能）：

| 技能 | 触发场景 |
|---|---|
| wps-office-next | 路由与通用约定：先 `wps_status`、参数名以 schema 为准、变更后复验、失败与 `warnings` 如何如实转达 |
| wps-excel | 表格、区域、格式、公式、图表、透视表、数据清洗 |
| wps-word | 文档、段落、标题、样式、批注、页眉页脚、目录、文档转换 |
| wps-ppt | 演示与幻灯片，含「组合配方」一节（KPI 卡片/时间线/流程图等由工具组合而成） |

每个技能目录下另有一份由注册表生成的 `reference.md`，列出全部工具、所属档位与派发方式。

## 性能

旧实现每次工具调用都要新起一个 PowerShell 进程并重新解析脚本。现在是一个常驻宿主
（`powershell.exe` 5.1 + `-STA`）通过 stdin/stdout 收发 JSON 行，工具层只做转发：

| 调用 | 改造前 | 改造后（warm） |
|---|---|---|
| wps_common_ping | 969ms | 1ms |
| wps_common_wire_check | 828ms | 1ms |
| wps_common_get_app_info | 1858ms | 2ms |
| 读 8000 格区域 | 8000 次 COM 往返 | 15–37ms（一次 `Range.Value2` + 二维编组） |

宿主首次调用约 1.0s（进程启动 + 259 个 action 预热），之后 1–2ms 稳态；宿主被 kill 或超时后下次调用自动重启。

## 质量门禁

装好后先跑一次环境自检（只读，不改任何东西；发现 ERROR 会以非零码退出）：

    node scripts/doctor.mjs                        # 平台 / Node / PowerShell STA / 包内容 / WPS COM / profile 接线

改完代码按这个顺序复跑（前三条不需要 WPS）：

    cd mcp && ./node_modules/.bin/tsc              # 类型检查 + 构建（dist 已入库）
    powershell -NoProfile -File scripts\build-host-actions.ps1
                                                   # 由 mcp/scripts/wps-com.ps1 生成 host/wps-actions.ps1
                                                   # 生成前会解析校验，写不出可加载的模块就直接失败
    node scripts/extract-spec.mjs                  # 从当前工具面 bootstrap 操作规格（迁移期用一次）
    node scripts/gen-tool-surface.mjs              # spec → spec/*.json（工具面、键表、广告集、签名索引）
    node scripts/gen-skill-tools.mjs               # 由注册表重生成 4 份 reference.md
    node scripts/verify.mjs                        # 23 项：门面、派发、预算、action 数量三方一致
    node scripts/param-contract.mjs                # 211 对参数契约对账，写出 docs/param-contract.md
    node test/xxx.test.mjs                         # 逐文件跑；需要本机装好 WPS
    node scripts/e2e.mjs --profile <name>          # 一键端到端验收；需要该 profile 已装本包 + 本机 WPS

端到端验收是**一条命令**：`node scripts/e2e.mjs --profile <name>` 会自己造 fixture 工作簿（本场景两份）
（裸 COM，刻意不走本插件，免得用具自己的 bug 伪造输入）→ 跑一个真实 headless 任务 → 逐帧解会话日志
打印工具调用轨迹 → 用裸 COM 重开产物核对内容 → 断言「没有残留文档」「操作结果里没有缺陷标记」
「模型没有自己写 COM 脚本」。**28 项检查、约 95 秒**；第二场景（P2-5）要求把另一份订单工作簿变成真正的表、加条件格式并设成可打印，见 docs/FIXES.md 第 44 条。profile 不存在时加 `--setup` 一步建好并安装本包。
轨迹、产物与 `report.json` 留在 `test/.artifacts/e2e/<run>/`。

当前数字：**443 项测试**（21 个文件，含 spec 复现验收 12 项）+ **23 项门禁**全绿；广告面
**59 工具 / 32,924 字节**（D1 上限 70 / 40,000）；桥 action **256**（生成器断言源码、生成物、期望值三方一致）；
注册工具 **252**，全部进操作规格（spec）。

### CI 覆盖到哪

`.github/workflows/ci.yml`（GitHub Actions，`windows-latest`）**只跑静态那部分**：

1. `npm ci` + `tsc` 构建，再断言 `mcp/dist` 与源码一致（dist 是入库的，不许过期）；
2. 跑宿主生成器，再断言 `host/wps-actions.ps1` 与桥源码一致（生成器自带解析校验）；
3. 由操作规格重生成工具面，再断言 `spec/*.json` 与 spec 一致；跑 `test/spec-reproduction.test.mjs`
   —— 12 项，含「209 个 schema 与序列化字节数与活体完全一致」「每个参数都有明确去向」
   以及「每个 pass-through 改名都在桥侧有声明」；
4. 重生成技能参考表，再断言 `skills/**/reference.md` 与注册表一致；
5. `node scripts/verify.mjs --static` —— 18 项：广告面、预算、桥 action 数量、`wps_help` 检索与派发守卫；
6. 参数契约对账，再断言 `docs/param-contract.md` 一致；
7. 两个不需要 WPS 的测试文件：`test/plugin.test.mjs`(32 项)、`test/com-host.test.mjs`(6 项)。

首次运行（2026-09-13，[run 34760241577](https://github.com/sueccku/dsh-plugin-wps-office-next/actions/runs/34760241577)）**13 步全绿、67 秒**。

需要真实 WPS 的测试与一键 e2e **留在本机**——没有自托管 runner，也不打算为了 CI 去装 WPS。
`verify.mjs` 的另外 5 项（`wps_status`、真实派发、`wps_batch`）只有完整模式（不带 `--static`）才会跑。

| 测试文件 | 项数 | 覆盖 |
|---|---|---|
| plugin.test.mjs | 32 | DSH 入口、技能注册、工具面过滤、预算 |
| word-lifecycle.test.mjs | 18 | e2e 暴露的 5 个缺陷（检索/列表/文档信息/公式/Word 建关） |
| excel-contract-fixes.test.mjs | 35 | 25 处参数错配逐个真实验证、跨应用批注污染 |
| ppt-contract-fixes.test.mjs | 57 | PPT 参数契约逐项修复 |
| merged-tools.test.mjs | 31 | 重复工具合并后的转发、改名与隐藏 |
| new-actions.test.mjs | 28 | 补齐的 action、closePresentation |
| sheet-ops.test.mjs | 21 | 工作表组目标、0 基 position、删表安全 |
| cell-format.test.mjs | 16 | 10 个格式属性逐个回读、平铺写法、空格式报错 |
| close-safety.test.mjs | 14 | 关闭不得弹模态框、不得泄漏工作簿/演示文稿 |
| find-replace.test.mjs | 14 | 查找替换的参数名契约与替换模式 |
| excel-range.test.mjs | 12 | sheet 回退、二维编组、单格、8000 格性能 |
| file-ops.test.mjs | 11 | 另存/打开的路径、按应用转换、页眉页脚分节 |
| deprecated.test.mjs | 8 | 弃用名合并、隐藏与转发 |
| warnings.test.mjs | 7 | 「尽力而为」失败如实回传 |
| com-host.test.mjs | 6 | 常驻宿主握手、串行队列、崩溃重启 |
| spec-reproduction.test.mjs | 12 | spec 逐字节复现模型可见面（P1 验收） |
| excel-missing-halves.test.mjs | 18 | P2 第一波 Excel 补全工具（真实 WPS） |
| excel-missing-halves-2.test.mjs | 30 | P2 第一波余项：格式、校验、重算、合并计算、分类汇总（真实 WPS） |
| excel-list-object.test.mjs | 25 | P2-2 表（ListObject）全族：建表/读结构/增删行/总计行/样式/范围/转回区域（真实 WPS） |
| excel-page-setup.test.mjs | 24 | P2-3 页面设置/打印标题/页眉页脚/可见性与标签色/分级显示/公式审计（真实 WPS） |
| excel-advanced.test.mjs | 24 | P2-4 透视表/全部刷新/单变量求解/迷你图/图表标题与删除（真实 WPS） |

## 目录结构

| 路径 | 用途 |
|---|---|
| cordis.patch.yml | DSH 接线：插件入口 + MCP 客户端（serverName 为 wps-office-next） |
| plugin.js | DSH 入口：发布包内绝对路径、注册 4 个技能 |
| mcp/src | MCP server（TypeScript）与工具实现、参数闸门、弃用表 |
| mcp/src/spec | **操作规格（唯一真源）**：工具名 ↔ 桥 action ↔ 参数/类型/必填/效果 |
| spec/ | 由 spec 生成的产物（入库、CI 漂移检查）：工具面、action 键表、广告集、紧凑签名索引、别名/容器/辅助键/动态动作声明 |
| mcp/dist | 预构建产物（已入库，安装即用） |
| mcp/scripts/wps-com.ps1 | **桥的唯一真源**：259 个 COM action |
| host/ | 常驻 COM 宿主 + 生成物 `wps-actions.ps1`（不要手改，改桥源码后重跑生成器） |
| skills/ | 4 个技能文档 + 生成的 reference.md |
| scripts/ | doctor、verify、参数契约、一键 e2e、生成器、工具面分析、延迟基准 |
| test/ | 21 个回归测试文件、443 项断言（多数需要本机 WPS） |
| baseline/ | 上游基线快照与 28 行缺陷清单（逐条标注 已修/部分修/仍开/不修） |
| docs/ | FIXES（修复记录）、PROGRESS（进度）、param-contract（生成的契约报告） |

## 文档地图

| 文档 | 内容 |
|---|---|
| [docs/FIXES.md](docs/FIXES.md) | 29 条修复记录，每条都带可复跑的验证方法与实测数字 |
| [docs/tool-roadmap.md](docs/tool-roadmap.md) | 工具面路线图：能力审计、缺口/冗余清单、分阶段任务表（P0–P5） |
| [docs/PROGRESS.md](docs/PROGRESS.md) | 分阶段进度与当前状态 |
| [baseline/known-defects.md](baseline/known-defects.md) | 上游 28 行缺陷清单 + 本仓库逐条状态 |
| [docs/param-contract.md](docs/param-contract.md) | 211 对工具/action 参数契约对账（生成物） |
| [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) | 上游来源与许可处理 |

## 已知限制

- **仅 Windows x64 + WPS**：macOS / Linux 支持已整体移除；不兼容 x86，不兼容多组件模式；
- **不做卸载/安装加载项**：走纯 COM 路径，不需要加载项；
- **不做「撤销」承诺**：COM 改动不一定进撤销栈，重要文档请先保存副本；
- **无人值守优先**：关闭文档时对「从未落盘」的文件会自动改为不保存关闭并回报 `warning`，避免弹出保存对话框卡死；
- 同一台机器上多个进程同时驱动 WPS 会互相干扰（共用一个 WPS 实例），自动化测试不要与交互使用并行；
- 结果里的 `warnings` 表示「操作已完成，但某个尽力而为的步骤失败了」，需要如实转述给用户。

## 已锁定的项目决策

| 决策 | 取值 |
|---|---|
| 平台 | 仅 Windows，仅 COM |
| 目标环境 | WPS 12.1+ x64 |
| 形态 | 单一 npm 包 = DSH bundle + 自带 MCP server + 自带 COM host + 全部 skills |
| MCP serverName | wps-office-next |
| 默认工具面 | standard 档 56 个工具；预算上限 70 工具 / 40,000 字节（D1） |
| 传输层 | 常驻 PowerShell STA 宿主 + stdin/stdout JSON 行 |
| 加载项 | 全部删除，零依赖 |
| 构建产物 | 预构建产物入库，安装后开箱可用 |
| 文档语言 | 中文 |
| 提交身份 | sueccku |

## 许可与来源

本仓库使用 MIT License。上游为 MIT，来源与处理方式见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

# dsh-plugin-wps-office-next

[![ci](https://github.com/sueccku/dsh-plugin-wps-office-next/actions/workflows/ci.yml/badge.svg)](https://github.com/sueccku/dsh-plugin-wps-office-next/actions/workflows/ci.yml)
![advertised tools](https://img.shields.io/badge/advertised%20tools-69%20%2F%20267-blue)
![platform](https://img.shields.io/badge/platform-Windows%20x64%20%C2%B7%20WPS%2012.1%2B-informational)
![license](https://img.shields.io/badge/license-MIT-green)

在 **DeepSeek Harness（DSH）** 里，用中文说话就能操作 **WPS 表格 / 文字 / 演示**。

装好之后，你可以直接提这样的要求：

> 把桌面上的《销售明细.xlsx》按地区汇总到一张新工作表，降序排列，再画一张柱状图，最后导出成 PDF。

DSH 会自己打开 WPS、一步步做完、保存文件，再把结果告诉你。**你不需要会编程，也不需要记任何命令。**

- 不用给 WPS 装加载项；
- 不用配置任何环境变量；
- 不用写代码——用中文把事情说清楚就行。

> 本插件只负责「操作 WPS」这一件事。所有改动都发生在你本机已经打开的 WPS 里，插件本身不联网、不会把文件传到别处。
> 真正会看到文档内容的是你正在使用的 AI 模型，敏感文件请按你所用的模型服务的隐私政策自行判断。

---

## 它适合谁

- **每天都在 WPS 里做重复操作的人**：整理表格、批量改格式、按模板生成文档、把 Excel 数据做成 Word 报告或 PPT。
- **已经用 DSH 处理文字工作的人**：希望 AI 直接动手改文件，而不是只给一段「你可以这样操作」的说明。
- **不想学 Office 自动化的人**：VBA、Python、PowerShell 脚本统统不用碰。

## 能帮你做什么

### 表格（WPS 表格 / Excel）

- 读取、写入、批量填充单元格区域；跨工作表搬运数据
- 排序、筛选、去重、分列、分类汇总、合并计算
- 公式与重新计算、单变量求解（从目标倒推输入）
- 单元格格式、数字格式、条件格式、数据验证
- 图表（含标题与数据标签）、透视表（创建 / 刷新 / 清除）
- 表格对象、命名范围、批注、冻结窗格、分级显示
- 页面设置与打印（打印标题、页眉页脚、横向 A4）
- 导出图片、转 PDF、另存为其他格式

### 文字（WPS 文字 / Word）

- 新建、打开、保存、另存、关闭文档
- 读写正文与段落；设置字体、样式、行距、对齐
- 查找与替换（「只查找」和「替换」是两种明确的行为，不会误改文档）
- 表格的读写与编辑（增删行列、合并拆分、套用样式）
- 页眉页脚、页码、分节、分栏、页面设置
- 目录、书签、超链接、脚注尾注、索引、交叉引用、内容控件
- 批注与修订（查看、接受 / 拒绝）、文档字数统计
- 按 CSV 数据做**邮件合并**：一次批量生成几十份新文档，母版不动

### 演示（WPS 演示 / PPT）

- 新建、打开、保存演示文稿
- 增删幻灯片、套用版式、设置标题与正文
- 文本框与形状、填充与字体、形状效果
- 插入表格、图片；统一表格样式
- 切换与动画、演讲者备注
- 把幻灯片导出成图片

### 跨应用与通用

- **在三个应用之间搬运数据**：例如「Excel 里算好的数据写进 Word 报告，再导出 PDF」
- 保存 / 另存为、格式互转（含转 PDF）
- 查看当前打开了哪些工作簿 / 文档 / 演示，以及连接状态
- 批量执行（一次最多 50 个操作），减少来回等待

> 完整工具清单见 [skills/](skills/) 各目录下的 reference.md；在 DSH 里也可以直接问「WPS 能做哪些事」。

## 你可以这样提要求

下面都是可以直接照抄、改一改就能用的例子：

| 想做的事 | 就这样说 |
|---|---|
| 汇总表格 | 把桌面《销售明细.xlsx》按地区汇总到新工作表，降序，并加一张簇状柱形图 |
| 排版与导出 | 这张表表头加粗、冻结首行、数字用千分位，然后转成 PDF 放到同一个文件夹 |
| 合并多份文件 | 把这个文件夹里所有 .xlsx 的第一张工作表合并成一张总表 |
| 改文档不改原件 | 读我打开的这份 Word，把所有「待定」换成「已确认」，另存一份副本，不要改原件 |
| 生成 PPT | 把这份报告的大纲做成 8 页 PPT，每页一个要点 |
| 批量套模板 | 把《合同模板.docx》里的 {甲方}、{金额} 换成这张表里的数据，逐行生成 30 份，并都导出 PDF |
| 清理收尾 | 看看我现在开着哪些 WPS 文件，没保存的先帮我存到桌面 |

## 安装前准备

请先对照下面这张表确认环境（缺一项都装不起来）：

| 项目 | 要求 | 怎么确认 |
|---|---|---|
| 操作系统 | **64 位 Windows 10 / 11** | 设置 → 系统 → 关于 → 看「系统类型」是否有「64 位」 |
| WPS Office | **12.1 或更高版本，且为 64 位** | 打开 WPS →「设置 / 关于」看版本号；不支持 32 位，也不支持「多组件模式」安装 |
| Node.js | **22.19 或更高** | 在 PowerShell 里运行 `node --version` |
| pnpm | 任意较新版本 | 在 PowerShell 里运行 `pnpm --version`；没有就 `npm install -g pnpm` |
| DeepSeek Harness | 能正常启动 | 在 PowerShell 里运行 `dsh --version` |

> Windows PowerShell 5.1 是 Windows 10 / 11 自带的，**不需要另外安装**。
> 还没装 Node.js 的话，去 https://nodejs.org 下载 LTS 版安装即可（会一并装上 npm）。

## 安装（一步一步来）

全程只要打开一次 PowerShell 窗口，把命令复制进去、回车就行。

### 第 0 步：完全关闭 DSH

关掉浏览器里的 DSH 页面或桌面窗口，并确认后台没有 DSH 在运行。
插件是在 **DSH 启动时加载**的，所以装完必须重启它才会生效。

### 第 1 步：打开 PowerShell

开始菜单搜索 `PowerShell`，打开「Windows PowerShell」。**不需要管理员权限**。

### 第 2 步：确认 pnpm 可用

输入：

```powershell
pnpm --version
```

- 能显示版本号（例如 `10.x` / `12.x`）→ 继续下一步。
- 提示「不是内部或外部命令」→ 先运行 `npm install -g pnpm`，再回来确认。

### 第 3 步：安装插件

```powershell
dsh plugin --profile web add github:sueccku/dsh-plugin-wps-office-next#v0.2.0
```

关于 `--profile web`：

| 你平时怎么用 DSH | 就用哪个 profile |
|---|---|
| 在浏览器 / 桌面窗口里聊天（最常见） | `web` |
| 在命令行里跑一次性任务 | `headless` |
| 自己起过别的名字 | 换成那个名字 |

命令会联网下载插件，然后输出 pnpm 的安装日志，最后回到命令提示符。首次为某个 profile 装插件时，
DSH 可能提示 `initialized profile ...`，这是正常的。

> 如果你是把本仓库克隆到本地来用，也可以直接用本地路径安装：
> `dsh plugin --profile web add D:\一些目录\dsh-plugin-wps-office-next`（改成你自己的实际路径）。

### 第 4 步：重新启动 DSH

按你平时的方式重新打开 DSH（例如重新运行 `dsh web`，或双击你平时的启动方式）。**不重启，插件不会生效。**

### 第 5 步：验证是否成功

1. 先启动 **WPS**，随便打开一个表格 / 文档 / 演示（`connected` 为 true 需要 WPS 正在运行）。
2. 在 DSH 里新建一个会话，发一句话：

   > 调用 wps_status，告诉我 WPS 连上了没有。

3. 预期能看到类似这样的结果：

   ```
   connected: true
   advertisedTools: 69
   registeredTools: 267
   ```

看到 `connected: true` 就说明装好了，可以开始用了。**第一次调用**可能慢 1 秒左右，之后每次只要 1–2 毫秒。

## 如果哪里不对

| 现象 | 原因 | 怎么办 |
|---|---|---|
| `pnpm not found on PATH` | 没装 pnpm | 运行 `npm install -g pnpm` 后重试 |
| `dsh` 不是内部或外部命令 | DSH 没装或不在 PATH | 确认 DSH 已正确安装，重开一个 PowerShell 窗口 |
| 装完在 DSH 里毫无反应 | 没有重启 DSH | 完全关闭再重新启动 DSH |
| `connected: false` | WPS 没启动，或没打开任何文档 | 启动 WPS 12.1+（64 位）并打开一个文件，然后重试；不要反复空转重试 |
| 提示平台 / Node 版本不符 | 系统或 Node 太旧 | 检查是否 64 位 Windows、`node --version` 是否 ≥ 22.19 |
| 提示 WPS 是 32 位 / 多组件模式 | 环境不支持 | 换成 64 位完整安装的 WPS |
| 操作时而成功时而失败 | 多个程序同时在驱动 WPS | 关闭多余的 DSH 会话或其它自动化工具 |

### 完整环境自检

想一次把所有环境项都查清楚，可以在安装目录里运行自检脚本（只读，不改任何东西）：

```powershell
cd "$env:DSH_HOME\profiles\web\node_modules\dsh-plugin-wps-office-next"
node scripts\doctor.mjs
```

它会逐项检查系统位数、Node 版本、PowerShell STA、包内容是否完整、WPS 能否连上，
最后打印 `DOCTOR OK` 或指出具体的错误。

## 卸载

```powershell
dsh plugin --profile web remove dsh-plugin-wps-office-next
```

然后重启 DSH。插件**不会往 WPS 里装任何东西**，所以卸载它不影响 WPS 本身。

## 重要注意事项

- **动手前先备份重要文件。** COM 改动不一定能进撤销栈，Ctrl+Z 不一定能恢复；批量改动前请先复制一份。
- **结果里的 `warnings` 要看清。** 它表示「主要操作已经完成，但某个次要步骤失败了」，通常需要你人工确认一下。
- **关闭「从未保存过」的文件时，插件会自动选择「不保存并关闭」**（并回报 warning），以免弹出保存对话框把流程卡死。
  有重要内容时请自己手动保存。
- **同一时间只让一个程序操作 WPS。** 多个 DSH 会话、或者你一边手动操作一边让 AI 操作，会互相干扰（它们共用同一个 WPS 实例）。
- **结果以文件为准。** 插件会在改动后复验，但涉及重要数据时，建议自己再打开确认一遍。

## 已知限制

- **仅 Windows x64 + WPS 12.1+（64 位）**：macOS / Linux 支持已整体移除；不兼容 32 位 WPS，也不兼容「多组件模式」。
- **不安装 / 卸载 WPS 加载项**：走纯 COM 路径，本来就不需要加载项。
- **不做「撤销」承诺**：COM 改动不一定进撤销栈。
- **以下能力受 WPS 自身限制，实测无法实现**（如实记录，不再尝试）：
  - 水印：WPS 页眉的 Shapes 集合不接受任何图形；
  - 文档属性（内置 / 自定义）：`BuiltInDocumentProperties` / `CustomDocumentProperties` 是坏壳；
  - 切片器：能建出缓存，但 `Slicers.Count` 恒为 0，用户可见的切片器不会出现；
  - 场景管理器：`Worksheet.Scenarios` 在 COM 里被暴露成方法，语义读不干净。
  - 以上可尝试隐藏逃生舱 `wps_execute_method`，但不保证成功。

---

## 进阶：技术细节

<details>
<summary><b>工具面：注册 267 个，每次请求只广告 69 个</b></summary>

一次请求塞几百个工具会浪费上下文。本插件把 `tools/list` 收敛为三档，由环境变量 `WPS_OFFICE_TOOLSET` 切换（默认 `standard`）：

| 档位 | 工具数 | schema 字节 | 约 tokens | 内容 |
|---|---|---|---|---|
| minimal | 4 | 1,348 | 385 | 仅 4 个门面工具 |
| **standard（默认）** | **69** | **37,573** | **10,735** | 门面 + 65 个精选工具 |
| full | 267 | 153,777 | 43,936 | 全量，保留完整描述 |

注册目录 **267** 个工具：Excel 118 / Word 59 / PPT 76 / 通用 14（含门面）。

**未广告的工具完全可用**，两条路都能走：按全名直接调用，或先查后调。

| 门面工具 | 作用 |
|---|---|
| wps_status | 连接状态 + 活动应用 + 当前档位 + 广告/注册/隐藏数量；编辑前先调用 |
| wps_help | 无参看分组概览；传 `app` 列该应用目录；传 `query` 模糊搜索；传 `tool` 取完整 inputSchema |
| wps_call | 执行任意已注册但未广告的工具 |
| wps_batch | 顺序批量执行，单次上限 50 项 |

配套两道契约防线：工具层发给桥的多余参数会被拒绝并列出可接受的键；
`scripts/param-contract.mjs` 零副作用地把 255 对工具/action 的参数契约对账一遍，
结果写入 [docs/param-contract.md](docs/param-contract.md)。

</details>

<details>
<summary><b>性能：常驻 COM 宿主</b></summary>

旧实现每次工具调用都要新起一个 PowerShell 进程并重新解析脚本。现在是一个常驻宿主
（`powershell.exe` 5.1 + `-STA`）通过 stdin/stdout 收发 JSON 行，工具层只做转发：

| 调用 | 改造前 | 改造后（warm） |
|---|---|---|
| wps_common_ping | 969ms | 1ms |
| wps_common_wire_check | 828ms | 1ms |
| wps_common_get_app_info | 1858ms | 2ms |
| 读 8000 格区域 | 8000 次 COM 往返 | 15–37ms（一次 `Range.Value2` + 二维编组） |

宿主首次调用约 1.0s（进程启动 + 267 个 action 预热），之后 1–2ms 稳态；宿主被 kill 或超时后下次调用自动重启。

</details>

<details>
<summary><b>技能（Skills）</b></summary>

`plugin.js` 注册 4 个技能，DSH 会按任务自动加载（任务是跨应用时先加载路由技能）：

| 技能 | 触发场景 |
|---|---|
| wps-office-next | 路由与通用约定：先 `wps_status`、参数名以 schema 为准、变更后复验、如何如实转达 `warnings` |
| wps-excel | 表格、区域、格式、公式、图表、透视表、数据清洗 |
| wps-word | 文档、段落、标题、样式、批注、页眉页脚、目录、文档转换 |
| wps-ppt | 演示与幻灯片，含「组合配方」一节（KPI 卡片 / 时间线 / 流程图等由工具组合而成） |

每个技能目录下另有一份由注册表生成的 `reference.md`，列出全部工具、所属档位与派发方式。

</details>

<details>
<summary><b>质量门禁与 CI</b></summary>

装好后先跑一次环境自检（只读，不改任何东西；发现 ERROR 会以非零码退出）：

```powershell
node scripts\doctor.mjs
```

改完代码按这个顺序复跑（前几条不需要 WPS）：

```powershell
# 1) 类型检查 + 构建（dist 已入库，构建产物要与源码一致）
cd mcp; npx tsc; cd ..

# 2) 由桥源码重新生成 host/wps-actions.ps1（生成前会解析校验）
powershell -NoProfile -File scripts\build-host-actions.ps1

# 3) 由操作规格重新生成 spec/*.json，再由注册表重生成 4 份 reference.md
node scripts\gen-tool-surface.mjs
node scripts\gen-skill-tools.mjs

# 4) 门禁：门面、派发、预算、action 数量三方一致 + 参数契约对账
node scripts\verify.mjs
node scripts\param-contract.mjs

# 5) 需要本机 WPS 的部分
node test\xxx.test.mjs                 # 逐文件跑
node scripts\e2e.mjs --profile <name>  # 一键端到端验收
```

当前数字：**541 项测试（25 个文件，多数需要真实 WPS）+ verify 23 项 + spec 复现 12 项**全绿；
广告面 69 工具 / 37,573 字节（内部预算上限 70 / 40,000）；桥 action 267，与注册表三方一致；
参数契约 255 对，四类静默失效均为 0。

`.github/workflows/ci.yml`（GitHub Actions，`windows-latest`）**只跑不需要 WPS 的静态部分**：
tsc 构建并对账 `mcp/dist`、重生成宿主并对账、重生成 spec 并对账、重生成技能参考表并对账、
`verify --static`、参数契约对账、以及两个纯静态测试文件。需要真实 WPS 的测试与一键 e2e 留在本机。

一键 e2e 是**一条命令**：`node scripts/e2e.mjs --profile <name>` 会自己造 fixture 工作簿（裸 COM，刻意不走本插件）
→ 跑一个真实 headless 任务 → 逐帧解会话日志打印工具调用轨迹 → 用裸 COM 重开产物核对内容 →
断言「没有残留文档」「结果里没有缺陷标记」「模型没有自己写 COM 脚本」。**28 项检查、约 95 秒。**
轨迹、产物与 `report.json` 留在 `test/.artifacts/e2e/<run>/`。

</details>

<details>
<summary><b>目录结构</b></summary>

| 路径 | 用途 |
|---|---|
| cordis.patch.yml | DSH 接线：插件入口 + MCP 客户端（serverName 为 wps-office-next） |
| plugin.js | DSH 入口：发布包内绝对路径、注册 4 个技能 |
| mcp/src | MCP server（TypeScript）与工具实现、参数闸门、弃用表 |
| mcp/src/spec | **操作规格（唯一真源）**：工具名 ↔ 桥 action ↔ 参数/类型/必填/效果 |
| spec/ | 由 spec 生成的产物（入库、CI 漂移检查） |
| mcp/dist | 预构建产物（已入库，安装即用） |
| mcp/scripts/wps-com.ps1 | **桥的唯一真源**：267 个 COM action |
| host/ | 常驻 COM 宿主 + 生成物 `wps-actions.ps1`（不要手改） |
| skills/ | 4 个技能文档 + 生成的 reference.md |
| scripts/ | doctor、verify、参数契约、一键 e2e、生成器、分析工具 |
| test/ | 25 个回归测试文件（多数需要本机 WPS） |
| baseline/ | 早期基线快照与缺陷清单 |
| docs/ | FIXES（修复记录）、PROGRESS（进度）、param-contract（生成的契约报告） |

</details>

<details>
<summary><b>文档地图</b></summary>

| 文档 | 内容 |
|---|---|
| [CHANGELOG.md](CHANGELOG.md) | 每个发布版本的用户可见变化 |
| [docs/FIXES.md](docs/FIXES.md) | 修复记录，每条都带可复跑的验证方法与实测数字 |
| [docs/tool-roadmap.md](docs/tool-roadmap.md) | 工具面路线图：能力审计、缺口 / 冗余清单、分阶段任务表 |
| [docs/PROGRESS.md](docs/PROGRESS.md) | 分阶段进度与当前状态 |
| [docs/param-contract.md](docs/param-contract.md) | 工具 / action 参数契约对账（生成物） |
| [baseline/known-defects.md](baseline/known-defects.md) | 早期基线缺陷清单与逐条状态 |
| [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) | 来源与许可处理 |

</details>

<details>
<summary><b>已锁定的项目决策</b></summary>

| 决策 | 取值 |
|---|---|
| 平台 | 仅 Windows，仅 COM |
| 目标环境 | WPS 12.1+ x64 |
| 形态 | 单一 npm 包 = DSH bundle + 自带 MCP server + 自带 COM host + 全部 skills |
| MCP serverName | wps-office-next |
| 默认工具面 | standard 档 69 个工具；预算上限 70 工具 / 40,000 字节 |
| 传输层 | 常驻 PowerShell STA 宿主 + stdin/stdout JSON 行 |
| 加载项 | 全部删除，零依赖 |
| 构建产物 | 预构建产物入库，安装后开箱可用 |
| 文档语言 | 中文 |
| 提交身份 | sueccku |

</details>

## 许可

本仓库使用 **MIT License**，见 [LICENSE](LICENSE)。

## 致谢

本项目的底层实现起步于两个早期的开源项目，感谢它们的作者：

- [lc2panda/wps-skills](https://github.com/lc2panda/wps-skills) —— WPS COM 操作实现与 MCP 工具层；
- [CatNebulaaaa/wps-dsh-plugin](https://github.com/CatNebulaaaa/wps-dsh-plugin) —— DSH bundle 的接线形态。

二者均为 MIT 许可，来源与处理方式见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
经过多轮重构与扩展，本项目的架构、工具面与代码已与上游有很大差异，是独立演进的结果。

本项目为非官方项目，与金山办公、DeepSeek 及上述上游作者均无隶属关系。

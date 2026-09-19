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

> **安装很简单**：见下面的「安装」一节——把那一段话整段复制粘贴给你的 DSH AI，剩下的交给它。

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

## 安装：把下面这段话复制给你的 DSH AI

你**不需要自己敲命令**。打开 DSH 新建一个会话，把下面整段复制粘贴进去发送，剩下的交给它：

```text
请帮我在本机安装 DeepSeek Harness 插件 dsh-plugin-wps-office-next
（仓库：github.com/sueccku/dsh-plugin-wps-office-next）。

用途：装上之后，我可以用中文让你操作本机的 WPS 表格 / 文字 / 演示
（读写数据、排版、做图表、转 PDF 等）。

请按下面的步骤执行，遇到不确定的地方先问我：

1. 环境检查：本插件只支持 64 位 Windows 10/11 + WPS Office 12.1+（64 位）；
   还需要 Node.js ≥ 22.19 和 pnpm。不满足的项目直接告诉我，不要硬装。
2. 判断我当前正在使用的 DSH profile（我从浏览器/桌面界面使用 DSH，通常是 web）。
3. 执行安装（把 <profile> 换成上一步判断出的名字）：
   dsh plugin --profile <profile> add github:sueccku/dsh-plugin-wps-office-next#v0.3.0
   如果报 ERR_PNPM_GIT_RESOLVE_FAILED 或连不上 github.com，改用这个地址重试：
   dsh plugin --profile <profile> add https://codeload.github.com/sueccku/dsh-plugin-wps-office-next/tar.gz/refs/tags/v0.3.0
4. 用下面这条命令确认接线成功（输出里应出现 wps-office-next-plugin 与 mcp-wps-office-next）：
   dsh --profile <profile> --dump-config | Select-String wps
5. 然后告诉我「请完全关闭并重新打开 DSH，回来再让我验证」。
   插件是在 DSH 启动时加载的，必须重启才会生效。
6. 我重启回来后，先提醒我启动 WPS 并打开一个表格/文档，然后调用 wps_status，
   确认返回 connected: true（此时 advertisedTools 应为 69，registeredTools 应为 267）。

安装、排错与卸载的完整说明见 README 的「给 AI 的安装指引」一节：
https://github.com/sueccku/dsh-plugin-wps-office-next#给-ai-的安装指引
```

> **唯一需要你亲自做的一步**：AI 装完后，按它说的**完全关闭再重新打开 DSH**。
> 重启后新建一个会话，问一句「WPS 连上了吗」即可。

动手前顺手确认两件事就够了：系统是 **64 位 Windows**，WPS 是 **12.1 及以上的 64 位版本**。
其余环境（Node.js、pnpm 等）AI 会自己检查，缺什么它会告诉你怎么办。

## 给 AI 的安装指引

> 这一节是写给你 DSH 里的那个 AI 看的。用户把上一节那段话发给你之后，按这里执行；
> 用户也可以把整节内容一起发给你。

### 1. 环境要求

| 项目 | 要求 | 检查方式 |
|---|---|---|
| 操作系统 | 64 位 Windows 10 / 11 | `(Get-CimInstance Win32_OperatingSystem).OSArchitecture` 为 `64-bit` |
| WPS Office | 12.1 或更高，64 位 | 能创建 COM 对象 `Ket.Application` / `Kwps.Application` / `Kwpp.Application` |
| Node.js | ≥ 22.19 | `node --version` |
| pnpm | 任意较新版本 | `pnpm --version`；缺失则 `npm install -g pnpm` |
| DSH | 能正常启动 | `dsh --version` |

Windows PowerShell 5.1 是 Windows 自带的，不需要安装。硬性条件不满足时如实告诉用户，不要硬装。

### 2. 判断要装进哪个 profile

插件必须装进**用户当前正在使用的那个 profile**，否则重启后没有任何效果。按顺序尝试：

```powershell
# 从正在运行的 dsh 进程命令行里读出 profile
$cl = (Get-CimInstance Win32_Process -Filter "Name='node.exe'" |
        Where-Object { $_.CommandLine -match 'dsh.*bin\.js' } |
        Select-Object -First 1).CommandLine
if ($cl -match '--profile\s+(\S+)') { $profile = $Matches[1] }
elseif ($cl -match 'bin\.js"?\s+(\S+)') { $profile = $Matches[1] }
$profile
```

- 常见结果：`web`（浏览器 / 桌面界面，最常见）、`headless`（命令行跑一次性任务）。
- 读不到时：若 `$env:DSH_WEB_URL` 有值，基本可以判定是 `web`。
- 仍不确定：看 `$env:DSH_HOME\profiles` 下有哪些目录，或者直接问用户「你平时是怎么启动 DSH 的」。
- `dsh plugin` 会把不存在的 profile 自动初始化，所以**装错 profile 不会报错，只会"装完没反应"**，
  宁可多问用户一句。

### 3. 执行安装

```powershell
dsh plugin --profile <profile> add github:sueccku/dsh-plugin-wps-office-next#v0.3.0
```

- 包内已带预构建产物，**没有任何构建步骤**，正常不需要额外处理，也不会出现"构建脚本被拦截"的问题。
- 报 `ERR_PNPM_GIT_RESOLVE_FAILED`、`could not connect to server`、连接超时等，是这台机器访问
  github.com 不稳定，改用 GitHub 打包地址重试（已实测可行）：

  ```powershell
  dsh plugin --profile <profile> add https://codeload.github.com/sueccku/dsh-plugin-wps-office-next/tar.gz/refs/tags/v0.3.0
  ```

- **注意 dsh 事后那句「构建脚本被拦截 / allowBuilds」提示**：本包没有任何 `prepare` 脚本，也没装
  `esbuild` 之类的原生依赖，所以**永远不需要**改 `allowBuilds`。dsh 在 pnpm 失败时会补上这句通用提示，
  照着改只会白费功夫——**先看真正的错误**：如果是 `ERR_PNPM_GIT_RESOLVE_FAILED` / `Failed to connect to
  github.com`，用上面的 codeload 地址重试即可（这一条已实测：4.1 秒装好）。
- 如果因为沙箱 / 权限限制写不了 `$env:DSH_HOME`（受限会话里常见），向用户申请一次授权后重试；
  或者把命令原样交给用户，让他在 PowerShell 里跑一次。
- 如果用本地克隆的仓库安装：`dsh plugin --profile <profile> add <仓库绝对路径>`。

### 4. 确认接线成功（不用重启就能查）

```powershell
dsh --profile <profile> --dump-config | Select-String wps
```

输出里应出现 `wps-office-next-plugin` 和 `mcp-wps-office-next`（分别对应 DSH 插件入口和 MCP 客户端）。
同时确认 profile 的 `package.json` 里 `dsh.profile.bundles` 已包含 `dsh-plugin-wps-office-next`
——`dsh plugin add` 会自动加，不需要手改。

### 5. 让用户重启，然后验证

插件在 DSH 启动时加载，**必须重启**。请明确告诉用户：

> 请完全关闭 DSH 再重新打开（例如关掉窗口后重新运行 `dsh web`），然后回来告诉我。

用户回来之后：

1. 先让他启动 WPS，并打开一个表格 / 文档 / 演示；
2. 调用 `wps_status`，确认 `connected: true`；
3. 正常结果形如：

   ```
   connected: true
   advertisedTools: 69
   registeredTools: 267
   ```

第一次调用可能慢约 1 秒（COM 宿主冷启动），之后 1–2 毫秒，属于正常。

### 6. 排错

| 现象 | 原因 | 怎么办 |
|---|---|---|
| `pnpm not found on PATH` | 没装 pnpm | `npm install -g pnpm` 后重试 |
| `ERR_PNPM_GIT_RESOLVE_FAILED` / 连不上 github.com | 网络抖动 | **先重试一次**（多数情况一次就好）；仍然失败再改用上面的 codeload 打包地址 |
| dsh 提示 `allowBuilds` / 「构建脚本被拦截」 | 那是 dsh 在 pnpm 失败后补的通用提示，不是真因 | 看真错误；本包没有构建步骤，别去改 `allowBuilds` |
| `dsh` 不是内部或外部命令 | DSH 未安装或不在 PATH | 确认 DSH 已安装，重开终端 |
| 装完毫无反应 | 装错 profile，或没重启 | 用第 2 步重新确认 profile；确认已重启 |
| `connected: false` | WPS 没启动 / 没打开文档 | 让用户启动 WPS 12.1+（64 位）并打开一个文件；不要反复空转重试 |
| 平台 / Node 版本不符 | 系统或 Node 太旧 | 需要 64 位 Windows 与 Node ≥ 22.19 |
| WPS 是 32 位 / 多组件模式 | 环境不支持 | 换 64 位完整安装的 WPS |
| 操作时好时坏 | 多程序同时驱动 WPS | 关掉多余的 DSH 会话或自动化工具 |
| 报错「检测到另一个 DSH 会话或自动化程序正在控制 WPS」 | 另一个会话的插件宿主正占着 WPS（同一时间只允许一个） | 关掉那个会话，或在那个会话里结束 WPS 操作后重试 |
| 报错「文件已加密，需要密码」 | 该文件有打开密码，插件不会弹密码框（弹框会卡死会话） | 先用 WPS 手工打开它，另存为一份不加密的副本，再对副本操作 |
| 报错「状态未知：WPS 可能正被一个对话框阻塞」 | 有个模态框在等人工确认，或 WPS 无响应 | 切到 WPS 窗口处理弹框；插件不会替你关掉 WPS |

### 7. 完整环境自检（可选）

```powershell
cd "$env:DSH_HOME\profiles\<profile>\node_modules\dsh-plugin-wps-office-next"
node scripts\doctor.mjs
```

逐项检查系统位数、Node 版本、PowerShell STA、包内容是否完整、WPS 能否连上、**已安装 WPS 的版本与架构**（≥ 12.1 且 64 位）
以及**插件接线**（`cordis.patch.yml` 里是否有 `wps-office-next-plugin` 与 `mcp-wps-office-next` 两个 id），
最后打印 `DOCTOR OK` 或具体错误。

## 卸载

卸载同样可以交给 AI，把下面这句发给它就行：

```text
请帮我卸载 DSH 插件 dsh-plugin-wps-office-next：先判断我当前使用的 profile，
然后执行 dsh plugin --profile <profile> remove dsh-plugin-wps-office-next，
完成后告诉我需要重启 DSH。
```

自己动手的话，就是这一条命令（把 `web` 换成你的 profile），然后重启 DSH：

```powershell
dsh plugin --profile web remove dsh-plugin-wps-office-next
```

插件**不会往 WPS 里装任何东西**，所以卸载它不影响 WPS 本身。

## 重要注意事项

- **动手前先备份重要文件。** COM 改动不一定能进撤销栈，Ctrl+Z 不一定能恢复；批量改动前请先复制一份。
- **结果里的 `warnings` 要看清。** 它表示「主要操作已经完成，但某个次要步骤失败了」，通常需要你人工确认一下。
- **关闭「从未保存过」的文件时，插件会自动选择「不保存并关闭」**（并回报 warning），以免弹出保存对话框把流程卡死。
  有重要内容时请自己手动保存。
- **同一时间只让一个程序操作 WPS。** 多个 DSH 会话、或者你一边手动操作一边让 AI 操作，会互相干扰（它们共用同一个 WPS 实例）。
  现在插件会**主动拦下第二个 DSH 会话**，回一句中文错误（「检测到另一个 DSH 会话或自动化程序正在控制 WPS…」）而不是让两边互相卡住；
  但「你手工操作」和「AI 操作」之间的干扰，仍然只能靠你避免。
- **结果以文件为准。** 插件会在改动后复验，但涉及重要数据时，建议自己再打开确认一遍。

## 已知限制

- **仅 Windows x64 + WPS 12.1+（64 位）**：macOS / Linux 支持已整体移除；不兼容 32 位 WPS，也不兼容「多组件模式」。
- **不安装 / 卸载 WPS 加载项**：走纯 COM 路径，本来就不需要加载项。
- **不做「撤销」承诺**：COM 改动不一定进撤销栈。
- **带密码的文件不会被打开**：打开加密文档时 WPS 会弹一个「文档已加密」的密码框，而模态框一旦弹出就会把整个会话卡住
  （杀宿主也没用，框还在）。插件现在**不弹这个框**，而是**在 1 秒内直接报错**并告诉你怎么办：
  「文件已加密，需要密码…请先用 WPS 手工打开它，另存为一份不加密的副本，再对副本操作。」
  加密的**演示文稿**（.pptx）是唯一没堵住的口子：`Presentations.Open` 没有密码参数，无法阻止弹框。
- **超时后的行为是「状态未知」，不是「已失败」**：如果某个操作超时，插件会告诉你 WPS 可能正被对话框阻塞或已无响应，
  **并且不会自动关闭 WPS**（那会丢掉你没保存的内容）；请切到 WPS 窗口处理弹框后重试。
  之后插件会先用较短的超时快速失败，成功一次即恢复正常。
- **以下能力受 WPS 自身限制，实测无法实现**（如实记录，不再尝试）：
  - 水印：WPS 页眉的 Shapes 集合不接受任何图形；
  - 文档属性（内置 / 自定义）：`BuiltInDocumentProperties` / `CustomDocumentProperties` 是坏壳；
  - 切片器：能建出缓存，但 `Slicers.Count` 恒为 0，用户可见的切片器不会出现；
  - 场景管理器：`Worksheet.Scenarios` 在 COM 里被暴露成方法，语义读不干净。
  - 以上可尝试**隐藏逃生舱** `wps_execute_method`（不在广告面里，技能里也不推荐）：它只是最后手段，不保证成功。

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
`scripts/param-contract.mjs` 零副作用地把 256 对工具/action 的参数契约对账一遍，
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

当前数字：**835 项测试（39 个文件，多数需要真实 WPS）+ verify 23 项 + spec 复现 13 项**全绿；
广告面 69 工具 / 37,573 字节（内部预算上限 70 / 40,000）；桥 action 267，与注册表三方一致；
参数契约 256 对，四类静默失效均为 0。

`.github/workflows/ci.yml`（GitHub Actions，`windows-latest`）**只跑不需要 WPS 的静态部分**：
tsc 构建并对账 `mcp/dist`、重生成宿主并对账、重生成 spec 并对账、重生成技能参考表并对账、
重生成工具覆盖矩阵并对账、`verify --static`、参数契约对账，以及七个不碰真实 WPS 的测试文件
（`plugin` / `com-host` / `host-lease` / `watchdog` / `silent-catch` / `install-selfcheck` / `arg-shape-guard`）。
需要真实 WPS 的测试与一键 e2e 留在本机。

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

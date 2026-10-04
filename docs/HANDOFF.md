# 交接文档（HANDOFF）

> 用途：把当前工作、进展、现状与下一步整理成**自包含**的一页，让一个**全新对话**无需回看历史即可接手。
> 核实时间：2026-10-04（**v0.6.2 发布后重核**；本文所有数字均从仓库/命令实测，非记忆）。
> 现值数字统一看 [current-numbers.md](current-numbers.md)（自动生成 + CI 对账）。
> 本文是工作文档，不随包发布（`docs/` 不在 `package.json` 的 `files` 白名单内）。

---

## 0. 新对话怎么开始

新开对话后，第一句话建议原样发送：

> 读 @docs/HANDOFF.md 和 @docs/current-numbers.md，然后按里面的「新对话怎么开始」做。

### 0.1 先读这五份（按顺序，约 10 分钟）

| # | 文件 | 读它为了知道什么 |
| --- | --- | --- |
| 1 | [HANDOFF.md](HANDOFF.md)（本文） | 现状、决策登记表、契约管线、纪律 |
| 2 | [current-numbers.md](current-numbers.md) | **所有现值的唯一来源**（自动生成 + CI 对账）。别用正则自己数 |
| 3 | [FIXES.md](FIXES.md) 的 84–91 号 | 最近一轮做了什么、踩过哪些坑（含"字段名不匹配""静默失配"这类反复出现的模式） |
| 4 | [stabilization-plan.md](stabilization-plan.md) | 稳定性加固的既定纪律（预算红线、账本只许缩、只读核对优先） |
| 5 | [tool-roadmap.md](tool-roadmap.md) | 还剩哪些规划项没做 |

### 0.2 再做只读核对（确认真实状态与本文一致，**再动手**）

```powershell
Set-Location "D:\dsh\a"
git log --oneline -6 ; git status --porcelain ; git describe --tags
git ls-remote --tags origin | Select-String "v0\.6"        # README 里的回退 tag 必须在这里
node scripts/verify.mjs --static                            # 不需要 WPS：预算/桥动作/契约/长动作（19 项）
node scripts/gen-numbers.mjs --check                        # 文档里的数字是否与生成值一致（20 项）
node scripts/doctor.mjs                                     # 环境自检，末尾应打印 DOCTOR OK
```

> ⚠️ **不要在「正在开着 WPS 干活」的会话里跑需要 WPS 的检查**：宿主的单实例租约会把测试挡在门外
> （`test/host-lease.test.mjs` 会打印中文「另一个 DSH 会话正在控制 WPS」）。这是环境冲突、不是代码缺陷。
> 上面这组**都是只读且不需要 WPS**（`verify.mjs` 不带 `--static` 时才连 WPS、会调一次 `wps_status`）。

### 0.3 需要时再跑整轮回归

跑完整测试（**48 个文件**，多数驱动真实 WPS，约 15 分钟）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\run-tests.ps1   # 跑完会刷新 test/summary.json
node scripts\gen-numbers.mjs          # 再用新计数刷新 docs/current-numbers.md（两步都要）
```

---

## 1. 一句话现状

**已发布 `v0.6.2`（2026-10-04）**：npm `dist-tags.latest = 0.6.2`，git tag 与 GitHub Release 均已就位（详见 §2）。
这一版收口了**假成功与数据破坏**（FIXES 86–91），并把广告面按真实调用数据从 69 重排到 84。

| 指标 | 现值 | 权威来源 |
| --- | --- | --- |
| 注册工具 / 广告面 | **268 / 84**（46,867 字节，预算上限 60,000） | [current-numbers.md](current-numbers.md) |
| 桥 action / 参数契约 | **263 / 257 对**（A/B/C/D 四类静默失效均为 0） | 同上 |
| 测试 | **1077 项 / 48 个文件** | 同上（来自 `test/summary.json`） |
| 本机整轮回归 | **1076 通过 / 1 失败** | 唯一失败是**已知环境问题** `confirm-dialog`（本机 Excel 确认框不弹；已用 `git stash` 证明与代码无关） |

**历史阶段（都是当时的读数，别当现值）**：S1–S9 加固分四波落地，`v0.3.0`（09-16）→ `v0.4.0`（09-19）→
`v0.5.0`–`v0.5.4`（10-02）→ `v0.6.0`（上架 npm，因漏发 `mcp/package.json` 而坏）→ `v0.6.1`（修好分发）→ `v0.6.2`（本版）。
各阶段的原始读数保留在 [PROGRESS.md](PROGRESS.md)（已标注为历史快照）。

---

## 2. 仓库与发布事实

| 项 | 值 |
| --- | --- |
| 仓库根 | `D:\dsh\a` |
| 分支 / HEAD | `main`（以 `git describe --tags` / `git log --oneline -1` 为准）——已推送，与 `origin/main` 一致 |
| 远程 | `https://github.com/sueccku/dsh-plugin-wps-office-next.git` |
| 提交身份 | `sueccku <18247499+sueccku@users.noreply.github.com>` |
| 标签 | `v0.2.0`…`v0.5.4`、`v0.6.1`、**`v0.6.2`**（**注意：没有 `v0.6.0` 标签**——0.6.0 是坏版本，已从 registry 撤销） |
| Releases | **Latest = v0.6.2**（2026-10-04，正文取自 CHANGELOG 的 0.6.2 段）：`gh release create v0.6.2 …` 已执行 |
| 包 | `dsh-plugin-wps-office-next@0.6.2` —— **2026-10-04 已上架 npmjs，`dist-tags.latest = 0.6.2`**（本次发布**不需要 OTP**：本机 token 已可免 2FA 发布，`PUT → 202`，约 2 分钟后 registry 上可见）。实测元数据：`fileCount=296`、`unpackedSize=3,294,700`、`shasum=3d9f2c3a…`、`integrity=sha512-agvvryjeImdzt…jAIKLAwGfqC8w==`（与本地 `npm publish` 打印的逐字符一致）。在架版本 `["0.6.1","0.6.2"]`。依赖 `@modelcontextprotocol/sdk`、`uuid`、`winston` |
| 构建脚本 | `snapshot` / `verify` / `gen:skills` / `gen:coverage` / `lint`——**没有 `prepare`**（安装时不需要构建） |

**工作区**：**干净**（`git status --porcelain` 无输出，与 `origin/main` 一致）。

**本机当前状态（新会话请看这里）**：

- `desktop` profile 里装的**就是 0.6.2**（`…/profiles/desktop/node_modules/dsh-plugin-wps-office-next/package.json` → `0.6.2`），
  与其他 profile（`web`）之外的临时 profile 都已删除。**你若在 DSH 里用插件，跑的就是本 HEAD 这版行为**。
- 但**仓库工作区是权威**：改完源码要重新 `tsc` + 跑管线，装出来的副本不会自动跟着变（除了 `desktop` 依赖的是本地路径时）。

**发布后验收（2026-10-04 实测）**：全新 profile 从 registry 装 `@0.6.2` → **296 文件、`doctor` OK**、`--dump-config` 两个 id 都在。

> ⚠️ **pnpm 元数据缓存（发下一版时会再遇到）**：刚发完新版本时，`dsh plugin add <包名>` 不写版本可能**装到上一版** ——
> pnpm 缓存了旧 packument，而 `dsh plugin add` 把依赖写成 `^上一版`（本次发 0.6.2 时就实测到解析出 `^0.6.1`）。
> 要拿到新版本：写全版本 `add <包名>@X.Y.Z`，或先 `pnpm store prune` / 等缓存过期后重装。
> 这不是包的问题（registry 侧 `dist-tags.latest` 是对的），但**发布后必须这样验证**，否则会误判"发布没生效"。

> 若本节与 registry 或 `git` 现状不符，**以命令实测为准**：`git describe --tags`、`git status --porcelain`、`npm view dsh-plugin-wps-office-next versions`。

---

## 3. 已定决策（不要再翻案）

### 3.1 当前开口的决策（**只有这些**，其余都已定）

| # | 待定事项 | 为什么还开着 | 需要谁决定 |
| --- | --- | --- | --- |
| **A3**（原 D1-E） | 产品扩张 / 跨应用工作流要做成什么样 | 缺**具体场景与目标客户**；没有场景就做等于猜需求 | 用户（给一个真实场景即可启动） |
| **A4** | `param-contract` 那次 2 小时挂起未定性 | 只出现过一次、未能复现；建议下次复现时抓现场（给它加心跳 + 自超时） | 等复现，不需要现在决定 |

**已交付、不要再当待办**（都在这两轮做完并发布）：B1–B4（账本更新 + 文档数字单一来源）、
A1（广告面重平衡 69→84）、A2（版式 × 占位符矩阵）、A6（`find_replace.selectFound` + `apply_style.range` 对齐）、
L1（71 个弱断言 → 0）、L2/L9（英文样式名 / `set_active_target`）、以及 FIXES 86 的用户报告全部条目。
详见 FIXES 86–91。

### 3.2 长期决策（不要再翻案）

- **D1 广告预算**：对外工具 ≤ **100 个**、schema ≤ **60,000 字节**（2026-09-30 第三次上抬，FIXES 68；当前用量 **84 / 46,867** —— 2026-10-04 按真实调用数据重平衡后，FIXES 90）。
- **D2 Word 长尾**：全都要（不做减法）。
- **D3 PPT 收敛**：**删** 媒体 / SmartArt / 讲义 / 3D 族 / 美化族；**保留** 版式 / 主题 / 尺寸 / 母版 / 节。
- **D4 废弃名处理**：18 个旧工具名**保留一个周期**作为 dispatch 别名；12 个 builtin 直接删除。
- **环境边界**：仅 Windows + COM；最低 **WPS 12.1 x64**；文档中文。（「仅 GitHub 发布」已由 2026-10-02 的发行决策取代，见下）
- **S1 落地决策（新增，实测依据见 FIXES 52）**：打开加密文件一律传**非空哨兵密码**
  （`$script:WpsNoPassword`，在桥头部）——空串等于「没给密码」，弹框照旧、会话照旧卡死。
- **S1 边界决策（新增）**：**超时后不自动关闭 WPS**。计划原本要求杀整个 WPS 进程树，实测
  「哪个进程是被卡住的那个」无法可靠指认（243 个 `wps`、Word/PPT 不给 `Hwnd`），照做会丢掉
  用户未保存的内容。改为讲清「状态未知」+ 短超时快速失败 + 宿主陈旧接管自愈。
- **S2 决策（新增）**：同一时间**只允许一个宿主**；第二个 DSH 会话得到中文错误而不是互相卡住。
- **`wps_execute_method`**：维持**隐藏**（P5-1 的最后手段契约不变），README「已知限制」已写清定位。

### v0.5.0 线决策（2026-09-30，用户拍板）

- **废弃名清理**：v0.5.0 里**清掉** 18 个废弃工具名与 `ALIAS_DEBT`（P1-4 目标 0），作为一次明确的
  breaking 变更处理，技能文档与 README 同步。这一条**推翻**上面 D4 的「保留一个周期」——窗口已过。
- **npm 发行（已被 2026-10-02 取代）**：当时只**准备**发布链路，**暂不真正发布**（D18 = B）。
- **真机整轮回归**：**不作为发版前必跑项**；`scripts/run-tests.ps1` 保持「需要时手动跑」。
- **广告预算**：见更新后的 D1（100 / 60,000，FIXES 68）。

### 2026-10-02 发行决策（用户拍板，**推翻上一条的「暂不真正发布」**）

- **D18 由 B 改为 A**：**开始对外发布 npm 包**（registry.npmjs.org，公开包），GitHub 标签作为备用渠道，
  两条渠道内容一致。首发版本 **0.6.0**（D3-A）。
- **D2-A 打包载荷**：`files` 由整个 `mcp` 目录改为**显式子路径** `mcp/dist` + `mcp/scripts`。
  原因：`files` 一写目录名，`.npmignore` 完全失效（已用仿真包实测），于是 `mcp/node_modules`（jest /
  ts-node / typescript 等 329 个顶层目录）被打进包：**10,280 文件 / 16.08 MB tarball / 解开 76.64 MB**。
  改后 **287 文件 / 约 0.6 MB / 解开 3.03 MB**。**不动**其他内容（`mcp/src`、`mcp/package.json`、
  `scripts/`、`LICENSES/` 都照旧发）——这是 D2 的 A 选项，不是最小集。
- **D4-B 源与凭证**：`publishConfig.registry = https://registry.npmjs.org/` **入库**（这样本机 `.npmrc`
  指向 npmmirror 时也不会发错地方）；**凭证绝不入库**。本机 `.npmrc` 目前**没有任何 npm 凭证**，
  `npm whoami`（含 npmjs）都是 `ENEEDAUTH`。
- **D5-A 只维护文档检查单**：新增 `docs/release-checklist.md`（手动流程）。**没有** `prepublishOnly`、
  **没有**发布用 CI 工作流——以后要加，得重新拍板。
- **运行时依赖的落点**（实测，是本决策的依据）：`mcp/dist` 的 bare import 只有
  `@modelcontextprotocol/sdk` / `uuid` / `winston`（外加三个 node 内建），三者都由**根** `package.json`
  声明、且与 `mcp/node_modules` 里的版本一致；在临时目录实测「删掉包内 `mcp/node_modules` 后仍解析成功、
  `doctor` 仍打印 `DOCTOR OK`」。所以删掉那 73.6 MB 不影响运行。
- **D6-A：`os` / `cpu` 挪到 package.json 顶层**（2026-10-02 用户拍板）。原来写在 `engines` 里，npm 静默忽略
  （`os`/`cpu` 只有顶层才生效）。挪动后两边都实测过：
  ① 非 Windows 平台会被 npm 拦下（这正是目的）；
  ② 本机 `dsh plugin add <tarball>` **照常装成功**（pnpm 1.1s、exit 0、`dump-config` 两个 id 都在、
  装出来的副本 287 文件 / 3.03 MB、`doctor` 打印 `DOCTOR OK`）—— 即平台约束不会挡住正常安装路径；
  ③ 根 `package-lock.json` 的 `packages[""]` 已同步镜像这两个字段，保持 lockfile 与 package.json 一致。
- **发布状态：已上架（2026-10-02 18:26 北京时间 / 10:26 UTC）**。用户走 npm 的浏览器授权（路线 1）完成 OTP，
  `dsh-plugin-wps-office-next@0.6.0` 已在 registry.npmjs.org 上，`dist-tags.latest = 0.6.0`、maintainer `sueccku`、
  `fileCount = 287`、`unpackedSize = 3,177,016`、integrity `sha512-rKZKMg/…Z5cuKUyCiPPnQ==`
  —— **与本地 dry-run 打印的 shasum/integrity 完全一致**，发出去的就是核对过的那个产物。
- **发布后验收（已做）**：全新隔离 profile `wpsnpmver` + 空 npm 缓存，`dsh plugin add dsh-plugin-wps-office-next`
  从 registry 拉取 → 2.5 秒、exit 0 → profile manifest 记成 `"dsh-plugin-wps-office-next": "^0.6.0"` 且
  `dsh.profile.bundles` 自动包含 → `--dump-config` 两个 id 都在 → 装出来的副本 **287 文件 / 3.03 MB**、
  `os=win32 cpu=x64`、无 `mcp/node_modules` → 从副本跑 `doctor` 打印 **`DOCTOR OK`**。验收 profile 已删除。
- **pnpm 的一个提示（记录备查）**：新包触发 pnpm 的 `minimumReleaseAge` 保护，它自动往
  `pnpm-workspace.yaml` 的 `minimumReleaseAgeExclude` 加了一行 `dsh-plugin-wps-office-next@0.6.0`；
  安装照常成功，不是错误。
- **后来都做完了**：`desktop` profile 已由用户自行升级到 `0.6.1` 并重启 DSH；真机调用也已通过**真实 MCP 工具**
  验证（见下一条）。

---

### 3.5 决策登记表（**跨会话的单一账本**，D12-A）

> **为什么要这张表**：决策散落在 CHANGELOG / FIXES / release-checklist / 本文件各处，很值钱但捞不出来。
> 以后**新决策写进这里**，并注明它在别处落在哪个版本/FIXES 号上。
>
> **编号消歧**：历史发布用过 D1–D18（0.6.0 那次），**本轮又用了 D5–D12** —— 所以同一个号在不同轮次含义不同；
> 引用时一律写「**哪一轮**的哪个号」或带上 FIXES/版本号，别只写「D5」。

> **读表须知**：这是**时间线**，早期行写的是当时的状态（例如 D9 那行写着「挂起」，但它在 FIXES 89 已做完）。
> **当前还有什么没定，一律看 §3.1**，不要从这张表里读"待办"。

| 轮次 | # | 问题 | 结论 | 落地 |
| --- | --- | --- | --- | --- |
| 加固期 | D1 | 广告预算上限 | ≤ 100 工具 / ≤ 60,000 字节 | FIXES 68 |
| 加固期 | D2 | Word 长尾做不做减法 | 全都要 | P3 |
| 加固期 | D3 | PPT 收敛范围 | 删媒体/SmartArt/讲义/3D/美化族，留版式/主题/尺寸/母版/节 | P4 |
| 加固期 | D4 | 18 个废弃工具名 | 保留一个周期 —— **已被 v0.5.0 线决策推翻**（清掉，ALIAS_DEBT 归零） | FIXES 77 |
| 0.5.x | D7 | 参数命名策略 | A：公开名对齐桥键，59 处全改 | FIXES 77 |
| 0.5.x | D8 | `set_cell_format` 扁平参数 | A：补进 schema（11 个扁平属性） | FIXES 78 |
| 0.5.x | D11–D15 | 桥键统一（一个概念一个键） | A / B / A / A / B：语义分组、统一 `transition`、统一 `url`、只收真同义、旧键本轮直接删 | FIXES 79 |
| 0.5.x | D17 | 未工具化 action 账本 | A：删 4 个重复/死代码 + 补 2 个真缺口（7 → 2） | FIXES 80 |
| 0.6.0 | D2 / D3 / D4 / D5 | 打包载荷 / 首发版本 / 源与凭证 / 流程 | A / A / B / A：`files` 改显式子路径、首发 0.6.0、`publishConfig.registry` 入库且凭证不入库、只维护文档检查单（无 `prepublishOnly`、无发布 CI） | [release-checklist.md](release-checklist.md) 开头 |
| 0.6.0 | D18 | 对外发布 | 由 B（只准备不发布）**改为 A**（上架 npm；首发 0.6.0 坏、0.6.1 修好） | FIXES 84 |
| **本轮**<br>2026-10-03 | **D1** | 下一步做哪一件 | **A 工程/文档收尾**（B 进 CI、C 广告面重平衡、D 弱覆盖补测、E 产品扩张 均未启动） | 提交 `471b48d` |
| **本轮** | **D5** | 这 3 个文档改动怎么处理 | **A**：只提交明确路径并推 `origin/main` | `471b48d` |
| **本轮** | **D6** | `verify-package` 是否进 CI | **B**：加 `npm run verify:package` + CI gate（实测 12.4 秒、11/11、不需要 WPS）——**接进去的第一分钟就抓到 FIXES 85** | CI `verify:package` gate |
| **本轮** | **D13** | `npm run` 下脚本必挂怎么办 | **A**（自动，无需另拍板）：既然是 D6-B 的门禁自身不可用，就修到可用——剥掉 npm 注入的环境再跑子 npm；直接调用 / `npm run` / 污染环境三种跑法全绿 | FIXES 85 |
| **本轮** | **D7** | 检查单的包体积期望值 | **A**：当场重测（`290 / 290 / 580KB / 3.04MB`）+ 注明会漂移 | release-checklist §2 |
| **本轮** | **D8** | `param-contract` 长时挂起 | **A**：先定性 —— 重跑正常（**秒级完成、257 对、A/B/C/D 全 0**），**未能复现**；挂起原因未定 | 见「挂起」栏 |
| **本轮** | **D9** | 75 个工具只有最弱断言（`any`） | **挂起**：先记着不做，**下一轮由我提醒**；再做时在 A（按域补真场景测试）/ B（只补最易错 5–8 个）之间选 | 未开工 |
| **本轮** | **D10** | 历史阶段数字（816/842/856） | **B**：保留原样，只加「当时的读数」标注（不改历史） | `docs/PROGRESS.md` 10 处 |
| **本轮** | **D11** | 与 DSH 自带离线 office 技能的边界 | **A**：只保留文档提醒，**不**改技能提示词（FIXES 82 已有共存提醒） | 无需改动 |
| **本轮** | **D12** | 决策记录放哪 | **A**：写进本文件这一节（`docs/` 入库、跨会话可见） | 本节 |
| **本轮** | **D14** | `replace_range` 越界（**数据破坏**） | **A**：越界**报错**（不静默钳制）+ 删超过 1 个段落标记需显式 `confirm` | FIXES 86 |
| **本轮** | **D15** | Word 没有「按范围设字体」 | **A 且不留兼容**：`range` 只收 `{start,end}` 或 `"all"`，旧的 `"selection"` 字符串彻底移除 | FIXES 86 |
| **本轮** | **D16** | 「假成功」怎么治 | **A**：格式类动作**写入后读回**，对不上/空范围走 warnings，不改 success 语义 | FIXES 86 |
| **本轮** | **D17** | `insert_text` 的 style 与 `new_paragraph` | **A**：style 落到刚插入的范围；`new_paragraph` 进键表并用真段落标记 | FIXES 86 |
| **本轮** | **D18** | 诊断与自验 | **C**：`affectedText` 改前置快照 + 回报受影响段落 + `get_paragraphs` 给字符坐标 | FIXES 86 |
| **本轮** | **D19** | 回归测试与弱断言 | **B**：新增 31 项回归；报告点名的 4 条从 `any` 升到 `ok` | FIXES 86 |
| **本轮** | **D20** | 交付节奏 | **B**：**攒批**，不单独发 0.6.2（数据破坏已在野，此选择是知情的） | 待发版 |
| **本轮** | **D21** | D9 挂起项 | **A**：**转正**——`any` 弱断言放行过真实的假成功，优先按域推进 | **已完成**：`any` 71 → 0（FIXES 89） |
| **本轮**<br>2026-10-04 | **D22** | 遗留清单里先做哪些 | **用户点名 A1 + A6 + A2 + B1**（C 类"看证据+说明"、A7 发版未表态，未动） | FIXES 90 |
| **本轮** | **D23** | B 类收尾的批 | **用户点名 B1–B4 全做**（B1 已在 FIXES 90 完成，本轮补 B2/B3/B4） | FIXES 91 |
| **本轮** | **A1（原 D1-C）** | 广告面重平衡 | **已做**：按真实调用数据把 15 个高频隐藏工具提到广告面，**69 → 84**（46,867 字节 / 预算 60,000）；**没换出任何工具**（零调用是语料偏置，不足为据） | FIXES 90 |
| **本轮** | **A6** | `find_replace` 不移动选区 / `range:"all"` 只有 `set_font` 收 | **已做**：新增 `selectFound`（默认 false，不动用户光标）；`apply_style.range` 与 `set_font` 对齐为 `{start,end}` 或 `"all"`，并让参数形状校验支持类型数组 | FIXES 90 |
| **本轮** | **A2** | PPT 版式 × 占位符没测清楚 | **已做**：真机扫出 6 个版式的占位符矩阵，据此修掉 3 个真缺陷（副标题判类型 2 应为 4、正文判 7 应为 2、`add_slide` 与 `set_slide_layout` 键名两套） | FIXES 90 |
| **本轮** | **B1** | `known-defects.md` 状态过期 | **已做**：S1 → fixed；C7 残余精确到"目的地解析"，并写明不修的理由 | FIXES 90 |
| **本轮**<br>2026-10-04 | **B2** | FIXES 验证表是旧快照 | **已做**：标注为"FIXES 63 快照 + 现值看哪里"，表头改"项数（当时）"（不毁历史证据） | FIXES 91 |
| **本轮** | **B3** | `PROGRESS.md` 阶段数字是历史值 | **已做**：最终数字一节标注为 v0.4.0 快照并指向现值；其余 26 处保持原样（D10-B） | FIXES 91 |
| **本轮** | **B4** | 文档数字没有单一来源（漂过 5 次） | **已做**：`gen-numbers.mjs` 从**真实 tools/list** + spec + 宿主键表算数、写 `docs/current-numbers.md`、逐项核对 README/HANDOFF 声明值；CI 加 `--check` + 生成物对账；`run-tests.ps1` 出 `test/summary.json` | FIXES 91 |
| **遗留轮**<br>2026-10-03 | **L1** | 71 个工具只有最弱断言（`any`） | **已做（用户选 A：按域全补）**：71 → **0**；过程中挖出并修掉 **12 个真缺陷**（Excel 2 / Word 1 / PPT 9），全是"回 success 但结果不对" | FIXES 89 |
| **遗留轮** | **L2** | 英文样式名不可用 | **已修**：工具层翻译（Heading 1→标题 1 等）+ 空格归一化 + 失败文案带下一步 | FIXES 88 + 28 项纯函数单测 |
| **遗留轮** | **L9** | `set_active_target` 校验失败仍回成功 | **已修**：两处失败改 `success: false`（锁照记，但不再谎报「校验通过」） | FIXES 88 + 真机回归 3 项 |

**挂起项（明确不做、但要记得）**

| # | 内容 | 状态 / 唤醒条件 |
| --- | --- | --- |
| D9（本轮） | 96 个工具没有专门测试、其中 **75 个只断言「没挂住」**（[tool-coverage.md](tool-coverage.md)：PPT 77 里 32、Excel 118 里 23、Word 59 里 15） | **已转正（D21-A）**：用户报告证明 `any` 放行过真实的假成功（`wps_word_replace_range` 越界删数据）。**已做完**：L1 把 71 个 `any` 全部补成真场景断言（`any` = 0，FIXES 89） |
| D8 残余 | `param-contract` 那次 2 小时挂起**未复现、未定性**（只读脚本，正常几秒） | 再出现一次时抓现场（别用 `Select-Object -Last N` 包住输出，会看不到进度）。建议顺手给它加心跳输出 + 自超时 |
| ~~D1-C~~ | 广告面重平衡 | **已做**（FIXES 90）：按 §8 的真实调用数据把 15 个高频隐藏工具提到广告面，69 → 84；**没有换出任何工具** |
| ~~D1-D~~ | 弱覆盖补测 | **已做**（FIXES 89）：71 个最弱断言 → 0，过程中修掉 12 个"回 success 但结果不对"的缺陷 |
| D1-E（本轮） | 产品扩张 / 跨应用工作流 | **仍挂起**：缺明确的目标客户与场景；没有场景就做等于猜需求。唤醒条件：用户给出具体场景 |

## 4. 架构与契约管线（顺序不能错）

三层，桥是唯一真源：

1. `mcp/scripts/wps-com.ps1` —— **桥，真源**（263 个动作分派；纯 CRLF、**无 BOM**）。
2. `host/wps-actions.ps1` —— **生成物**（`scripts/build-host-actions.ps1`），**UTF-8 BOM**，
   目标是与重新生成的结果字节一致。
3. `mcp/src/` —— TS 工具面（MCP server）。

**生成顺序（颠倒会拿到旧数据；完整顺序如下，一次都不能省）**：

```
cd mcp; npx tsc        # 1) 先把 TS 编进 mcp/dist（extract 读的是这里）
node scripts/extract-spec.mjs      # 2) 写 mcp/src/spec/operations.ts
cd mcp; npx tsc        # 3) 关键：把新 spec 再编进 mcp/dist
node scripts/gen-tool-surface.mjs  # 4) 读 mcp/dist/spec → spec/*.json
node scripts/gen-skill-tools.mjs   # 5) 技能清单
node scripts/gen-tool-coverage.mjs # 6) docs/tool-coverage.md
node scripts/gen-numbers.mjs       # 7) docs/current-numbers.md（并核对文档里的数字）
```

**为什么中间那次 tsc 不能省**：`extract-spec.mjs` 写的是 **`mcp/src/spec/operations.ts`**，而
`gen-tool-surface.mjs` / `gen-tool-coverage.mjs` 读的是 **`mcp/dist/spec/operations.js`**。少了第 3 步，
生成器看到的是**上一次编译的旧 spec**——实测症状是「明明 extract 说某个工具的 action 已修好，
surface 出来的 `action` 还是 `null`」，而且门禁只在 `spec-reproduction` 的 untooled 计数上叫（FIXES 90/91 各踩一次）。

**加一个新工具**：写 TS 定义 + handler（显式字面量 `executeMethod('action', {...})`，**绝不要工厂函数**）→ 按上面 7 步重跑管线。
**动态动作**必须在 `mcp/src/spec/aliases.ts` 的 `dynamicParamActions` 里声明，**否则宿主生成器直接 throw**。
守卫不变量：0 个未解析参数；桥的每个参数都必须映射到桥真正读取的 key。

> ⚠️ **两个静态抽取陷阱**（都真踩过，症状都是"门禁不叫但功能没了"）：
> 1. handler 的 `executeMethod<泛型>` 里**泛型部分不能超过 300 字符**（抽取器的窗口），超了会静默失配，
>    整个工具从 `bridge` 降级成 `opaque`、不进广告面（FIXES 90 踩过）。
> 2. **别在注释里写出抽取用的那个正则字面量**（`executeMethod…'action'`），解析器会把它当成一次真实调用，
>    于是 `calls.length` 变 2、该工具的**参数契约整条降级为 UNPARSED**（FIXES 91 踩过）。

**S1 新增的桥内助手**（都在 `wps-com.ps1` 头部，生成物里会一并出现，不要手改生成物）：
`Set-WpsAlertsSuppressed` / `Restore-WpsAlerts` / `Open-WordDocument` / `Open-ExcelWorkbook` /
`Open-PptPresentation` / `Format-WpsOpenError`。三条打开路径（`openWorkbook` / `openDocument` /
`openPresentation`）已经统一走它们（第四条 `openFile` 已在 FIXES 80 删除）；**再加打开路径时必须复用，不要直接调 `.Open()`**。

---

## 5. 当前权威数字（**引自代码内常量与生成物，不要用正则重数**）

> **本表的机器可读版**：[current-numbers.md](current-numbers.md) —— 由 `node scripts/gen-numbers.mjs` 生成，
> 并逐项核对本文件与 README 里的声明值（CI 跑 `--check`，对不上就红，FIXES 91）。
> 改了工具面/测试就先跑生成器，别手改数字。

| 指标 | 值 | 权威来源 |
| --- | --- | --- |
| 注册动作 | **263** | `scripts/verify.mjs` 的 `EXPECTED_ACTIONS` |
| 对外工具 | **84**（80 curated + 4 facade） | `mcp/src/server/toolset.ts` + `spec/advertised.json` |
| 广告面字节 | **46,867** / 上限 60,000 | `node scripts/gen-numbers.mjs`（真实 tools/list 载荷） |
| 全量 schema | 157,854 字节（268 工具，≈45,101 tokens） | 同上 |
| 预算 | `{ maxTools: 100, maxSchemaBytes: 60000 }` | `scripts/verify.mjs` |
| 测试 | **1077 断言 / 48 个测试文件**（口径见下方注；FIXES 89 +90、FIXES 90 +4、FIXES 91 发版审计 +3） | `test/*.test.mjs`（S3–S9 后 595 → 816，P2 +19，FIXES 65/66 +7，P3 +14，FIXES 80–83 +8；静态点名 `check(` 共 922 处，差额来自循环内断言） |
| e2e | 29 项检查，约 2–4 分钟（含归属记录一项） | `scripts/e2e.mjs` |
| 账本 | `ALIAS_DEBT = 0`、`UNTOOLED_ACTIONS = 2`（只剩 `getActivePresentation` / `getActiveWorkbook`，故意留着） | `test/spec-reproduction.test.mjs` |
| 参数契约 | **257** 对（A/B/C/D 四类均为 0，未解析 5） | `scripts/param-contract.mjs` |
| FIXES | 1～91 号 | `docs/FIXES.md` |

按能力域：Excel 118 / Word 59 / PPT **77** / 通用 14（含 4 个门面 + 转换）＝ 注册 268。

> **「断言数」的口径说明（2026-10-03 补）**：928 是 FIXES 81 那轮记录的**运行时合计**（46 个文件跑完各自打印的总数），
> 文档里历史上的 918 是它之前的读数。静态点名只能得到 782 —— 差额来自**循环里重复执行的 `check()`**，所以
> 「按名字数」永远数不出这个数。**下次整轮回归（`scripts/run-tests.ps1`）后，用各文件末尾打印的实际数字替换本行**，
> 不要让这个数字继续靠推算维持。

---

## 6. 已完成的阶段（历史；现状看 §1/§5）

P0 清理 → P1 规格真源 → P2 Excel 做深（5 波）→ P3 Word 做深（4 波）→ P4 PPT 收敛（2 波，88→76）
→ P5 收尾（逃生舱决策、广告面重定、文档重生成）→ **P5-4 发布（v0.2.0，随后 v0.2.1）**
→ **加固第 1 波（S1 + S2，FIXES 52）**。
详见 `docs/PROGRESS.md` 与 `docs/tool-roadmap.md`。

---

## 7. 稳定性实测（第 2 波计划的立足点；数字是当时读数，现值看 §5）

1. **模态弹窗的真实触发条件只有一个：打开加密文件**（实测）
   - 加密 `.docx`：裸开**永久卡死**（弹 `文档已加密` 模态框，class `Qt5QWindow`）；
     `PasswordDocument:=""` **无效**；**非空哨兵** → 4 秒内报错。
   - 加密 `.xlsx`：裸开卡死；哨兵 → 1 秒内 `0xFFF40006`。
   - 扩展名不符的老 `.doc`：**不弹**（18 秒内直接转换成功）——原计划的假设不成立。
   - `DisplayAlerts` 不是解（Word 侧本来就是 0，密码框照样弹）。
   - **已修**；残余的加密 **.pptx** 口子已由 **FIXES 69** 用文件头预检堵死（`Presentations.Open` 仍然没有密码参数，所以预检放在调用 Open 之前）。
2. **宿主并发**：已加命名互斥体 + 租约文件 + 陈旧接管，第二个会话得到可读中文错误（**已修**）。
3. **破坏性调用面**：`Delete()` **27 处**、`Clear()` 3、`ClearFormats()` 2、`ClearContents()` 2、
   `Unlist()` 1、`ResetAllPageBreaks()` 1（合计 35 站点 / 29 动作）。**S3 三批已落地**：**25 / 25 个用户数据动作**
   回传前置影响统计（范围类 5 + 对象类 8 + 批注/验证/Word/PPT 12）；清单见 `docs/destructive-operations.md`（FIXES 53 / 54 / 55）。
4. **测试覆盖缺口**（以下三个数字是 S4 当时的口径）：267 个工具里**只有 159 个被测试点名**（PPT 最弱，76 中仅 26）。
   历史上 **7 个「从来没工作过」的缺陷（FIXES 38 / 39 / 43，按正文逐条数共 7 处）全部落在无测试覆盖的路径上**。
   **S4 已完成**：覆盖率 **267/267**（当时；现在 **268/268**，ratchet 进 `spec-reproduction`，只许涨），`scripts/smoke-tools.mjs`
   出矩阵并可 `--live` 只读冒烟。**2026-10-04 收官（FIXES 89 + 90：D21-A / L1-A / D1-C）**：268 个工具里
**最弱断言（`any`）为 0**，广告面 **84 个**（按真实调用数据从 69 重平衡而来） —— 之前那 71 个 `any` 已按域全部补成真场景断言，
过程中挖出并修掉 **12 个真缺陷**（详见 FIXES 89，全是"回 success 但结果不对/字段是空的"）。
**这就是弱断言的代价**：只验"没挂住"会把"调用返回了"当成"事情做成了"。
5. **静默失败**：空 catch 走账本（**S5 已完成**，FIXES 59）：桥 34 + 宿主 6 全部登记，新增即红（`test/silent-catch.test.mjs`）。
6. **恢复路径原有的两个缺陷已修**（由本轮新测试抓出）：陈旧子进程的 `exit` 会反杀新宿主；
   `ready` 帧误清 `suspect` 标志。二者都在 `mcp/src/client/com-host.ts`。
7. **新的 WPS 静默失效（FIXES 52）**：`Document.Password` 按长度失效——15 字符卡死调用、
   **17 字符静默写出完全不加密的文件**。只影响 `wps_execute_method` 逃生舱。
8. **WPS 进程泄漏（FIXES 57，已修）**：`Get-WpsApp` 曾无条件执行 `New-Object -ComObject`，
   每次宿主获取实例都会启动并**丢弃**一个 WPS 实例；测试把它放大到 21 个 `et.exe` / 296 个
   `wps.exe` / 约 35 GB。已改为「先 `GetActiveObject`，拿不到才启动」，并记录自启实例、
   优雅关闭时按「无未保存内容」退出。这是**生产缺陷**，只是测试放大了它。

---

### 7.1 发布事故与教训（FIXES 84，2026-10-02）

**事故**：`0.6.0` 发到 npm 后，用户装进 `desktop` profile，**所有 WPS 工具全不可用**。MCP server 一启动就崩：

```
ReferenceError: exports is not defined in ES module scope
This file is being treated as an ES module because it has a '.js' file extension and
'.../node_modules/dsh-plugin-wps-office-next/package.json' contains "type": "module".
```

**根因**：`mcp/dist` 是 **CommonJS** 编译产物，`mcp/package.json`（现显式写着 `"type": "commonjs"`）是它的
模块边界。D2-A 把 `files` 从整个 `mcp` 目录改成 `mcp/dist` + `mcp/scripts` 时，把这个文件漏在了包外 ——
旧写法 `files: ["mcp"]` 是**顺带**把它带上的，改窄之后这个隐含依赖就断了。

**为什么没抓住（这条最值钱）**：发布前的验证只证明了三件事 —— 依赖能从顶层解析、从装出来的副本跑
`doctor.mjs` 打印 `DOCTOR OK`、「删掉包内 `mcp/node_modules` 也能解析」。**三件事全是真的**，但它们都不是
「server 真的能启动」。少的是一个**执行入口**上的断言；文件清单类的检查永远抓不到运行期缺件。

**修复与加固**

- `files` 加回 `mcp/package.json`（包内 288 文件）；`mcp/package.json` 显式写 `"type": "commonjs"`，
  让这份边界声明不再是隐含的。
- 新增 **`scripts/verify-package.mjs`**：`npm pack` → 装进临时目录 → **真的把 MCP server 拉起来**做一次
  JSON-RPC 握手 → 断言 `tools/list` 广告 **69** 个工具、`wps_status` 在列。**不需要 WPS**。
  这个脚本先在临时目录里**复现了 0.6.0 的崩溃**，再用同样的方式确认修复 —— 已写进
  `docs/release-checklist.md` §3.1 作为**发布前必跑**。
- 版本推进到 **0.6.1**；`CHANGELOG.md` 的 0.6.0 条目被打上「这一版是坏的」标记。
- 另加 **`scripts/probe-installed.mjs`**：对**已安装的副本**做同样的握手并真调一次 `wps_status`（需要 WPS），
  用于发布后与排错时分辨「包坏了」还是「环境没起 WPS」。

**0.6.1 的验收（2026-10-02 13:1x）** — *历史记录：其中"广告 69 个工具"是**当时**的数字，现值看 [current-numbers.md](current-numbers.md)（84）。*

- registry：`dist-tags.latest = 0.6.1`，integrity `sha512-jcztFCKPkhTpSPEPmnFcWNQeShNA707cy911D0uBoy26QROIEPoop2M7YdrL/0yOfPoimGbZUJ0v7fUued3T2A==`
  —— 与本地 dry-run 打印的**逐字符一致**；`fileCount = 289`、`unpackedSize = 3,187,511`。
- 全新隔离 profile 从 registry 装：4.2 秒、exit 0；装出来 **289 文件**、`mcp/package.json` 在、`type = commonjs`；
  `--dump-config` 两个 id 都在；握手 `serverInfo = {name: wps-office-mcp, version: 0.6.1}`、**广告 69 个工具**、
  `wps_status → connected: true`（真实 WPS 12.1.0.28488，latency 710ms）。验收 profile 已删除。
- **坏掉的 0.6.0 已撤销（2026-10-02 14:53 UTC，用户手动在交互终端完成）**。三次换 token 的尝试都失败，
  原因已实测并记进 `docs/release-checklist.md` §0：Granular **没勾** Bypass → `EOTP`；Granular **勾了** Bypass →
  npm 直接 `403`（*Granular access tokens that bypass two-factor authentication may not perform this action*）；
  且授权 URL 与发起命令的**进程**绑定，非交互终端里 npm 不等待、直接退出。
  **结论**：`unpublish` / `deprecate` 只能由人在交互终端跑。现在 registry 上**只剩 `0.6.1`**（`latest`）。
- **真实用户路径已验收（最终确认）**：`desktop` profile 已是 `0.6.1`（289 文件、`mcp/package.json` 在、
  `type = commonjs`），用户重启 DSH 后，**通过真实 MCP 工具调用** `wps_status` 返回
  `connected: true` / `advertisedTools: 69` / `registeredTools: 268` / `latencyMs: 695`（WPS 12.1.0.28488）。
  至此这条线闭环：包 → registry → 用户 profile → 模型真正调到工具。

**办法**：凡是改动「包内文件集合」的发布，验证必须落到**执行入口**（启动、握手、再跑一条真实调用），
只看文件清单与依赖解析不够。

---

## 8. 待办（下一步）

**当前没有必须做的技术债。** 未工具化 action 账本只剩 **2** 个（FIXES 80）：`getActivePresentation` / `getActiveWorkbook`，
两者的信息已由 `getOpenPresentations` / `getOpenWorkbooks` 的 `active` 标记覆盖，属对称性缺口（Word 的 `getActiveDocument` 是工具化的），
**故意留着**。全部弱断言（`any`）已归零（FIXES 89），四类参数静默失效均为 0。

### 8.1 开口的两项（与 §3.1 同源，别再重复列）

| # | 事项 | 状态 |
| --- | --- | --- |
| A3 | 产品扩张 / 跨应用工作流 | **等用户给具体场景**；没有场景不做 |
| A4 | `param-contract` 2 小时挂起 | 未复现；下次复现时抓现场（加心跳 + 自超时） |

### 8.2 发布状态（2026-10-04 实测）

- **在架版本：`0.6.1` + `0.6.2`，`dist-tags.latest = 0.6.2`**；坏掉的 `0.6.0` 已由用户 `npm unpublish` 撤销（2026-10-02），
  但 **CHANGELOG 保留那次事故的记录**（不要删）。
- 版本归属：FIXES 79→v0.5.3、80–83→v0.5.4、84→v0.6.1、**86–91→v0.6.2**。
- 「尚未 `npm publish`」「唯一在架版本是 0.6.1」这类句子都已过期，不要再当现状——发布事实一律看 §2。

计划之外、审计出来的可选工作（文档同步是其中 P0 项，本轮已做）：

| 优先级 | 内容 | 说明 |
| --- | --- | --- |
| ~~P0~~ | 同步陈旧文档 + 清 9 处 macOS 注释 | **本轮已做**：HANDOFF / stabilization-plan / `baseline/known-defects.md` / 工具层注释 |
| ~~P1~~ | 补 `docs/tool-coverage.md`；对齐版本号与 lockfile | **本轮已做**（生成器 `scripts/gen-tool-coverage.mjs` + CI 对账；`index.ts` serverInfo 改为读根 `package.json`，根/mcp lockfile 对齐 0.3.0） |
| ~~P2~~ | C7 目标漂移的统一前置；TS 层类型与枚举校验 | **本轮已做**（FIXES 64）：三处共用解析点加歧义警告（多文件且未指定目标才提示）+ 15 处内联解析回迁到共用解析点；TS 层只做**结构**校验，枚举仍由桥裁决。C7 残留（Word、第一方工具丢 warnings）记在 `known-defects.md` |
| ~~P3~~（可选） | 运行时 WPS 版本前置检查、fresh-profile 安装验收、e2e 扩充、lint 门禁 | **已做**（FIXES 67，随 v0.4.0 发布） |
| ~~P3~~ | 跨会话回收 WPS 孤儿实例 | **已做**（FIXES 66）：归属写盘 + 下一个宿主在租约下回收；有未保存内容不动 |

**发版**：S3–S9、审计 P2、进程残留治理（FIXES 65/66）与 P3 已作为 **v0.4.0（2026-09-19）** 打包发布；
**v0.5.0（2026-10-02，FIXES 68–76）** 同样走完整流程：升 `package.json` / lockfile、写 CHANGELOG、
更新 README 安装 pin（`#v0.5.0`）、tag + Release，并跑了一键 e2e **29/29**。**v0.5.1（2026-10-02）** 只做参数命名对齐、**v0.5.2** 补声明 `set_cell_format` 的扁平参数、**v0.5.3** 统一桥键（旧拼写不再接受）、**v0.5.4** 补两个缺口 + 修长动作超时名单 + DSH 0.2.0 兼容核验 + 文档审计，都走完整流程（pin 分别到 `#v0.5.1` … `#v0.5.4`）。

**真实用法数据（2026-10-03 采集，是「广告面该放哪些工具」的唯一实测依据；2026-10-04 已据此完成 A1 重平衡）**

汇总 `~/.dsh/sessions/**/session.v4.jsonl.zstd` 里所有真实 `wps_*` 调用（17 个会话、**742 次调用**，失败 6 次）：

- 调用量前两名是**门面**：`wps_call` **151**、`wps_help` **130** —— 说明模型大量在走「先查后调」；
- 被 `wps_call` 反复调用的**未广告**工具：`wps_excel_close_workbook` **36**、`wps_excel_export_chart_as_image` **16**、
  `wps_word_close_document` **15**、`wps_excel_get_conditional_formats` **15**、`wps_word_get_open_documents` **14**、
  `wps_excel_create_sheet` **14**、`set_conditional_format` / `set_sheet_header_footer` / `set_sheet_print_titles` 各 6–7；
- 模型 2 次**直接**调 `wps_excel_create_sheet` → `unknown tool`（注册表里有、广告面没有），只好退回 `wps_call`。

**读法（2026-10-04 更新）**：这段数据已经**用过一次**了 —— FIXES 90 / A1 按它把 15 个高频隐藏工具提到广告面
（69 → 84，「打开 → 干活 → 关闭」的尾巴补齐：`close_workbook` / `close_document` / `get_open_documents`）。
后续再要动广告面，请注意它同时影响预算与 `spec/advertised.json` 等生成物，改完必须重跑契约管线 + `gen-numbers`。

<details><summary>复算办法（可重复，不需要 WPS）</summary>

Node 26 的 `zlib.zstdDecompressSync` 可用；会话日志是**多帧 zstd 拼接**，必须按帧魔数 `28 B5 2F FD` 逐帧解，
直接对整文件解只拿到 190 字节（第一帧的会话头）。解出后逐行 JSON：`tool/call` 的 `data.name` 是工具名、
`data.arguments` 是 JSON 字符串；`tool/result` 的 `data.message.toolCallId` 与其配对，`isError` + `content[].text` 判失败。
`wps_call` 的真实目标在 `arguments.tool` 里。

</details>

**顺手可清**：无（`docs/PROGRESS.md` 的阶段数字按 D10-B 保留为历史快照，现值统一看 `docs/current-numbers.md`）。

**2026-10-03 的文档收尾（本轮，只动文档）**：把 **FIXES 84** 补进 `docs/FIXES.md`（此前只在本文与 `release-checklist.md` 里被引用，
修复日志本身缺号）；§2 的仓库/发布事实与 §8 的发布状态按命令实测重写（此前仍停在「0.6.0 已备好、尚未 publish」）；
§5 的账本、参数契约对数、FIXES 编号、断言口径一并校正；§9 补上「别在占用 WPS 的会话里跑全套」。
改动当时 `verify --static` 19 项、`lint` 0 违规、`gen-tool-coverage` 无漂移。

---

## 9. 环境与验证命令

- **本机 DSH 版本：`0.2.0-rc.2`**（文档旧值 `0.1.5-rc.1` 已过时）。兼容性已逐面核验 + 真机验收（FIXES 83）：
  `accept-install` 14/14、`e2e` 29/29；契约自查用本版新增的 `dsh --profile <name> --dump-config-schema`。
  升级 DSH 后重跑这两条即可复验；若失败，先看 `dsh plugin add` 的报错与 `--dump-config` 里两个接线 id 是否还在。
- 本机 DSH **是桌面版打包运行时**（FIXES 76）：`%LOCALAPPDATA%\Programs\DeepSeek Harness\` 下的
  `DeepSeek Harness.exe` + `resources\app.asar\dsh\...\dsh-desktop-host\lib\cli.js`（用 `ELECTRON_RUN_AS_NODE=1` 启动）；
  旧的 npm 全局安装 `%APPDATA%\npm\node_modules\@deepseek-ai\dsh\lib\bin.js` **已不存在**。
  `scripts/e2e.mjs` 两种装法都认（`resolveDshLauncher()`），找不到时用 `--dsh-bin <lib/bin.js>` 显式指定。
  `DSH_HOME = C:\Users\qwer\.dsh`。
- GUI：本会话通过 `http://127.0.0.1:19387` 交互（端口以实际启动为准，不要照抄旧值）。
  **现存 profile 只有 `desktop`（在用）与 `web`**；`wpsdoc2` / `wpse2e*` / `wpsnpm*` / `wps-acceptance` 都是历史临时 profile，
  每次用完都删掉了（若要再建，用完要 `dsh plugin --profile <name> remove dsh-plugin-wps-office-next`
  **（必须带包名，不带会报 `ERR_PNPM_MUST_REMOVE_SOMETHING`）**，再删 profile 目录。`node scripts/e2e.mjs --clean`
  只清产物、不动 profile。
- 会话日志：`~/.dsh/sessions/<编码的工作目录>/session-<id>/session.v<N>.jsonl.zstd`（**v4**，e2e 需要它做行为断言）。
- 真机 e2e（需要 WPS）：
  ```powershell
  node scripts/e2e.mjs --setup --timeout 420 --profile <name>
  node scripts/e2e.mjs --clean --profile <name>
  ```
- **宿主的单实例租约文件**：`%USERPROFILE%\.wps-office-mcp\com-host.json`
  （`hostPid` / `clientPid` / `心跳` / `phase` / `lastAction`）——售后排查「谁占着 WPS、卡在哪个动作」看它。
- 一次性 COM 探针放在 `test/.artifacts/`（**gitignored**，不会被提交）。
- **FIXES 87：`scripts/run-tests.ps1` 在**每个测试文件 spawn 之前**把 `WPS_OFFICE_MCP_ENTRY` / `WPS_OFFICE_HOST_SCRIPT` 指向仓库内路径**。
  从「已装本插件」的会话里起终端时，`plugin.js` 会把这两个变量指向 profile 里那份副本，测试于是测**上一个发布版**
  （症状：宿主沿用旧键表，报 `unknown parameter confirm` 之类，红得像是自己刚改的代码坏了）。
  **只删变量不够**：`plugin.test.mjs` 断言「`plugin.js` 发布的入口 == 包内入口」，副本路径会让它假红 —— 要"钉住"，不要"清理"。
  单独手跑某个 `test/*.test.mjs` 时，自己把这两个变量指向仓库内路径（`$env:WPS_OFFICE_HOST_SCRIPT = "$PWD\host\wps-com-host.ps1"`）。
- **改完 `scripts/*.ps1` 若行为"没变化"，先怀疑 PowerShell 的脚本缓存**（FIXES 87 实测踩过）：文件被 `git stash` 或直接覆写改过时，
  `-File` 启动可能仍读缓存里的旧内容 —— 判据是"内容长度一变行为就跟着变"。让文件内容长度发生变化（或改文件名）即可稳定绕过。
- **`test/confirm-dialog.test.mjs` 在本机会稳定红一条**（2026-10-03 实测）：它等 Excel 的确认框（`Qt*` 窗口），本机不弹，
  于是 `no modal confirmation dialog appeared` 失败。**已用 `git stash` 在改动之前复跑确认：同样红** —— 环境相关，不是某次改动引入的。
- **别在「正在用 WPS 的 DSH 会话」里跑全套**（2026-10-03 实测）：宿主的单实例租约正好把测试挡在门外 ——
  `test/host-lease.test.mjs`（10/18）与 `test/encrypted-preflight.test.mjs`（3 项）会打印**中文「另一个 DSH 会话正在控制 WPS」**而失败。
  这是**环境冲突，不是代码缺陷**（租约保护正常工作）。要拿到可信读数，就在没有活跃 WPS 宿主的终端里跑 `scripts/run-tests.ps1`，
  或先结束占用者（`%USERPROFILE%\.wps-office-mcp\com-host.json` 里的 `hostPid`）。

---

## 10. 纪律与已踩过的坑

**提交卫生（重要）**
- 每次开工先 `git status`；**提交只写明确路径，禁止 `git add -A`**。
- 原因：用户会在**另一条 DSH 会话**里同时改 `README.md` / `package.json`，已因此误扫过两次别人的未提交改动。

**方法**
- 先测量再改代码；每个结论都要有可运行命令；优先用**真机 WPS** 验证而不是推理。
  **教训**：计划里「补 `PasswordDocument:=""`」这条**实测是错的**（空串=没给密码，照样弹框），
  如果照计划直接改，会得到一个「看起来修好了、其实照样卡死」的版本。**计划也是假设，一样要验证。**
- **按危险度排序找缺陷：静默错 > 崩溃 > 挂起。** 这两轮挖出的 17 个缺陷里，**没有一个**是崩溃或超时——
  全是「回了 success 但结果不对」（字段名不匹配、占位符类型判错、参数被静默忽略）。
  所以：**断言要读回真值，不能只断言 `success`**。最弱的 `check(x.status === "ok")` 放行过真实的假成功；
  「`any`（只验没挂住）」这种断言等于没测（FIXES 86 的用户报告就是这么漏出去的）。
- **验收别人的产出要跑，不要读。** 发版前那次独立审计抓到的问题（README 指向不存在的 tag、生成器里写死的数字、
  我自己写的注释让静态解析器误判），**全都没有任何门禁会叫** —— 只有"真的去执行/核对"才能发现。
- **账本**（`ALIAS_DEBT` / `UNTOOLED_ACTIONS` / `EXPECTED_ACTIONS` / `BUDGET`）**涨即失败**，缩小要显式改。
- **杀掉后台作业 ≠ 杀掉它的子进程**（2026-10-03 实测，代价半小时）：`run-tests.ps1` 被中断后，它的主进程变成了**孤儿**，
  继续在循环里逐个启动 `test/*.test.mjs`；于是「刚跑通的回归测试」反复报**单实例租约被占**（写着别人的 `hostPid`/`clientPid`），
  看起来像自己的改动坏了。**排查任何"租约被占"，先查这两个**：`Get-CimInstance Win32_Process | ? { $_.CommandLine -like "*run-tests.ps1*" }`
  与 `… -like "*.test.mjs*"`（node）、`… -like "*wps-com-host.ps1*"`（宿主）。清理顺序：先杀 run-tests 主进程，再杀测试 node 与宿主，
  最后删 `%USERPROFILE%\.wps-office-mcp\com-host.json`。
- **别在会话里一边调 WPS 一边跑测试**：宿主是**单实例**的，你自己的探针/诊断调用会占住租约，把测试挡在门外（错误是中文的「另一个 DSH 会话…」）。
  跑测试前先释放自己的宿主（杀 `wps-com-host.ps1` + 删租约文件），跑完再继续。

**PowerShell / WPS 实测坑（都已踩过并修复，别再踩）**
- 生成的 `.ps1` 必须带 **UTF-8 BOM**，否则中文乱码。`host/wps-com-host.ps1` 本轮加了中文错误文案，
  **因此它现在必须保持 BOM**（`host/wps-actions.ps1` 由生成器写 BOM）。
- **不要用 `Get-Content -Raw` + `Set-Content -Encoding UTF8` 改含中文的 UTF-8 无 BOM 文件**
  （`.mjs` / 桥源码）：PS 5.1 会按 ANSI 读入、写成乱码，本轮把一个测试文件的 `密码|加密` 正则毁过一次。
  要么用编辑工具，要么 `[System.IO.File]::ReadAllText/WriteAllText` 并显式指定编码。
- `mcp/scripts/wps-com.ps1` 现在是**纯 CRLF、无 BOM**（7126 行、77 函数——以 `scripts/build-host-actions.ps1` 打印的 `functions` 为准）；改完要重新生成 host 并对账。
- PS 逗号优先级高于 `+`：`$m[$a + $r, $b + $c]` 会被解析错，必须加括号。
- `return $range` 会把多格 Range 展开成数组——要写 `return ,$range`。
- `Worksheet.Scenarios` 是 PSMethod（`$s.Scenarios()`）。
- ListObject 的几何信息在 `Delete()` / `Add()` 之后是**陈旧的**，必须先重新解析。
- 脚注正文要读 `.Reference.Text`，不是 `.Range.Text`；加完脚注后光标还在注释故事里，需要 `Get-MainTextRange` 兜底。
- **`$pid` 是只读变量**，探针里别拿它当局部变量名。
- **WPS 的 IDispatch 不支持命名参数**（`InvokeMember` + namedParameters → `E_INVALIDARG`），只能按位置传。
- **一次性临时文件（如测试 fixture）必须自己验证**：本轮「加密」fixture 一开始根本没加密，
  若不加验证，整份测试会为了错误的原因通过。
- **裸 COM 探针要按桥的规矩来，否则会把「共享的」WPS 实例搞坏，连累后面不相干的测试**（本轮真实翻车）：
  ① `GetActiveObject` 可能回一个**空壳实例**——`Workbooks` 看着正常，`Add()` 出来的工作簿 `Sheets` 是 null
  （桥里 `Test-WpsAppUsable` 防的就是这个）；② 对「脏」文档直接 `Close()` 会弹**模态保存框**，
  模态框会把整个实例钉住，之后所有调用回 `RPC_E_CALL_REJECTED`。
  规矩：取实例要「`GetActiveObject` → 校验 → `New-Object` → 校验」，open/close 一律用 `DisplayAlerts` 包住并还原。
- **WPS 卡住了怎么救（本轮实测有效，按顺序试）**：
  1) 列出可见窗口，只关 **`class` 以 `Qt*` 开头**的那些（那是对话框：密码框 / 保存框 / 恢复提示）——
     `XLMAIN` / `OpusApp` / `PP12FrameClass` 才是真正的文档窗口，**不要关**；
  2) 关了对话框后 COM 通常立刻恢复，这时再把残留文档/工作簿关掉（`Workbooks.Count=0`）；
  3) 仍不行才 `Stop-Process` **那一个**应用进程（`et` / `wps` / `wpp`），桥下次会自己重建实例。
  **不要**按进程名批量杀：实测本机有 **243 个**进程叫 `wps`、34 个叫 `et`，那是 WPS 的多进程架构，
  按名字杀会连带关掉用户所有文档。
- 跑全套前建议先看 `test/.artifacts/s1/visible.ps1`（一次性探针，未入库，丢了就照上面重写）——
  环境脏会让 `open-safety` 之外的测试也成片失败，别把它误判成代码缺陷。
- **清残留只清「从未保存过」的文档/工作簿**（`Path` 为空）：那是测试僵尸的特征；
  用户真正打开着的文件一律别碰——体检/清理自己变成丢数据的凶手，比不清理更糟。

**实测不可实现（是 COM 层面的限制，禁止重试，推广材料应写「不支持」）**
- Word 水印（页眉 `Shapes` 拒绝一切添加，`Count` 恒为 0）
- 文档属性（`BuiltInDocumentProperties` / `CustomDocumentProperties` 是空壳）
- Excel 切片器（`SlicerCaches.Add2` 可调用但 `Slicers.Count` 恒为 0）
- 方案管理器（`Scenarios` 行为与预期不符）
- **加密 .pptx 的密码框无法拦截**（`Presentations.Open` 没有密码参数）

---

## 11. 文件地图

| 路径 | 作用 |
| --- | --- |
| `mcp/scripts/wps-com.ps1` | **桥，真源**（263 个动作分派、77 函数、7126 行、纯 CRLF 无 BOM） |
| `host/wps-actions.ps1` | 生成物（字节一致证明目标，UTF-8 BOM） |
| `host/wps-com-host.ps1` | 常驻 STA 宿主：**单实例租约 + 心跳 + 陈旧接管**（手写，非生成，必须有 BOM） |
| `mcp/src/client/com-host.ts` | 宿主客户端：`timeoutFor` + suspect 短超时 + 陈旧子进程守卫 + 等旧宿主退出 |
| `mcp/src/client/wps-client.ts` | WPS 客户端门面（`executeMethod` / `invokeAction`） |
| `mcp/src/server/toolset.ts` | `STANDARD_TOOLS`（65 curated） |
| `mcp/src/server/mcp-server.ts` | 注册与描述（含 `wps_execute_method`、`wps_batch` 上限 50） |
| `mcp/src/tools/{excel,word,ppt}/` | 各能力域的 TS 工具定义 |
| `mcp/src/spec/aliases.ts` | `dynamicParamActions`——动态动作必须登记 |
| `scripts/{extract-spec,gen-tool-surface,gen-skill-tools}.mjs` | 契约管线（顺序见 §4） |
| `scripts/{verify,doctor,param-contract,e2e}.mjs` | 验证入口 |
| `scripts/build-host-actions.ps1` | 生成宿主动作表，打印 `switch_cases` / `functions` / `guard_installed` |
| `scripts/run-tests.ps1` | **整轮测试入口**：跑完每个文件回收无头 WPS 孤儿（FIXES 65），`-KeepOrphans` / `-Filter` 可调 |
| `scripts/lint.mjs` | 项目化 lint：手写 PowerShell 的 BOM/CRLF、制表符与行尾空白、`console.*`、测试退出码（进 CI） |
| `scripts/lint-alerts.mjs` | **弹框守卫门禁**（FIXES 70）：可能弹框的调用必须关 `DisplayAlerts` 并还原，且不许绕过两个共用助手 |
| `scripts/lint-com-boundary.mjs` | **COM 边界门禁**（FIXES 74）：`return $range` 必须 `return ,$range`；裸 `catch { continue }` 必须登记理由 |
| `scripts/lib/coverage-tiers.mjs` | **覆盖率口径唯一实现**（FIXES 74）：bespoke / matrixOk / matrixAny / notDriven 四层，CLI 与 CI 共用 |
| `mcp/src/utils/tool-warnings.ts` | **warnings 通道**（FIXES 73）：AsyncLocalStorage 收集桥侧 warnings，`tool-registry` 追加到结果文本 |
| `test/deprecated.test.mjs` | FIXES 76：18 个废弃工具名已全部消失（**不需要 WPS**，已进 CI） |
| `test/silent-catch.test.mjs` | S5：空 `catch { }` 必须登记理由，账本式（桥 34 + 宿主 6） |
| `scripts/accept-install.mjs` | 全新一次性 profile 的安装验收：装一遍、验接线与产物、跑 doctor、再拆掉 |
| `mcp/scripts/wps-com.ps1` 的 `owned-apps.json` 记录 | FIXES 66 的归属记录（`~/.wps-office-mcp/owned-apps.json`），只由强杀留下 |
| `test/*.test.mjs` | 47 个文件、949 断言（口径见 §5）；账本在 `spec-reproduction.test.mjs`；FIXES 86 的回归在 `word-range-format.test.mjs` |
| `test/target-ambiguity.test.mjs` | P2/C7 目标歧义警告：多文件且未指定目标才有 warning（需要真实 WPS） |
| `test/arg-shape-guard.test.mjs` | P2 入参形状守卫：数组/对象错位被拒，标量放行（不需要 WPS，已进 CI） |
| `test/orphan-reclaim.test.mjs` | FIXES 66 跨会话回收：强杀后新宿主收孤儿，无归属记录则不动（需要真实 WPS） |
| `test/wps-version.test.mjs` | P3 版本前置检查的纯函数单测（不需要 WPS，已进 CI） |
| `test/host-lease.test.mjs` | S2 单实例租约（**不需要 WPS**，已进 CI） |
| `test/open-safety.test.mjs` | S1 打开加密/异常文件不得卡死；**开头有环境体检**（需要真实 WPS） |
| `test/watchdog.test.mjs` | S1 超时契约（**不需要 WPS**，已进 CI） |
| `test/encrypted-preflight.test.mjs` | FIXES 69：加密 OOXML 在打开前被文件头认出来（**不需要 WPS**，已进 CI） |
| `test/alerts-gate.test.mjs` | FIXES 70：弹窗守卫门禁自己的断言（**不需要 WPS**，已进 CI） |
| `test/honest-reporting.test.mjs` | FIXES 71：close 的 saved 如实、protect 读回、删除列整段、PPT 导出不改指源文件（需要真实 WPS） |
| `test/range-limits.test.mjs` | FIXES 72：大范围预算、长路径文案、一次插整段（需要真实 WPS） |
| `test/warning-channel.test.mjs` | FIXES 73：warnings 经第一方工具到模型、原样透传不重复（需要真实 WPS） |
| `docs/FIXES.md` | 1～91 号修复记录（**新 bug 继续追加编号**；84 = 0.6.0 发布事故、85 = verify-package 在 `npm run` 下必挂、86 = 用户报告的 Word 假成功/数据破坏、87 = 测试跑的是旧副本、88 = 英文样式名 / `set_active_target`、89 = 补弱断言时挖出的 12 个真缺陷、90 = 广告面重平衡 + 版式矩阵 + 语义对齐、91 = 文档数字收成单一来源） |
| `docs/error-contract.md` | **错误与超时契约**：结果信封、批量部分失败、三档超时、调用方该做什么 |
| `docs/PROGRESS.md` / `tool-roadmap.md` | 阶段进展 / 路线图 |
| `docs/param-contract.md` | 生成物（重新生成后应无漂移） |
| `docs/stabilization-plan.md` | **加固计划**：S1–S9 全部完成（FIXES 52–63），含实测修正与排期状态 |
| `docs/destructive-operations.md` | **S3 破坏性操作清单**：动作 / 影响 / 回传统计 / 弹窗抑制 / 测试 + 剩余项 |
| `docs/tool-coverage.md` | **工具覆盖矩阵**（生成物）：工具 × 应用 × action × 广告 × 测试证据 |
| `scripts/gen-tool-coverage.mjs` | 生成 `docs/tool-coverage.md`；CI 重新生成并对账 |
| `README.md` | **面向客户的唯一契约**：安装由 AI 照做（7 步），AI 需逐条实测 |
| `CHANGELOG.md` | 在 `files` 白名单内，随包发布 |

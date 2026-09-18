# 交接文档（HANDOFF）

> 用途：把当前工作、进展、现状与下一步整理成**自包含**的一页，让一个**全新对话**无需回看历史即可接手。
> 核实时间：2026-09-16（v0.3.0 发布后重核；本文所有数字均从仓库/命令实测，非记忆）。
> 本文是工作文档，不随包发布（`docs/` 不在 `package.json` 的 `files` 白名单内）。

---

## 0. 新对话怎么开始

新开对话后，第一句话建议原样发送：

> 读 @docs/HANDOFF.md 和 @docs/stabilization-plan.md，然后继续。

接手时**先做只读核对**，确认真实状态与本文一致，再决定动手：

```powershell
Set-Location "D:\dsh\a"
git log --oneline -6 ; git status --porcelain ; git tag --list
node scripts/verify.mjs          # 预算 + 桥动作数 + 契约（23 项）
node scripts/doctor.mjs          # 环境自检（末尾应打印 DOCTOR OK）
```

跑完整测试（28 个文件，多数驱动真实 WPS，约 12 分钟）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "test\.artifacts\run-tests.ps1"
```

---

## 1. 一句话现状

**功能面仍与 v0.2.1 完全一致；加固第 1 波（S1 模态弹窗围堵 + S2 宿主单实例）已作为 `v0.3.0`
于 2026-09-16 推送并发布（提交 `d7736e1`，GitHub CI 全绿）；本机完整回归 **804/0**、一键 e2e
**28/28** 均通过——现在「不卡死、不撞车」这条线已经守住，可以开始小范围推广。
第 2～4 波（S3～S9）尚未开工，见 §8。**

---

## 2. 仓库与发布事实

| 项 | 值 |
| --- | --- |
| 仓库根 | `D:\dsh\a` |
| 分支 / HEAD | `main` / `d7736e1` release: v0.3.0（稳定性加固第 1 波）——已推送，与 `origin/main` 一致 |
| 远程 | `https://github.com/sueccku/dsh-plugin-wps-office-next.git` |
| 提交身份 | `sueccku <18247499+sueccku@users.noreply.github.com>` |
| 标签 | `v0.2.0`、`v0.2.1`、`v0.3.0` |
| Releases | v0.3.0（Latest，2026-09-16）、v0.2.1、v0.2.0——三份正文均已修正为正常 UTF-8 |
| 包 | `dsh-plugin-wps-office-next@0.3.0`，依赖 `@modelcontextprotocol/sdk`、`uuid`、`winston` |
| 构建脚本 | 只有 `snapshot` / `verify` / `gen:skills`——**没有 `prepare`**（安装时不需要构建） |

**工作区**：干净（`docs/HANDOFF.md` 本身已随仓库跟踪）。临时 profile `wpsdoc2` / `wpse2e` 已删除。

---

## 3. 已定决策（不要再翻案）

- **D1 广告预算**：对外工具 ≤ **70 个**、schema ≤ **40,000 字节**（当前 69 / 37,573）。
- **D2 Word 长尾**：全都要（不做减法）。
- **D3 PPT 收敛**：**删** 媒体 / SmartArt / 讲义 / 3D 族 / 美化族；**保留** 版式 / 主题 / 尺寸 / 母版 / 节。
- **D4 废弃名处理**：18 个旧工具名**保留一个周期**作为 dispatch 别名；12 个 builtin 直接删除。
- **环境边界**：仅 Windows + COM；最低 **WPS 12.1 x64**；仅 GitHub 发布；文档中文。
- **S1 落地决策（新增，实测依据见 FIXES 52）**：打开加密文件一律传**非空哨兵密码**
  （`$script:WpsNoPassword`，在桥头部）——空串等于「没给密码」，弹框照旧、会话照旧卡死。
- **S1 边界决策（新增）**：**超时后不自动关闭 WPS**。计划原本要求杀整个 WPS 进程树，实测
  「哪个进程是被卡住的那个」无法可靠指认（243 个 `wps`、Word/PPT 不给 `Hwnd`），照做会丢掉
  用户未保存的内容。改为讲清「状态未知」+ 短超时快速失败 + 宿主陈旧接管自愈。
- **S2 决策（新增）**：同一时间**只允许一个宿主**；第二个 DSH 会话得到中文错误而不是互相卡住。
- **`wps_execute_method`**：维持**隐藏**（P5-1 的最后手段契约不变），README「已知限制」已写清定位。

---

## 4. 架构与契约管线（顺序不能错）

三层，桥是唯一真源：

1. `mcp/scripts/wps-com.ps1` —— **桥，真源**（267 个动作分派；纯 CRLF、**无 BOM**）。
2. `host/wps-actions.ps1` —— **生成物**（`scripts/build-host-actions.ps1`），**UTF-8 BOM**，
   目标是与重新生成的结果字节一致。
3. `mcp/src/` —— TS 工具面（MCP server）。

**生成顺序（颠倒会拿到旧数据）**：

```
scripts/extract-spec.mjs  →  tsc  →  scripts/gen-tool-surface.mjs  →  scripts/gen-skill-tools.mjs
```

注意：`gen-tool-surface.mjs` 读的是 **`mcp/dist`**，所以**必须先 tsc 再生成**。

**加一个新工具**：写 TS 定义 + handler（显式字面量 `executeMethod('action', {...})`，**绝不要工厂函数**）→ extract → tsc → gen。
**动态动作**必须在 `mcp/src/spec/aliases.ts` 的 `dynamicParamActions` 里声明，**否则宿主生成器直接 throw**。
守卫不变量：0 个未解析参数；桥的每个参数都必须映射到桥真正读取的 key。

**S1 新增的桥内助手**（都在 `wps-com.ps1` 头部，生成物里会一并出现，不要手改生成物）：
`Set-WpsAlertsSuppressed` / `Restore-WpsAlerts` / `Open-WordDocument` / `Open-ExcelWorkbook` /
`Open-PptPresentation` / `Format-WpsOpenError`。四条打开路径（`openFile` / `openWorkbook` /
`openDocument` / `openPresentation`）已经统一走它们；**再加打开路径时必须复用，不要直接调 `.Open()`**。

---

## 5. 当前权威数字（**引自代码内常量，不要用正则重数**）

| 指标 | 值 | 权威来源 |
| --- | --- | --- |
| 注册动作 | **267** | `scripts/verify.mjs` 的 `EXPECTED_ACTIONS` |
| 对外工具 | **69**（65 curated + 4 facade） | `mcp/src/server/toolset.ts` + `spec/advertised.json` |
| 广告面字节 | **37,573** / 上限 40,000 | `node scripts/verify.mjs` |
| 全量 schema | 153,777 字节 | 同上 |
| 预算 | `{ maxTools: 70, maxSchemaBytes: 40000 }` | `scripts/verify.mjs` |
| 测试 | **804 断言 / 35 个测试文件** | `test/*.test.mjs`（S3+S4+S5+S6+S7 后 595 → 804） |
| e2e | 28 项检查，约 94 秒 | `scripts/e2e.mjs` |
| 账本 | `ALIAS_DEBT = 59`、`UNTOOLED_ACTIONS = 7` | `test/spec-reproduction.test.mjs` |
| 参数契约 | 255 对（A/B/C/D 四类均为 0） | `scripts/param-contract.mjs` |
| FIXES | 1～61 号 | `docs/FIXES.md` |

按能力域：Excel 118 / Word 59 / PPT 76 / 通用 14。

---

## 6. 已完成的阶段

P0 清理 → P1 规格真源 → P2 Excel 做深（5 波）→ P3 Word 做深（4 波）→ P4 PPT 收敛（2 波，88→76）
→ P5 收尾（逃生舱决策、广告面重定、文档重生成）→ **P5-4 发布（v0.2.0，随后 v0.2.1）**
→ **加固第 1 波（S1 + S2，FIXES 52）**。
详见 `docs/PROGRESS.md` 与 `docs/tool-roadmap.md`。

---

## 7. 稳定性实测（**本轮新增证据**，是第 2 波计划的立足点）

1. **模态弹窗的真实触发条件只有一个：打开加密文件**（实测）
   - 加密 `.docx`：裸开**永久卡死**（弹 `文档已加密` 模态框，class `Qt5QWindow`）；
     `PasswordDocument:=""` **无效**；**非空哨兵** → 4 秒内报错。
   - 加密 `.xlsx`：裸开卡死；哨兵 → 1 秒内 `0xFFF40006`。
   - 扩展名不符的老 `.doc`：**不弹**（18 秒内直接转换成功）——原计划的假设不成立。
   - `DisplayAlerts` 不是解（Word 侧本来就是 0，密码框照样弹）。
   - **已修**；残余口子是加密 **.pptx**（`Presentations.Open` 无密码参数，无法阻止弹框），已写入 README。
2. **宿主并发**：已加命名互斥体 + 租约文件 + 陈旧接管，第二个会话得到可读中文错误（**已修**）。
3. **破坏性调用面**：`Delete()` **27 处**、`Clear()` 3、`ClearFormats()` 2、`ClearContents()` 2、
   `Unlist()` 1、`ResetAllPageBreaks()` 1（合计 35 站点 / 29 动作）。**S3 三批已落地**：**25 / 25 个用户数据动作**
   回传前置影响统计（范围类 5 + 对象类 8 + 批注/验证/Word/PPT 12）；清单见 `docs/destructive-operations.md`（FIXES 53 / 54 / 55）。
4. **测试覆盖缺口**：267 个工具里**只有 159 个被测试点名**（PPT 最弱，76 中仅 26）。
   历史上 **7 个「从来没工作过」的缺陷（FIXES 38/39/40/43/47/49）全部落在无测试覆盖的路径上**。
   **S4 已完成**：覆盖率 **267/267**（ratchet 进 `spec-reproduction`，只许涨），`scripts/smoke-tools.mjs`
   出矩阵并可 `--live` 只读冒烟；剩余是把 PPT（25/76）等未覆盖工具补上场景测试（FIXES 56）。
5. **静默失败**：桥里 242 个 `catch`，其中 **13 个是空的**（S5 待做）。
6. **恢复路径原有的两个缺陷已修**（由本轮新测试抓出）：陈旧子进程的 `exit` 会反杀新宿主；
   `ready` 帧误清 `suspect` 标志。二者都在 `mcp/src/client/com-host.ts`。
7. **新的 WPS 静默失效（FIXES 52）**：`Document.Password` 按长度失效——15 字符卡死调用、
   **17 字符静默写出完全不加密的文件**。只影响 `wps_execute_method` 逃生舱。
8. **WPS 进程泄漏（FIXES 57，已修）**：`Get-WpsApp` 曾无条件执行 `New-Object -ComObject`，
   每次宿主获取实例都会启动并**丢弃**一个 WPS 实例；测试把它放大到 21 个 `et.exe` / 296 个
   `wps.exe` / 约 35 GB。已改为「先 `GetActiveObject`，拿不到才启动」，并记录自启实例、
   优雅关闭时按「无未保存内容」退出。这是**生产缺陷**，只是测试放大了它。

---

## 8. 待办（下一步）

`docs/stabilization-plan.md` 共 9 项；**第 1 波（S1 + S2）已完成并随 v0.3.0 发布**，剩下：

| 波次 | 条目 | 目的 | 估工 |
| --- | --- | --- | --- |
| **2** | S3 破坏性操作加守卫/回传统计（**25/25 已落地**，仅剩确认框实测；见 `docs/destructive-operations.md`）· S4 逐工具冒烟矩阵 + 覆盖率账本 | 不丢数据、有兜底 | 3.5 日 |
| 3 | S5 13 处空 catch 建 allowlist · S6 失败/超时契约 · S7 中文错误文案 | 可解释、可归因 | 2 日 |
| 4（可选） | S8 `wps_execute_method` 定位 + WPS 版本前置检查 · S9 安装摩擦兜底 | 体验与售前 | 1 日 |

**建议下一轮直接开第 2 波**：S3 的验收物是「破坏性操作清单」表 + 每处「是」的断言；
S4 的大头是**真机跑数**（PPT 50 个未覆盖 → Excel 33 → Word 18），不是写代码。

**上一轮遗留的两个待拍板问题已经落地**：`wps_execute_method` 维持隐藏（README 已写清定位）；
S1 + S2 已开工并完成。所以现在**没有阻塞项**，可以按上表继续。

**顺手可清**：无（`docs/PROGRESS.md` 已同步到 804 项 / 35 文件）。

---

## 9. 环境与验证命令

- 本机 DSH：`0.1.5-rc.1`，位于 `C:\Users\qwer\AppData\Roaming\npm\node_modules\@deepseek-ai\dsh\`；`DSH_HOME = C:\Users\qwer\.dsh`。
- profile：`web`（GUI :3080，**活跃中，勿扰**）、`headless`（模板）。临时 profile `wpsdoc2` / `wpse2e`
  已在 2026-09-16 清理；下次真机 e2e 用 `--setup` 新建，跑完记得 `dsh plugin --profile <name> remove …` 再删目录。
- 真机 e2e（需要 WPS）：
  ```powershell
  node scripts/e2e.mjs --setup --timeout 420 --profile <name>
  node scripts/e2e.mjs --clean --profile <name>
  ```
- **宿主的单实例租约文件**：`%USERPROFILE%\.wps-office-mcp\com-host.json`
  （`hostPid` / `clientPid` / `心跳` / `phase` / `lastAction`）——售后排查「谁占着 WPS、卡在哪个动作」看它。
- 一次性 COM 探针放在 `test/.artifacts/`（**gitignored**，不会被提交）。

---

## 10. 纪律与已踩过的坑

**提交卫生（重要）**
- 每次开工先 `git status`；**提交只写明确路径，禁止 `git add -A`**。
- 原因：用户会在**另一条 DSH 会话**里同时改 `README.md` / `package.json`，已因此误扫过两次别人的未提交改动。

**方法**
- 先测量再改代码；每个结论都要有可运行命令；优先用**真机 WPS** 验证而不是推理。
  **本轮的最大教训**：计划里「补 `PasswordDocument:=""`」这条**实测是错的**（空串=没给密码，照样弹框），
  如果照计划直接改，会得到一个「看起来修好了、其实照样卡死」的版本。**计划也是假设，一样要验证。**
- 账本（`ALIAS_DEBT` / `UNTOOLED_ACTIONS` / `EXPECTED_ACTIONS` / `BUDGET`）**涨即失败**，缩小要显式改。

**PowerShell / WPS 实测坑（都已踩过并修复，别再踩）**
- 生成的 `.ps1` 必须带 **UTF-8 BOM**，否则中文乱码。`host/wps-com-host.ps1` 本轮加了中文错误文案，
  **因此它现在必须保持 BOM**（`host/wps-actions.ps1` 由生成器写 BOM）。
- **不要用 `Get-Content -Raw` + `Set-Content -Encoding UTF8` 改含中文的 UTF-8 无 BOM 文件**
  （`.mjs` / 桥源码）：PS 5.1 会按 ANSI 读入、写成乱码，本轮把一个测试文件的 `密码|加密` 正则毁过一次。
  要么用编辑工具，要么 `[System.IO.File]::ReadAllText/WriteAllText` 并显式指定编码。
- `mcp/scripts/wps-com.ps1` 现在是**纯 CRLF、无 BOM**（6542 行、60 函数）；改完要重新生成 host 并对账。
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
| `mcp/scripts/wps-com.ps1` | **桥，真源**（267 个动作分派、60 函数、6542 行、纯 CRLF 无 BOM） |
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
| `test/*.test.mjs` | 28 个文件、595 断言；账本在 `spec-reproduction.test.mjs` |
| `test/host-lease.test.mjs` | S2 单实例租约（**不需要 WPS**，已进 CI） |
| `test/open-safety.test.mjs` | S1 打开加密/异常文件不得卡死；**开头有环境体检**（需要真实 WPS） |
| `test/watchdog.test.mjs` | S1 超时契约（**不需要 WPS**，已进 CI） |
| `docs/FIXES.md` | 1～61 号修复记录（**新 bug 继续追加编号**） |
| `docs/error-contract.md` | **错误与超时契约**：结果信封、批量部分失败、三档超时、调用方该做什么 |
| `docs/PROGRESS.md` / `tool-roadmap.md` | 阶段进展 / 路线图 |
| `docs/param-contract.md` | 生成物（重新生成后应无漂移） |
| `docs/stabilization-plan.md` | **加固计划**：第 1 波已完成并标注实测修正，第 2～4 波待做 |
| `docs/destructive-operations.md` | **S3 破坏性操作清单**：动作 / 影响 / 回传统计 / 弹窗抑制 / 测试 + 剩余项 |
| `README.md` | **面向客户的唯一契约**：安装由 AI 照做（7 步），AI 需逐条实测 |
| `CHANGELOG.md` | 在 `files` 白名单内，随包发布 |

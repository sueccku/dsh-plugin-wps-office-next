# 更新日志

本文件记录每个发布版本的用户可见变化；逐条修复的原因与实测证据见 [docs/FIXES.md](docs/FIXES.md)。

## 0.6.2（假成功与数据破坏收口 · 广告面重平衡 · 文档数字进 CI 对账）

> **建议直接升到这一版。** 0.6.1 里已知有两类问题：**格式类工具谎报成功**、以及 **Word 替换越界会删数据**。

### 数据与结果正确性（来自用户报告，FIXES 86）

- **`wps_word_replace_range` 越界会静默删数据（严重）**：范围超出文档时以前按"文档能到哪就改到哪"处理，
  于是"替换第 3 段的第 5-80 个字符"这种请求会把**后面所有段落**吞掉。现在越界**直接报错**，并告诉你文档有多长、
  下一步该用什么坐标；删除超过 1 个段落标记需要显式 `confirm`。
- **格式类工具"报成功但没生效"**：按范围设字体、应用样式、设段落格式等改为**写入后读回**，对不上就出 warning；
  失败信息统一带"动作 / 下一步"，不再只说一句 COM 报错。
- `wps_word_insert_text`：`style` 现在作用在**刚插入的那段**上；`new_paragraph` 生成的是真段落标记。

### 假成功清理（FIXES 88 / 89 / 90）

- 英文样式名可用：`Heading 1` → `标题 1`、`Normal` → `正文` 等 15 个别名，失败时会列出可用的样式名。
- `set_active_target` 校验失败不再回 `success: true`。
- **12 个"回 success 但结果不对"的缺陷**（把 71 个只有最弱断言的工具补成真场景断言时挖出来的）：
  - Excel：`get_selection` 读错字段（行数显示 undefined）、`text_to_columns` 必抛错；
  - Word：`generate_toc` 参数多传 2 个 → **目录从来建不出来却回成功**；
  - PPT：`get_slide_master`/`get_table_cell` 字段名错位、`set_slide_title` 在空白版式上静默不写、
    `set_slide_layout` 只认英文键、图表数据与动画顺序、`set_font_color`、`beautify` 崩溃等。
- **封面/标题页的副标题与正文其实从来写不进去**：`set_slide_subtitle` 判的是占位符类型 2（实为 4）、
  `set_slide_content` 判 7（实为 2）；`add_slide` 与 `set_slide_layout` 还各用一套版式键名。已修，并附实测的
  "版式 × 占位符"对照表（写进 WPS 演示技能）。

### 工具面（默认档）

- **广告面按真实调用数据重平衡：69 → 84 个工具**（schema 38,878 → 46,867 字节，上限 60,000）。
  依据是本机 742 次真实工具调用：`close_workbook` 36 次、`get_open_documents` 19 次、`export_chart_as_image` 17 次…
  这些以前都要先 `wps_help` 查、再 `wps_call` 调，现在直接可见。**没有移除任何工具。**
- `wps_word_find_replace` 新增 `selectFound`：查找后把选区定位到命中处，可以接着对"刚找到的那处"施加格式；
  默认仍是**不动你的光标**。
- `wps_word_apply_style` 的 `range` 与 `set_font` 统一：`{start,end}` 或 `"all"`。

### 工程与文档

- 测试从 918 项涨到 **1,076 项 / 48 个文件**；新增回归覆盖用户报告的全部问题。
- **文档里的数字改为机器生成 + CI 对账**（`docs/current-numbers.md`）：README/HANDOFF 里的工具数、字节数
  与测试数从此对不上就会让 CI 变红（历史上这组数字漂过 5 次）。
- 修掉测试环境的一个真隐患：从"已装本插件"的会话里跑测试时，测的其实是**上一个发布版的副本**。

**破坏性变更**：无（工具名、参数含义、技能文档调用方式均未变；新增的都是可选参数）。

数字：注册工具 **268**、广告面 **84 / 46,867 字节**、桥 action **263**、测试 **1,076 项 / 48 文件**、
包内 **296 文件 / 611.9 KB（解包 3.13 MB）**。

## 0.6.1（修掉 0.6.0 的致命打包缺陷：`mcp/package.json` 漏发）

> **0.6.0 在 npm 上是坏的，请直接装 0.6.1。** 症状是 MCP server 起不来、所有 WPS 工具全不可用。

- **根因**：`mcp/dist` 是 **CommonJS** 编译产物，`mcp/` 目录里那份 `package.json`（`"type": "commonjs"` 的隐含来源）
  是它的**模块边界声明** —— 有了它，Node 才不把 `dist/*.js` 当成 ESM。0.6.0 把 `files` 从整个 `mcp` 目录改成
  `mcp/dist` + `mcp/scripts` 时，把这个文件漏在了包外，于是 `mcp/dist/index.js` 继承了根 `package.json` 的
  `"type": "module"`，一启动就抛：

  ```
  ReferenceError: exports is not defined in ES module scope
  This file is being treated as an ES module because it has a '.js' file extension and
  '.../node_modules/dsh-plugin-wps-office-next/package.json' contains "type": "module".
  ```

- **修复**：`files` 加回 `mcp/package.json`（288 个文件）。**验证方式也换了**：不再只看文件在不在，而是
  **真的把 MCP server 拉起来做一次 JSON-RPC 握手**并调用 `wps_status` —— 握手成功、广告 69 个工具、
  `connected: true`（真实 WPS 12.1.0.28488）才算过。这条已写进 `docs/release-checklist.md` §3。
- 这一版还包含 0.6.0 的全部内容：npm 上架、载荷从 76.64 MB 瘦到 3.03 MB、`publishConfig.registry`、
  `os`/`cpu` 挪到顶层。

**破坏性变更**：无（工具名、参数、行为、技能文档均未变）。

数字：注册工具 **268**、广告面 **69 / 38,878 字节**、桥 action **263**、包内 **288 文件 / 3.03 MB**。


## 0.6.0（发布到 npm：载荷从 76.6 MB 瘦到 3.0 MB）— **npm 上的这一版是坏的，请用 0.6.1**

> ⚠️ **这一版已发布但不可用**：MCP server 无法启动（`exports is not defined in ES module scope`），
> 所有 WPS 工具均不可调用。原因与修复见上面的 0.6.1。GitHub 标签 `v0.6.0` 同样是坏的。

> 功能面与 0.5.4 **完全一致**（注册工具 268、广告面 69、桥 action 263 都没动）。这一版只改分发方式。

- **上架 npm**：包名 `dsh-plugin-wps-office-next`，安装可以只写包名，不必再依赖 GitHub 可达性：

  ```powershell
  dsh plugin --profile <profile> add dsh-plugin-wps-office-next
  ```

- **打包载荷修正**：`files` 之前写的是整个 `mcp` 目录，于是 `mcp/node_modules`（jest / ts-node / typescript 这些
  开发依赖，共 329 个顶层目录）连同 `mcp/src` 一起被打进包里 —— tarball **16.08 MB**、**10,280 个文件**、解开
  **76.64 MB**。改成显式列出 `mcp/dist` 与 `mcp/scripts` 之后是 **287 个文件 / 3.03 MB**（解开），tarball 约 0.6 MB。
  运行不受影响：`mcp/dist` 的依赖（`@modelcontextprotocol/sdk` / `uuid` / `winston`）由根 `package.json` 声明，
  安装时照常解析（已在临时目录实测：删掉包内 `mcp/node_modules` 后仍能解析成功、`doctor` 仍打印 `DOCTOR OK`）。
- **`publishConfig.registry`**：固定为 `https://registry.npmjs.org/`，这样本机 `.npmrc` 指向镜像源时
  `npm publish` 也不会发错地方。仓库里不存任何凭证。
- 发布流程与检查单见 [docs/release-checklist.md](docs/release-checklist.md)。

**破坏性变更**：无（工具名、参数、行为、技能文档均未变）。

数字：注册工具 **268**、广告面 **69 / 38,878 字节**、桥 action **263**、测试 **918 项 / 46 文件**。

## 0.5.4（能结束放映、能设圆角 · 长动作超时修正 · DSH 0.2.0 兼容核验）

### 补上的两个能力（FIXES 80）

- **新工具 `wps_ppt_end_slide_show`**：以前能开始放映却没有工具能结束 —— 放映一旦开始就占满屏幕，只能让用户自己按 Esc。
- **`wps_ppt_set_shape_style` 增加 `roundness`**：圆角半径；写入后会**读回**，结果里回报实际生效值（读不回或不一致会告警）。
- 顺带删掉 4 个重复/死代码 action（`openFile` / `unfreezePanes` / `replaceInSheet` / `setShapeRoundness`），
  未工具化账本 **7 → 2**（剩下两个已被 `getOpenPresentations` / `getOpenWorkbooks` 的 active 标记覆盖）。

### 修掉一个静默降级（FIXES 81）

- `LONG_ACTIONS` 里有三个名字不是真实 action（`beautify` / `recalculate` / `proofreadBasic`）：写错不报错，
  只会让那个动作退回 **60s 默认超时** —— `calculateSheet` 这类慢操作会被杀成「状态未知」。已改成真实名字，
  并加静态门禁：**每个长动作名都必须是真实桥 action**（写错 CI 直接红）。

### DeepSeek Harness 0.2.0-rc.2 兼容性核验（FIXES 83）

- 逐面核验：包清单 `dsh.bundle.patch`、patch 操作 `insert`、MCP 客户端九个配置键、`toolCallTimeoutMs`（无上限截断，
  330s 有效）、技能注册 API、CLI 用法、`DSH_PERMISSION_MODE` / `DSH_TOOLS_MODE`、会话日志 v4；**无需改接线**。
- 真机验收：全新 profile 安装 **14/14**、一键 e2e **29/29**（83 次工具调用 / 3 次技能加载）。
- README 注明与 DSH 自带离线 office 技能（LibreOffice）的共存关系：一个改你眼前的 WPS 窗口，一个是离线读写文件。

### 文档（FIXES 82）

- 33 份文档全量审计：修掉过期数字、已删除的死名字（模型面向的技能示例也在内）、失效链接与不存在的文件引用；
  机器扫描剩余项全部是有标注的历史引用。

**破坏性变更**：4 个 action 被删除（均为重复品或死代码，仓库内无测试与脚本调用）。
数字：注册工具 **268**、广告面 **69 / 38,878 字节**、桥 action **263**、测试 **918 项 / 46 文件**。
## 0.5.3（桥键统一：一个概念一个键 + set_cell_format 的 numberFormat）

> 破坏性变更两处：① 参数键统一，**旧拼写不再被接受**（`filePath` / `dataRange` / `address` / `transition` 等）；
> ② `wps_common_save_as` 的路径键由 `path` 改为 `outputPath`。

- **一个概念一个键**（FIXES 79）：输入文件 `path`、输出目标 `outputPath`、图片 `imagePath`、幻灯片切换
  `transition`、链接 `url`、区域 `range` —— 三个应用一致，可以直接类推。桥侧每个 action 只认自己的规范键，
  兼容别名按 D15-B 删除；旧写法会明确报错，不再静默忽略。
- `set_cell_format` 的扁平 `numberFormat` 一并声明（D10，FIXES 78 的收尾）。
- 顺带修掉三个导出工具重复发 `path` 造成的静默失效（由 `param-contract` 抓出）。
- 数字：测试 **914 项 / 46 文件**；注册工具 267、广告面 **69 / 38,878 字节**（内部上限 100 / 60,000）。
## 0.5.2（模型能发现 set_cell_format 的扁平参数）

- `wps_excel_set_cell_format` 的扁平写法（`bold` / `fontSize` / `fontColor` / `bgColor` / `horizontalAlignment` …）
  现在**声明在 schema 里**。处理器一直支持、测试也一直断言，但模型看不到就等于不存在 —— 这一版补上，与嵌套的
  `format.*` 等价，两种写法都可以（FIXES 78）。
- 广告面字节 37,573 → **38,713**（上限 60,000），工具数不变（69）。
- 顺带修好生成器 `scripts/extract-spec.mjs` 里一处**从未被触发的死引用**（`CONTAINER_PARAMS`）—— 参数分类一旦
  落到那一行就会让整个生成器崩掉；并清掉 P1-4 遗留的 9 条旧名映射表（名字已在 v0.5.1 对齐）。
- 两处测试矩阵把 `filePath` 写成 `path`（期望值宽松，所以写错也一直「通过」），已修正；现在它们报的是真正的原因。
## 0.5.1（参数命名对齐：公开名 = 桥键，ALIAS_DEBT 归零）

> 破坏性变更一处：**59 处工具参数改名**（例如 `filePath` → `path`、`app_type` → `appType`）。桥侧仍然接受
> 旧拼写，所以按旧技能或旧习惯写的调用照旧可用。

- **公开参数名与桥键一一对应**，别名表不再被任何工具依赖；`ALIAS_DEBT` 从 59 变成**硬 0** —— 新增工具
  必须照桥键命名，否则 CI 直接红。
- 41 组改名：最大一族是 `filePath` → `path`（8 处），其余是 `app_type` / `data_range` / `chart_type` /
  `slide_index` / `find_text` / `font_name` 这类 snake_case → camelCase，以及 `url`→`address`、
  `shapeIndices`→`names`、`effect`→`transition`、`chartType`→`type`、`order`→`zOrder` 等同义改名。
- 技能文档同步 5 处；已过时的反例换成**仍然存在**的真陷阱（Word 的 `insert_image` 收 `imagePath`，PPT 的
  `insert_ppt_image` 收 `path`；`set_slide_transition` 收 `effect`，`apply_transition_to_all` 收 `transition`）。
- 顺手修掉上一版漏掉的 3 个真机测试文件：`merged-tools` 因 `require` 已删除的模块而直接崩，另两个文件
  还在调 v0.5.0 已删除的旧工具名。
- 数字：测试 **914 项 / 46 文件**；注册工具 267、广告面 **69 / 37,573 字节**（内部上限 100 / 60,000）。
## 0.5.0（试点就绪：不卡死、结果如实、能解释、按不变量防守）

> 破坏性变更一处：**18 个废弃工具名已删除**（D4 的兼容窗口到期，见 docs/FIXES.md 76）。

### 稳定性（FIXES 68–72）

- **加密的 .pptx 不再能钉死会话**：所有打开路径在调用 COM 之前先做文件头预检（加密 OOXML 是 OLE/CFB，
  明文是 ZIP）；加密的 docx / xlsx 也从「1–4 秒后报错」变成 0 毫秒确定性拒绝。
- **弹窗守卫从「按清单」改成「按不变量」**：11 个 action（SaveAs / 转换 / 刷新 / 关闭 / 删表…）补齐
  DisplayAlerts 保护，并加静态门禁 scripts/lint-alerts.mjs（覆盖 / 配对 / 不绕过 / 例外有账）。
  实测发现 **DisplayAlerts 拦不住「另存为」**，因此对「从未落盘」的文档调用保存会被明确拒绝，
  而不是卡住或把文件静默写进「文档」目录。
- **结果如实**：close 的 saved 只在真的 Save 成功后才为真（失败就不关）；protect / unprotect 读回真实状态；
  deleteColumns 的影响统计覆盖整段；refreshLinks 标量化外链列表；PPT 转 PDF 改用 SaveCopyAs，不再把
  当前演示文稿改指成 PDF。
- **大范围预算四档**：读 5 万格 / 扫描 20 万 / 排版 20 万 / 逐格写 2 万，超限在执行前明确报错；
  超长路径（>260 字符）不再被误报成「文件不存在」。

### 可解释（FIXES 73）

- 桥侧收集的 warnings 现在能经**任意第一方工具**到达模型 —— 以前只有 wps_call / wps_execute_method /
  wps_batch 三条原样透传面能看到，约 260 个工具上的警告等于不存在。
- Word 侧补上共用的文档解析点：多文档同时打开而调用方没说目标时，会明确提醒落在哪一份上。

### 工程化（FIXES 74–76）

- **覆盖率口径分层**（有专门测试 / 仅矩阵 ok / 仅矩阵 any / 未驱动）并做成双向棘轮；三个矩阵测试里 81 行
  恒真断言改成要求返回体非空。
- scripts/lint-com-boundary.mjs 变成真门禁：return $range 必须 return ,$range；裸 catch { continue }
  必须登记理由。首次运行就抓到两处未包逗号的 return。
- **SKILL.md 里会漂的数字改成生成物**（GENERATED:advertised 标记，缺失即生成失败）：模型读到的广告清单
  与计数不再和注册表脱节；覆盖矩阵补上漏掉的「逃生舱」桶。
- 一键 e2e 适配**桌面版 DSH 运行时**与 **v4 会话日志**（session.v<N>.jsonl.zstd）。
- **删除 18 个废弃工具名**：wps_excel_zoom → wps_excel_set_zoom、wps_ppt_set_animation →
  wps_ppt_add_animation 等，旧名现在返回「未知工具」。

### 数字

注册工具 267、广告面 **69 / 37,573 字节**（内部上限 100 / 60,000）；测试 **927 项 / 46 文件**（多数需要本机 WPS）。

## 0.4.0（稳定性加固第 2 波：目标不漂、进程不堆、版本可见）

功能面与 0.3.0 完全一致：69 个广告工具 / 37,573 字节、267 个 action、256 对参数契约，一个都没变。
这一版继续只做**稳定性与可解释性**。

### 亮点

- **多文件打开时，目标歧义会明确告诉你**：不指定工作表 / 演示文稿时，动作落在「当前活动」的对象上，
  而它跟着窗口焦点走——一次超时重试就可能改到另一个文件。同时开着多个文件时，插件现在会在结果里附一条
  `warnings`，提示显式指定目标；只开一个文件、或你已经显式指定，都不会打扰。
- **参数形状错了立刻说清楚**：`data` 该传数组却传了 `"[[1,2]]"`、`background` 该传对象却传了字符串这类错误，
  在进 WPS 之前就被拒绝，报错直指参数名与期望类型。标量之间依旧宽松（`value: 42` 照旧可用）。
- **不再堆积 WPS 进程**：客户端被强杀（DSH 崩溃、任务被强杀）时，Windows 会把整棵进程树一起结束，
  宿主来不及收尾，它启动的 WPS 实例就留在那儿。现在宿主退出前会尽量收掉自己启动的实例，并把「归属」
  写进 `~/.wps-office-mcp/owned-apps.json`；下一个会话启动时会把上一个会话留下的孤儿收掉。
  **有未保存内容的实例一律不碰**，也不会去关你自己打开的 WPS。
- **`wps_status` 会报真实 WPS 版本**，低于 12.1 时给一句明确提醒。注意 WPS 通过 COM 报的 `Version` 是
  Office 兼容值（12.1 报 12.0），所以这里读的是 exe 的文件版本。

### 工程侧

- 新增项目 lint（手写 PowerShell 的 BOM/CRLF、制表符与行尾空白、`mcp/src` 的 `console.*`、测试退出码），已进 CI；
- 新增 `scripts/accept-install.mjs`：在一个全新的一次性 profile 里完整装一遍再拆掉；
- 新增 `scripts/run-tests.ps1`：整轮测试入口，跑完每个文件回收无头 WPS 残留；
- 全套测试 **816 项 / 37 文件 → 856 项 / 41 文件**；一键 e2e 28 → 29 项（新增归属记录断言）；
- 实测记录：`Application.Quit()` 返回成功并不保证 WPS 进程退出，因此进程级清理交给测试运行器，而不是断言。

## 0.3.0（稳定性加固第 1 波：不再卡死、不再撞车）

功能面与 0.2.1 完全一致：工具数、广告面（69 工具 / 37,573 字节）、参数契约、桥 action 都没有任何改动。
这一版只做**稳定性**——目标从「再加能力」转为「让客户敢用、出问题能解释、坏不了数据」。

### 亮点

- **不再被密码框卡死**：打开带密码的文件时 WPS 会弹一个模式对话框，此前会让整个 DSH 会话**永久卡住**
  （杀宿主也没用，框还在）。现在改成**不弹框、1 秒内直接报错**，并告诉用户怎么办。
- **同一时间只允许一个宿主**：同时开两个 DSH 会话操作 WPS 时，第二个会话会收到一句**可读的中文错误**，
  而不是两边互相卡死。
- **超时不再等于「已失败」**：超时文案明确说明「状态未知（可能被对话框阻塞）」，并且**不会自动关闭 WPS**
  （那会丢掉你没保存的内容）；之后会先用短超时快速失败，**成功一次即恢复**。
- **修掉两个恢复路径缺陷**（由本轮新测试抓出）：陈旧子进程的 `exit` 会反杀刚拉起的新宿主；
  `ready` 帧误清 `suspect` 标志，使短超时形同虚设。

### 实测依据

`docs/FIXES.md` 第 52 条记录了完整探针数据。两处计划假设被实测推翻：`PasswordDocument:=""` **挡不住**
密码框（空串等于没给密码）；扩展名不符的老 `.doc` 实测**不弹**转换框。真正有效的是传**非空哨兵密码**——
加密 `.docx` 从「永久卡死」变成 4 秒内报错，加密 `.xlsx` 1 秒内报错（`0xFFF40006`）。
另一处偏离计划：超时后**不杀整个 WPS 进程树**，因为实测本机有 243 个进程叫 `wps`、且 Word/PPT 的 COM
对象不提供 `Hwnd`，「被卡住的是哪一个」无法可靠指认，照做会连带杀掉用户整套 WPS。

### 已知限制（如实记录）

- 加密的**演示文稿（.pptx）**是唯一没堵住的口子：`Presentations.Open` 没有密码参数，无法阻止弹框；
  请先用 WPS 手工打开它并另存为不加密副本，再对副本操作。
- `wps_execute_method` 逃生舱下，`Document.Password` 按长度失效：15 字符卡死调用，**17 字符静默写出
  完全不加密的文件**。本插件的正式工具里没有「设置打开密码」的能力。

### 验收

`test/open-safety.test.mjs` 21 项、`test/host-lease.test.mjs` 18 项、`test/watchdog.test.mjs` 15 项全绿；
全套 **595 项 / 28 文件** + `verify.mjs` 23 项。后两个测试不需要 WPS，已并入 GitHub Actions 静态门禁。

### 安装

```powershell
dsh plugin --profile <profile> add github:sueccku/dsh-plugin-wps-office-next#v0.3.0
```

连不上 github.com 时改用：

```powershell
dsh plugin --profile <profile> add https://codeload.github.com/sueccku/dsh-plugin-wps-office-next/tar.gz/refs/tags/v0.3.0
```

## 0.2.1（仅文档与打包修正，无行为变化）

功能与 0.2.0 完全一致：工具数、广告面、参数契约、桥 action 都没有任何改动。这一版只修文档与打包。

### 改动

- **README 重写为面向使用者的两段式**：新增「它适合谁 / 能帮你做什么 / 你可以这样提要求」，安装改成
  「把一段话复制给你的 DSH AI」+ 一节写给 AI 的 7 步安装指引（环境要求、判断 profile、执行安装、
  确认接线、重启验证、排错表、环境自检），并补上卸载、重要注意事项与已知限制；技术细节收进折叠区。
- **`CHANGELOG.md` 现在会随包安装**（此前 `files` 白名单漏了它，0.2.0 的安装包里没有这个文件）。
  `README.md`、`LICENSE`、`THIRD_PARTY_NOTICES.md` 原本就在白名单里，现在文档齐了。
- **安装指引逐条实测过**：profile 探测脚本、codeload 备用地址、`--dump-config` 的接线检查、
  `scripts/doctor.mjs`、`remove` 子命令、批量上限 50——全部与本仓库实际行为一致。
- **修掉一处会把人带偏的提示**：`dsh plugin add` 失败时，dsh 会补一句「构建脚本被拦截 / 请加 allowBuilds」
  的通用兜底文案；本包没有 `prepare` 脚本、也没有原生依赖，**永远不需要改 allowBuilds**。
  真因通常是这台机器连不上 github.com，用 codeload 地址重试即可（README 已写明）。

### 安装

```powershell
dsh plugin --profile <profile> add github:sueccku/dsh-plugin-wps-office-next#v0.2.1
```

连不上 github.com 时改用：

```powershell
dsh plugin --profile <profile> add https://codeload.github.com/sueccku/dsh-plugin-wps-office-next/tar.gz/refs/tags/v0.2.1
```

## 0.2.0（首个发布）

把 [lc2panda/wps-skills](https://github.com/lc2panda/wps-skills) 与
[CatNebulaaaa/wps-dsh-plugin](https://github.com/CatNebulaaaa/wps-dsh-plugin) 合并成**一个 DSH bundle**：
DSH 插件 + 自带 MCP server + 常驻 COM 宿主 + 全套技能。仅 Windows、仅 COM，不需要任何 WPS 加载项。

**广告面比上游基线小 73.5%，能力反而更全**：250 工具 / 141,872 schema 字节 → **69 工具 / 37,573 字节**；
注册工具 267 个（其余经 `wps_call` 与技能参考表触达）；541 项测试全部驱动**真实 WPS**。

### 亮点

- **常驻 COM 宿主**：取代「每次新起一个 PowerShell」，冷启动约 1.0s、稳态 1–2ms，支持串行队列与崩溃重启
- **契约真源**：工具 schema、桥的键表、技能参考表都由一份操作规格生成，CI 断言「源码与产物逐字节一致」
- **Excel 做深**：表（ListObject）全族、条件格式与数据验证的读与删、页面设置与打印（打印标题/页眉页脚/横向 A4）、
  单变量求解、迷你图、透视表刷新与清除、公式审计（引用与被引用）、命名范围读写
- **Word 做深**：表格读写编辑（增删行列/合并拆分/样式）、页码、分栏、修订列表与接受拒绝、批注读删、
  内容控件、脚注尾注、索引、交叉引用、**邮件合并**（CSV → 生成新文档，母版不动）
- **PPT 做减法**：放弃 3D 族与美化族，把碎片 setter 归并（形状效果四合一、动画三合一、表格样式三合一、
  页脚三合一），88 → 76 个工具
- **一键 e2e**：一条命令造 fixture、跑真实 headless 任务、逐帧解会话日志、用裸 COM 重开产物核对，28 项检查

### 修好的「从未生效」缺陷

这些能力早就写在桥里，却因为宿主环境差异或参数写法而**永远不工作**，此前没有任何测试覆盖：

1. **查找**：`Range.Find()` 结果的 `Address()` 在常驻宿主里不可用 → 改为按 A1 串取范围、一次 `Value2` 扫描
2. **二维数组下标**：PowerShell 的逗号优先级高于 `+`，`$m[$a + $r, $b + $c]` 被解析成
   `$a + ($r, $b) + $c`，异常又被裸 `catch { continue }` 吞成「0 命中」→ 下标加括号
3. **合并计算 / 分类汇总**：`XlConsolidationFunction` 常量写成了假值，每次调用都 HRESULT 0x800A03EC
4. **表（ListObject）**：`Delete()`/`Add()` 之后同一个对象仍返回旧几何，工具会把变更前的结构报给用户
5. **透视表**：`Get-RangeFromAddress` 的 `return $range` 被 PowerShell 展开成 21 个单格，创建透视表**从来没有成功过**
6. **脚注之后**：光标留在注释正文里，此后所有「往光标处插内容」的动作都插进了注释（邮件合并因此静默出错）

### 实测的 WPS 限制（不做，如实记录）

- **水印**：WPS 的页眉 `Shapes` 集合不接受任何图形（`AddTextEffect`/`AddShape`/`AddTextbox` 都返回对象但 `Count` 恒为 0）
- **文档属性**：`BuiltInDocumentProperties` / `CustomDocumentProperties` 是坏壳（`.Item()`、`GetType()` 直接抛 null 引用）
- **切片器**：`SlicerCaches.Add2` 建得出缓存，但 `Slicers.Count` 恒为 0，用户可见的切片器不会出现
- **场景管理器**：`Worksheet.Scenarios` 在 COM 里被暴露成方法，语义读不干净

以上四项都可以用隐藏逃生舱 `wps_execute_method` 自行尝试（默认不在广告面里）。

### 安装

```powershell
dsh plugin --profile <name> add github:sueccku/dsh-plugin-wps-office-next#v0.2.0
```

要求：Windows、WPS Office 12.1+ x64。预构建产物已入库，安装后开箱可用，无需构建步骤。

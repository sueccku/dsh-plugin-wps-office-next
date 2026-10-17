# Skill 重构与代码修复报告（阶段 D 交付物）

> 依据：改进指南 §11 阶段 D —— 列出**修改文件、验证结果、未完成项与剩余限制**。
> 审计依据（改动前的逐条核对）见 [skill-audit.md](skill-audit.md)；固定场景与比较方法见 [skill-evaluation.md](skill-evaluation.md)。
>
> **一句话结论**：四个 Skill 已按"任务决策指南 + 按需参考"重写，生成器不再往主文件里注入工具清单，
> 三项 P0 语义缺陷与一项编码缺陷已修复并有真机证据；**模型侧的 Skill 效果对比没有跑**，
> 因此本报告不给"成功率提升"之类的结论。

## 1. 修改文件清单

### 1.1 Skill 内容（重写 / 新增）

| 文件 | 类型 | 说明 |
|---|---|---|
| `skills/wps-office-next/SKILL.md` | 重写 | 共享执行协议：路由、任务分类、目标与输出策略、工具发现、验证、保存、失败恢复、按需参考入口 |
| `skills/wps-excel/SKILL.md` | 重写 | 表格决策流程：三层级、先读懂数据、标准工作流、值/公式/数字格式三分、输出策略、核对 |
| `skills/wps-word/SKILL.md` | 重写 | 文字决策流程：内容 vs 版式、三层格式、字符坐标规则、排版顺序（先全局后局部）、目录页码、模板 |
| `skills/wps-ppt/SKILL.md` | 重写 | 演示决策流程：先画布→结构→版式、**混合单位表**、占位符矩阵、页序回读、信息图组合、视觉检查 |
| `skills/*/reference.md` | 重新生成 | 完整工具表 + **按任务分组索引**（9/7/8/4 组）+ 直达/`wps_call` 标记 |
| `skills/wps-office-next/references/execution-contract.md` | 新增 | 目标身份读回协议、结果字段语义、保存/副本/备份、失败分流、验证清单 |
| `skills/wps-office-next/references/compatibility.md` | 新增 | 版本、宿主超时、单实例租约、孤儿回收、对话框规避、各应用实测异常、不要依赖的东西 |
| `skills/wps-office-next/recipes/cross-app-report.md` | 新增 | 跨应用任务配方（表格→文字/演示→PDF），身份切换与对应关系核对 |
| `skills/wps-excel/references/ranges-and-formulas.md` | 新增 | `set_formula` 区域写入的实测行为、逐行公式的三条正确做法、重算与核对、范围预算分档 |
| `skills/wps-excel/recipes/summarize-report.md` | 新增 | 汇总报表配方（口径、公式策略、图表源、打印导出、总额核对、原件保护） |
| `skills/wps-word/references/ranges-styles-sections.md` | 新增 | 字符坐标来源/单位/失效、段落标记语义、`set_font` vs `apply_style`、分节与页眉页脚、目录、字段类型 |
| `skills/wps-word/recipes/format-existing-report.md` | 新增 | 排版已有报告（先全局后局部、逐节、目录、核对表） |
| `skills/wps-word/recipes/fill-template.md` | 新增 | 模板批量填写（字段类型判定、邮件合并 vs 手工、只重试未完成项） |
| `skills/wps-ppt/references/layout-and-shapes.md` | 新增 | 单位换算表（磅/像素/秒）、画布对照、版式占位符、形状索引、版式计算模板、文字容量 |
| `skills/wps-ppt/recipes/create-deck.md` | 新增 | 从零做演示的完整流程 |
| `skills/wps-ppt/recipes/infographic-layout.md` | 新增 | 11 类信息图的坐标算法与验收点 |

### 1.2 生成器与门禁

| 文件 | 变更 |
|---|---|
| `scripts/gen-skill-tools.mjs` | 重写：`reference.md` 出完整目录 + 任务分组索引；SKILL.md 的生成块**只留计数与参考入口**；新增 `--check`（只读、比对、有差异非零退出）；加请求/整体超时、子进程退出处理、`StringDecoder` 收帧；只做 `initialize` + `tools/list`（不发起任何 COM 业务操作） |
| `spec/skill-tool-groups.json` | 新增：任务分组元数据（只写工具名，不重复维护 schema）。生成器**硬校验**：未知名字、重复归类、注册工具未被归类 → 生成失败 |
| `test/skill-quality.test.mjs` | 新增静态门禁 123 项：注册名/元数据、生成块不含工具清单、相对引用存在、`wps_*` 名字都在注册目录、**调用示例的顶层参数必须在该工具 runtime schema 里**、误导说法黑名单、`--check` 幂等且不写盘 |
| `test/plugin.test.mjs` | 扩展：每个 Skill 必须有 `references/` 与 `recipes/` 且非空 |
| `scripts/verify.mjs` | 扩展 `--static`：**每个注册工具都必须出现在某个 skill reference.md 里**，且 reference 能分辨「直达 / wps_call」 |
| `scripts/doctor.mjs` | 扩展：按需参考结构存在性 + SKILL.md 里 `references/`/`recipes/` 链接可达 |
| `scripts/lib/static-tests.mjs` | 把 `test/skill-quality.test.mjs` 纳入静态测试清单（CI 与本地共用唯一来源） |
| `.github/workflows/ci.yml` | 新增 `gen-skill-tools.mjs --check` 步骤（先证明只读模式无差异，再走写入 + git diff 对账） |
| `README.md` | 数字刷新（standard 55,335 字节 / full 165,241）；**修掉「未广告工具可以按全名直接调用」这条误导**；补 `wps_batch` 的 `stopOnError`/`status` |
| `docs/error-contract.md` | 批量契约补 `stopOnError`、`status/succeeded/failed/stoppedAt`、逐项 `index`、错误不截断、`partial`/`shortfalls` |

### 1.3 代码修复

| 文件 | 变更 | 级别 |
|---|---|---|
| `mcp/scripts/wps-com.ps1` | `closeWorkbook`/`closeDocument`/`closePresentation`：未落盘 **且** 有未保存改动 **且** 要求保存 → **不关闭**、保留内容、返回明确错误（原来降级成"不保存关闭"并静默丢弃内容）；新增 `Test-WpsObjectClean`、`Format-WpsUnsavedCloseRefusal` | **P0** |
| `host/wps-com-host.ps1` | 租约：身份核对通过就不再判过期，心跳只作诊断（原来"心跳 >120s"会杀掉仍存活的健康宿主） | **P0** |
| `mcp/scripts/wps-com.ps1` (`Save-WpsOwnedApps`/`Invoke-WpsOrphanReclaim`) | 回收记录升级到 `version:2` 并带**实例身份**（类名 / 安装目录 / 窗口句柄 / 当时打开的文件清单）；回收前逐项核对，身份不符或"实例里出现了记录时没有的文件"一律不回收；旧记录向后兼容 | **P0** |
| `mcp/scripts/wps-com.ps1` | `Get-Content -Encoding UTF8`：修掉"记录里中文文件名被 ANSI 解码成乱码 → 身份核对永不相等 → 孤儿永远回收不掉"的真实缺陷 | **P0 伴生** |
| `host/wps-com-host.ps1` | 每次动作后刷新归属记录（让身份核对拿到最近状态，而不是"实例刚建好那一刻"的旧快照）；`Read-HostState` 显式 UTF-8 | P0 伴生 |
| `mcp/src/client/com-host.ts` | **按字节收帧**：`Utf8LineReader`（`StringDecoder`）替换 `chunk.toString()`，chunk 边界切开中文不再变 U+FFFD；stderr 同样走解码器；换子进程时 `reset()` | **P1（改进指南 §9 明确列出）** |
| `mcp/src/server/mcp-server.ts` + `mcp/src/spec/operations.ts` | `wps_batch` 增加 `stopOnError`；整体返回 `status: completed/partial/failed`、`requested/succeeded/failed/stoppedAt`；逐项带 `index`；失败项的 `error` **不截断**；整批 `success` 语义不变（只表示"批处理跑完"），已在文档与 Skill 里写清 | **P1** |
| `mcp/scripts/wps-com.ps1` (`getExcelContext`, `getSlideInfo`) | 目标身份可读回：Excel 加 `workbookPath`；PPT 加 `presentationName`/`presentationPath` 与画布 `slideWidth/slideHeight`（磅）+ `slideWidthPx/slideHeightPx`（像素） | **P1** |
| `mcp/src/tools/excel/missing-halves.ts`, `mcp/src/tools/ppt/slide-ops.ts` | 把上面的身份字段**打进模型看得到的那段文本**（桥有、handler 不打印 = 能力对模型不存在） | **P1** |
| `mcp/src/tools/excel/workbook.ts` | `wps_excel_open_workbook` 补 `readOnly`/`updateLinks`（桥本来就读这两个键，schema 没声明 → 工具层会拒绝，导致"读来源数据不动原件"没有入口） | **P1** |
| `mcp/src/tools/excel/formula.ts` | `wps_excel_set_print_area` 补 `sheet`（桥按 `sheet` 定位工作表，schema 没声明） | **P1** |
| `mcp/src/tools/ppt/slide-ops.ts` | `wps_ppt_set_slide_size` 描述里"常用尺寸"混用磅与像素（16:9 写 960×540 其实是磅值）→ 明确标单位并给出正确像素值 | **P1** |
| `host/wps-actions.ps1`, `spec/*.json`, `mcp/dist/**` | 由源文件重新生成 | — |

### 1.4 测试

| 文件 | 变更 |
|---|---|
| `test/close-safety.test.mjs` | 新增 A4：脏 + 未落盘 + `save:true` 在 Excel/Word/PPT 三侧都必须**拒绝关闭**、内容仍在、`save:false` 仍可关；含裸 COM 的 `Saved` 前提断言 |
| `test/orphan-identity.test.mjs` | **新增**：归属记录写出身份；类名不符不回收；文件清单出现新文档不回收；身份相符才回收且退出 ROT；旧记录向后兼容；身份读不出来则回绝 |
| `test/target-identity.test.mjs` | **新增**：Excel/Word/PPT 的身份读回（含未落盘时如实说"尚未保存到磁盘"）、画布磅↔像素换算与裸 COM 交叉验证 |
| `test/host-lease.test.mjs` | 新增 A4：**心跳停更不能把仍存活的宿主判成过期**（改老心跳后第二个宿主仍须拒绝、宿主仍须存活） |
| `test/error-contract.test.mjs` | 新增：`status/succeeded/failed`、逐项 `index`、错误不截断、`stopOnError` 停止位与跳过、全失败 → `status:"failed"` |
| `test/com-host.test.mjs` | 新增：中文帧在 6 种字节切法下都能还原、逐字节收帧、半帧不吐、`reset()` 不串帧、端到端"桥的中文原样到达客户端" |
| `test/honest-reporting.test.mjs` | 按新契约改判 2 条断言（旧断言编码的是"降级关闭 + warning + success"） |
| `test/silent-catch.test.mjs` | 登记新加的一条 best-effort 空 catch（门禁要求每条都有理由） |

## 2. 实测范围与结果

### 2.1 静态门禁（不需要 WPS）

| 命令 | 结果 |
|---|---|
| `npm run lint` | scanned=225 files, violations=0 |
| `node scripts/verify.mjs --static` | **VERIFY OK (21 checks)** |
| `node test/spec-reproduction.test.mjs` | **SPEC REPRODUCTION OK (16)** |
| `node scripts/run-static-tests.mjs` | **15/15 文件全绿**（plugin 40、skill-quality 123、com-host 17、host-lease 23、watchdog 15、silent-catch 8、install-selfcheck 6、mcp-node 16、arg-shape-guard 10、wps-version 14、encrypted-preflight 15、alerts-gate 10、com-boundary-gate 16、deprecated 10、style-names 28） |
| `node scripts/gen-skill-tools.mjs --check` | 8 个文件全部 ok，未写工作区 |
| `node scripts/gen-numbers.mjs --check` | CLAIMS OK (20 checked) |
| `node scripts/param-contract.mjs` | tools=270 checked=259；A/B/C/D 四类静默失效均为 0 |
| `npm run verify:package` | **PACKAGE RUNTIME OK (11)**（打包 305 项、装包、MCP 握手、tools/list=95） |
| `cd mcp; npx tsc` + 重新生成 `host/wps-actions.ps1` / `spec/*` | 全部成功，265 个 action、解析 0 错误、0 个 `exit` 残留 |

### 2.2 真机测试（本机 WPS 12.1.0.29158，全部显式把 `WPS_OFFICE_HOST_SCRIPT`/`WPS_OFFICE_MCP_ENTRY` 钉到本仓库）

> 为什么必须钉：本 DSH 会话把插件装在了 profile 里，`WPS_OFFICE_HOST_SCRIPT` / `WPS_OFFICE_MCP_ENTRY`
> 指向**安装副本**；不钉就会测到 0.8.2 那份旧代码（第一次跑 close-safety 时正是这样先"假绿"后"假红"的），
> 连 `test/plugin.test.mjs` 也会因为 `plugin.js` 保留预置的 `WPS_OFFICE_MCP_ENTRY` 而报一条假失败。
> `scripts/run-static-tests.mjs` / `run-tests.ps1` 已经这么做了，**单独跑某个测试文件时要手动钉或清掉这两个变量**。

| 文件 | 结果 | 关键证据 |
|---|---|---|
| `test/close-safety.test.mjs` | **36/36** | 脏+未落盘+`save:true` 三侧都拒绝关闭、内容仍在、`save:false` 可关；裸 COM 交叉验证没有工作簿/演示泄漏 |
| `test/target-identity.test.mjs` | **26/26** | `workbookPath` = 落盘路径（裸 COM 一致）；未落盘如实回报；PPT 画布 `960×540 磅（1280×720 像素）` 与裸 COM `PageSetup.SlideWidth` 一致 |
| `test/orphan-identity.test.mjs` | **22/22** | 记录带身份；类名不符/文件清单不符都不回收（实例仍在 ROT）；身份相符才回收并退出 ROT；旧记录放行、身份读不出回绝 |
| `test/error-contract.test.mjs` | **20/20** | `status`、逐项 `index`、错误不截断、`stopOnError` 停在失败项、全失败 `status:"failed"` |
| `test/honest-reporting.test.mjs` | **21/21** | 新契约下"不把未保存说成已保存"更强（直接拒绝而不是丢弃内容） |
| `test/error-wording.test.mjs` | **31/31** | 三段式错误文案与幂等性未被新错误破坏 |
| `test/warning-channel.test.mjs` | **14/14** | 批量/透传通道的 warning 仍逐字保留 |
| `test/word-lifecycle.test.mjs` | **18/18** | 关闭/保存语义 |
| `test/excel-missing-halves.test.mjs` | **20/20** | `get_sheet_info` 文本新增字段未破坏既有断言 |

**没有跑的**：`scripts/run-tests.ps1` 整轮（按 `AGENTS.md` §1/§2，需要用户明确同意）、`scripts/e2e.mjs`、
`scripts/accept-install.mjs`、`test/orphan-reclaim.test.mjs`（它会强制结束本机所有 WPS 进程，属于破坏性；
本轮的 `orphan-identity` 用受控记录覆盖了同一段逻辑，但没有覆盖"强杀客户端→下次接管"那条端到端路径）。

## 3. 未完成项与剩余限制（明确不假装已支持）

1. **模型侧效果对比未跑**：见 `docs/skill-evaluation.md` §1/§3。场景与成功标准已固定，但新旧 Skill 的
   成功率、调用数、加载 tokens 一个都没有实测；本轮**不能**声称"模型更容易选对工具"。
   （后续补充：可执行脚手架已落地在 [test/skilleval/](../test/skilleval/README.md) —— 6 个场景有夹具与
   确定性断言，正例/反例都验证过会 PASS/FAIL；跑模型会话属于下一步，需要用户同意。）
2. **Harness 的 Skill 加载 / 相对路径读取未验证**：改进指南 §3.2 要求用真实 Harness 会话确认
   "应用 Skill 被单独选中时总 Skill 是否自动加载""`references/`/`recipes/` 能否读到"。
   本环境没有交互会话 → 已按降级方案处理（四个应用 SKILL.md 都保留关键协议短摘要）。
   （后续补充：`wpse2e` profile 已用 junction 指向工作区，是这条验证的可用入口。）
3. **没有 documentId**：本轮实现的是"**身份读回**"（名称 + 路径 + 画布 + 前后各读一次对账），
   这**只能缓解**调用之间窗口焦点切换的竞争，不能消除。Skill 里没有一处声称"已锁定目标"。
4. **静态值/公式的边界**：`set_formula` 对多格区域不做相对引用调整是**当前实现**，本轮把它写成
   明确的限制与三条正确做法，没有改成广播语义（那会是底层重写）。
5. **嵌套参数仍不递归校验**：ToolRegistry 只校验顶层键；Skill 里明确写"嵌套对象按 schema 自己核对"，
   没有把承诺扩大到所有嵌套参数。
6. **`wps_batch` 的 `success` 语义保持兼容**：整批 `success:true` 仍表示"批处理跑完了"。已用
   `status`/逐项 `success` 补上结构化判断，并在契约文档与 Skill 里写清；**没有**改成"有子项失败就整体失败"
   （那会推倒既有契约与测试）。
7. **视觉验收未做**：导出图片与图像读取能力未在本环境验证；Skill 只允许说"内容与属性已核对，视觉效果未验证"。
8. **孤儿回收的正面路径**（强杀客户端 → 下次会话回收）本轮没有端到端跑；`test/orphan-reclaim.test.mjs`
   是它的回归文件，属于需要用户同意的整轮范围。另外本次修复也暴露了这条路径上的一个新增保护：
   记录刷新是**按动作**做的，客户端恰好在"动作完成、记录刷新之间"被强杀时，新会话会因文件清单不符而
   **跳过回收**（安全方向：宁可留下孤儿，也不误关用户实例）。
9. **PowerShell 模块化（改进指南 §9 P2）未做**：`host/wps-actions.ps1` 仍是文本改写生成。
10. **`wps-office-next/SKILL.md` 142 行**，高于改进指南 §4.2 的 80–130 行建议值。它承载的是共享协议
    （路由 + 目标 + 保存 + 验证 + 失败分流 + 参考入口）；按"必要知识优先于字数"保留，未做删减。

## 4. 与改进指南 §2.3 的逐条对账

| # | 条目 | 结果 |
|---|---|---|
| 1 | 「可选参数写错名不会被发现」 | 已重写为当前规则（未知顶层参数被拒 + 嵌套不递归） |
| 2 | Word：`close_document` 广告状态矛盾 | 已删除手写广告状态，状态一律由 `reference.md` 生成 |
| 3 | 「Save 可保留回退版本」 | 已改成保存/另存副本/备份三分，并写明 Save 覆盖磁盘原件 |
| 4 | 「未连接就不能继续」 | 已改成状态查询只作诊断；打开/新建会启动 WPS |
| 5 | PPT：「插图用 path」 | 已改成 `imagePath`，并写进 schema 检查门禁 |
| 6 | PPT：「x/y 会静默忽略」 | 已删除旧失败行为描述 |
| 7 | PPT：「演示工具最多」 | 已删除该比较句 |
| 8 | PPT：已移除的高层工具又出现在已知坑 | 已不复现；并把"文档里的 `wps_*` 名字必须都在注册目录"固化成静态检查 |
| 9 | Word：先标题后统一正文字体会覆盖标题 | 已改成「先全局、后局部覆盖」并给出只改正文坐标的做法 |
| 10 | Excel：逐行公式建议写静态值 | 已改成保留公式（逐格 `set_formula` / `fill_series`）+ 回读公式与结果 |
| 11 | 收尾默认关闭文档 | 已改成按任务决定；保存不必关闭 |
| 12 | 按 `accepted` 键重试 | 已改成区分调用错误与插件内部映射错误 |
| 13 | 新文档 `save=true` 仍丢弃关闭 | **已修代码**（保留打开 + 明确错误），并有真机回归 |
| 14 | 「未广告工具可按全名直接调用」 | 已删除；README 同步修正 |
| 15 | 演示单位笼统称"像素" | 已重写为混合单位表（形状=磅、画布/导出=像素、1px=0.75pt），真机验证 |
| 16 | `warnings` 表述不完整 | 已补 `success`/`partial`/`shortfalls`/`warnings` 四件套语义 |
| 17 | `wps_batch` 说明不足 | 已补 `stopOnError` 与整体状态；文档同步 |
| 18 | 无 documentId 却像已锁定 | 已写清"仍依赖活动对象"，并实现身份读回 |
| 19 | 超时"宿主自动重启"说太少 | 已补"状态未知 + 非幂等操作禁止盲目重放" |

# 错误与超时契约

> 目的：让「失败长什么样、超时后该怎么办」有唯一可查的说法，并且**每条语义都有断言**。
> 真源：`mcp/src/client/com-host.ts`（超时）、`mcp/src/server/mcp-server.ts`（批量与门面）、
> `mcp/scripts/wps-com.ps1`（动作错误与 warnings）。断言位置见文末。

## 1. 结果信封

每个工具返回 `{ success, content, error? }`，MCP 层把它映射成 `isError: !success`：

- `success: true`：`content` 里有结果；桥收集到的 `warnings` 会附在结果 JSON 的 `data.warnings`，**并且**（FIXES 73 起）第一方工具还会把它追加到结果文本末尾，见 §1c。
- `success: false`：`content[0].text` 与 `error` 是同一句中文原因。

`warnings` 的语义是「主要操作完成，但某个次要步骤失败了」——不是失败，但需要看。

## 1b. 面向客户的错误文案（S7）

桥里所有 `success:false` 的错误（含宿主捕获的异常）都会经过 `Format-WpsErrorText` 统一成三段式：

```
<中文一句话 或 原文>（动作：<bridge action>）下一步：<可执行的建议>。（原始信息：<英文/HRESULT>）
```

- 命中的英文短语（`WPS Excel not running`、`No active document`、`table not found on this sheet`…）换成中文，
  原始英文放进「原始信息」；
- **没命中的文案原样保留**，只补动作名与下一步（例如 `shapeIndex/shapeName is required（动作：setAnimation）下一步：…`），
  这样历史断言与模型重试都不受影响；
- 动作名是桥的 action（如 `openWorkbook`），模型可以直接拿它重试；
- 文案是**幂等**的：已经带「（动作：」的不会再包一层。

断言：`test/error-wording.test.mjs`（最近一次整轮 **31 项**：抽样必失败路径 + 动作名 + 幂等；分组明细以测试文件为准，读数见 `test/summary.json`）。

## 1c. 目标歧义警告（C7）

不指定目标时，动作落在「活动」对象上（`ActiveWorkbook` / `ActiveSheet` / `ActivePresentation`），而它跟着窗口焦点走：
一次超时重试、或同一会话里的第二次调用，就可能落到另一个文件上。所以桥在**确有歧义**时补一条 warning：

- **触发**：该解析点没拿到显式目标名，且**打开的工作簿 / 演示文稿 > 1**；同一动作内按消息去重。
- **解析点**：`Get-WorksheetByParam`（`sheet`）、`Resolve-Worksheet`（`sheet` / `name` / `oldName`）、
  `Get-TargetPres`（`presentationName`）。
- **消除**：显式传目标名；单文件时不打扰。
- **仍未覆盖**（记在 `baseline/known-defects.md` C7）：Word 的文档解析、`transpose` 目标表、`copySheet` 源表、已用范围探测。
- **能到哪儿**：**所有工具**。FIXES 73 起，桥侧 warnings 会经第一方工具到达模型 —— `wps_call` / `wps_execute_method` /
  `wps_batch` 原样透传；其余工具由 `mcp/src/server/tool-registry.ts` 收集后追加到结果文本
  （`mcp/src/utils/tool-warnings.ts`，AsyncLocalStorage）。
- **注意**：`wps_batch` 单项结果会截断到 2000 字符，追加的 warnings 可能被截掉（FIXES 73 的残留，单项调用可避免）。

断言：`test/target-ambiguity.test.mjs`（真实 WPS）。

## 2. 批量契约（`wps_batch`）

| 情况 | 行为 |
| --- | --- |
| `calls` 为空 | 整体失败：`calls 不能为空` |
| `calls` > 50 | 整体失败：`单次批量最多 50 项`；**一项都不执行** |
| 某一项工具名无效 / 是门面 / 未知 | 该项记 `{index, tool, success:false, error:"无效或不可调用的工具名"}`，**默认继续执行后面的项** |
| 某一项执行失败 | 该项记 `{index, tool, success:false, result:"<结果文本>", error:"<完整错误，不截断>"}`（有 `partial`/`shortfalls` 时一并带上），**默认继续执行后面的项** |
| 某一项成功 | 该项记 `{index, tool, success:true, result:"<结果文本，最多 2000 字符>"}` |
| `stopOnError: true` | 遇到第一个失败项即停止，`stoppedAt` 记下失败项下标，后面的项**不执行** |

- 整批的返回值**恒为 `success: true`**（它表示"批处理本身跑完了"），结构
  `{count, requested, succeeded, failed, status, stopOnError, stoppedAt, results[]}`；
  局部失败不会把整批标成失败，所以调用方**必须看 `status` 或逐项看 `results[i].success`**。
- `status` 取值：`completed`（全部成功）/ `partial`（有成功也有失败）/ `failed`（一项都没成功）。
- `result` 会被截断到 2000 字符（长度压力在正文上）；`error` **不截断** —— 截断过的错误没法诊断。
- 有依赖关系的步骤（先建表再写数据）用 `stopOnError: true`，或干脆逐项调用；默认的"继续执行"会让
  后续步骤落到错误的前置状态上。

## 3. 超时契约

| 项 | 值 | 环境变量 |
| --- | --- | --- |
| 默认超时 | 60000 ms | `WPS_OFFICE_TIMEOUT_MS` |
| 长动作超时 | 300000 ms | `WPS_OFFICE_LONG_TIMEOUT_MS` |
| 可疑（suspect）短超时 | 15000 ms | `WPS_OFFICE_SUSPECT_TIMEOUT_MS` |

长动作集合（`LONG_ACTIONS`，11 个）：`convertToPDF`、`convertFormat`、`exportChartAsImage`、
`exportRangeAsImage`、`exportSlideAsImage`、`createPivotTable`、`updatePivotTable`、
`beautifySlide`、`insertSlidesFromFile`、`calculateSheet`、`saveAs`。

> FIXES 81 修掉了这张表里的三个**废名**（`beautify` → `beautifySlide`、`recalculate` → `calculateSheet`、
> `proofreadBasic` —— 它根本不是桥 action，校对是纯 TS 实现）：名字写错不会报错，只会让那个动作悄悄退回
> 60s 默认超时、慢操作跑到一半被杀成「状态未知」。`scripts/verify.mjs` 现在会断言每个名字都是真实 action。

超时后**不是**「已失败」，而是：

1. 宿主进程被杀（它卡在 COM 里，不可复用）；
2. **WPS 不关**（可能还有未保存内容，也可能只是被对话框阻塞）；
3. 置 `suspect`：后续调用先用 15s 短超时快速失败，**成功一次即恢复**常规超时；
4. 文案固定为三段式：动作 + 超时毫秒 + 「状态未知」+ 下一步。

## 4. 调用方该做什么

| 症状 | 该做什么 |
| --- | --- |
| 超时，文案含「状态未知」 | 切到 WPS 处理弹框（密码 / 保存 / 恢复），然后重试；**不要**以为操作已经失败 |
| 批量里某项 `success:false` | 只看那一项；它前面的项**已经执行**，默认后面的项**也会继续执行**（要停就用 `stopOnError: true`） |
| 批量的 `status` | `completed` / `partial` / `failed` —— 整批 `success:true` 只表示"跑完了" |
| 结果里带 `warnings` | 主操作已生效，按 warning 内容人工确认次要步骤 |
| 结果里带 `partial` / `shortfalls` | 你要求的事**没有全部生效**（`success` 已经是 `false`）：读已生效范围，只补做缺失部分 |
| `输出目录不存在：<路径>`（`save_as`） | 先创建该目录或把 `outputPath` 换到已存在的目录；桥**不会**替调用方创建目录。以前这种情况只回一句「文档保存失败」，模型会白试两轮（D1 实测） |

## 5. 断言位置

| 语义 | 断言 |
| --- | --- |
| 超时文案含「状态未知」、短超时、成功即恢复、不偷偷重试 | `test/watchdog.test.mjs` |
| 批量上限 50、空批量、部分失败继续、`stopOnError` 停止位、`status/succeeded/failed`、结果截断 2000、错误不截断、门面/未知工具拒绝 | `test/error-contract.test.mjs` |
| 空 catch 的静默必须登记 | `test/silent-catch.test.mjs` |
| 错误文案三段式、动作名、幂等 | `test/error-wording.test.mjs` |
| `save_as` 缺输出目录的错误文案与"不替调用方建目录" | `test/file-ops.test.mjs` |
| 广告面、桥动作数、门面行为、长动作名单都是真实 action、Skill 参考覆盖全部注册工具 | `scripts/verify.mjs` |
| Skill 结构、引用可达、示例参数符合 schema、生成块不含工具清单 | `test/skill-quality.test.mjs` |

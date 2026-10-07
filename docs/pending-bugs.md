# 待修清单（P1–P8 **全部已修**）

> **状态：P1–P8 全部已修（2026-10-04，见 [FIXES.md](FIXES.md) 的 92 号）。**
> 本文件保留为**发现过程的证据记录** —— 每条都带复现步骤与实测数据，修法与验证结果见 FIXES 92。
> 确认无误后可以删掉本文件（它不在 `package.json` 的 `files` 白名单内，不随包发布）。

发现日期：2026-10-04（只读核对 + 0.6.2 真机测试）

---

## P1 · confirm-dialog 测试的「机器范围」误报（测试缺陷，**不是产品缺陷**）

- **位置**：`test/confirm-dialog.test.mjs`（断言在 :46，看门狗口径在文件头 :5-7）
- **现象**：本机整轮回归稳定红 1 条（1076 / 1077），失败项 `no modal confirmation dialog appeared`
- **实测根因（2026-10-04）**：看门狗用 `EnumWindows` **全机**枚举可见窗口，只按「类名以 `Qt*` 开头」判定。
  本机**微信**（`Weixin.exe`，PID 3112 / 29388）的窗口类名是 `Qt51514QWindowIcon`，与 WPS 对话框**同类名前缀**
  （微信也是 Qt 5.15.14）。
  证据：`test/.artifacts/dialog-watch-wps.log` 显示该窗口在**看门狗启动后 25 ms**（`15:45:27.937` → `15:45:27.962`）
  就被记账 —— 早于任何场景开跑（测试自身还要先等 2500 ms 编译 `Add-Type`）。
- **结论**：不是 WPS 弹框，也不是代码缺陷。
- **影响**：任何装了 Qt 应用（微信就是）的机器上这条**永远红**，把「整轮全绿」变成常态红，训练人忽略门禁 ——
  比缺陷本身更贵。
- **建议修法**：把监视范围限定到 WPS 进程（`et` / `wps` / `wpp` 的 PID）；或在场景开跑前建立基线，**只记新增**窗口。

## P2 · 同一个工具面有两个 `schemaBytes` 读数

- **位置**：`scripts/verify.mjs:52`（逐个 `JSON.stringify(t)` 相加，**不含数组分隔符**）vs `scripts/lib/tool-face.mjs:40`（真实 `tools/list` 数组载荷）
- **现象**：`verify.mjs --static` 打印 **46,782**；`gen-numbers.mjs` / `docs/current-numbers.md` 的权威值是 **46,867**
- **差额**：85 字节 = 84 个工具之间的 **83 个逗号 + 2 个方括号**（正好 `n+1`）
- **影响**：两个门禁都过，但只有 tool-face 那个是模型**真正收到**的字节数；verify 系统性少算 `n+1` 字节。
  当前离 60,000 上限还远，无害；但贴着预算调工具面时，两个读数会给出不同答案。
- **建议修法**：verify 改成量数组载荷（或直接复用 `scripts/lib/tool-face.mjs`），让两处同源。

## P3 · `docs/HANDOFF.md` §5 自相矛盾的静态断言计数

- **位置**：:301 写静态点名 `check(` 共 **922** 处；:313 写 **924** 处
- **实测**：`924`（`test/*.test.mjs` 中 `check(` 的出现次数）
- **说明**：:313 那段正是 FIXES 91 为「修掉自相矛盾的三行」写的，但**表里的 922 漏改了**。
- **门禁盲区**：`gen-numbers.mjs --check` 的 20 项声明核对**不覆盖**这个数，所以没人会叫。
- **建议修法**：改成 924；或干脆把这个数也纳入生成器对账。

## P4 · `docs/HANDOFF.md` §9 对 confirm-dialog 失败原因的解释**写反了**

- **位置**：:510-511
- **现文**：「它等 Excel 的确认框（`Qt*` 窗口），本机不弹，于是 `no modal confirmation dialog appeared` 失败。」
- **实际**：该断言**看到 Qt 窗口才失败**（`dialogs.length === 0` 才通过）。本机是**看到了**微信的 Qt 窗口才红的。
- **建议修法**：与 P1 一起改 —— 结论（非代码缺陷）是对的，**原因写反了**。

---

以下为 **0.6.2 真机测试**（2026-10-04，走真实 MCP 工具；被测副本 `mcp/dist` 240 文件与仓库逐字节一致）。

## P5 · `wps_word_insert_text` 带 `style` 时**静默改写上一段的样式**（严重 · 静默错）

- **类别**：回 success 但结果不对（项目自定的最高危险度），且**零警告**。
- **复现（决定性，4 步）**：新建文档 → 连续插入 3 个**不带样式**的段落（甲/乙/丙）→ 再插入一段
  `wps_word_insert_text { text:"新的标题", position:"end", new_paragraph:true, style:"标题 1" }`
  - 插入前：`[4] (正文) 普通段落丙`
  - 插入后：`[4] (标题 1) 普通段落丙` ← **被改了**，而工具只报告「应用样式: 标题 1（作用在第 5 段）」
- **证据链**：工具回报「插入范围: 18-23」，而第 4 段是 `@13-19` —— 它的**段落标记正好在第 18 位**被包含进样式范围。
- **根因**：[wps-com.ps1:4953](../mcp/scripts/wps-com.ps1#L4953) 的锚点是
  `$doc.Range($doc.Content.End - 1, $doc.Content.End - 1)`（**落在最后一段的段落标记上**），
  `InsertAfter(``\r`` + 文本)` 会把 `$range` 扩到包含那个标记；而**段落标记决定整段样式**，
  于是 [wps-com.ps1:4972](../mcp/scripts/wps-com.ps1#L4972) 的 `$range.Style = $p.style` 连带把上一段也改了。
- **实际后果**：上一段变成标题 → 套上标题的字号/间距，并**混进自动生成的目录**。
  实测那次目录里就多出一条 `背景细节略。 1`（一个本该是正文的段落）。
- **影响面**：所有 `position:"end"` 的带样式插入（最常见路径）；
  `position:"start"` 也会给**新建的前导空段**套上样式（较轻，不入目录）。
- **建议修法**：样式只施加到**刚插入的那一段**（例如插入后按段落索引取
  `$doc.Paragraphs.Item($idx).Range.Style`，或把范围收窄为 `[$insertStart + 1, $insertEnd]` 以跳过前导段落标记）；
  并补一条回归：**断言插入前那一段的样式未被改动**（现有测试只断言"插入的段拿到样式"，所以漏了）。

## P6 · `wps_word_insert_text` 不翻译英文样式名（`apply_style` 会）

- **对比取证**（同一台机器、同一会话）：
  - `wps_word_insert_text { style: "Heading 2" }` → `success: true` + 警告
    「样式 'Heading 2' 未生效：对 COM 组件的调用返回了错误 HRESULT E_FAIL」，段落实际停在「正文」
  - `wps_word_apply_style { styleName: "Heading 2", range:{...} }` → **成功**，解析为「标题 2」
  - `wps_word_insert_text { style: "标题 2" }` → **成功**
- **根因**：`resolveStyleName()` 只在 [format.ts:89](../mcp/src/tools/word/format.ts#L89) 被调用；
  [content.ts:109](../mcp/src/tools/word/content.ts#L109) 把 `style` **原样**传给桥。
- **契约说明（避免误判）**：`insert_text` 的 schema 只写了中文示例（"标题 1"、"正文"），
  所以**不算违约**，属**两个工具的同名概念行为不一致** + 失败信息把原始 `HRESULT E_FAIL` 漏给用户。
- **建议修法**：`insert_text` 的 `style` 也过一遍 `resolveStyleName()`（与 `apply_style` 对齐）；
  失败文案改成"可用样式名"清单（`apply_style` 已有这套文案）。

---

## 本轮已验通过（无缺陷，备查）

- **Word 边界**（FIXES 86 的修复在 0.6.2 里生效）：越界 `replace_range` **报错且文档零改动**；
  跨段删除不带 `confirm` **被拒绝**并报出将删除的段落标记数；带 `confirm` 才执行且如实回报。
- **Excel**：写入/读回、数字类型保留、`#,##0.00`、粗体+底色、条件格式规则、列宽
  —— 全部经 **openpyxl 独立核对**，真的落进了文件；`sort_range` 正确保留表头，且**全文本无表头**时也排对；
  `write_range` 写 `"=1+1"` 是真公式、`"'=1+1"` 是文本。
- **PPT**：标题页副标题、`title_content` 正文占位符都**真的写进去了**（历史缺陷已修）；
  `export_slide_as_image` 产出真实 **1280×720 / 29,942 字节** PNG。
- **公式整块写入**：`set_formula D2:D5` 给 4 格写同一个公式（COM 语义），但**工具主动警告了**，不算缺陷。

## P7 · `insert_text` 的「所在段落」在文本重复时报错段（报告误导）

- **复现**：文档已有 `[2] 重复文本`、`[3] 中间段落`；再插入一段**与第 2 段完全相同**的文本：
  - 工具报：「插入范围: 10-15 **所在段落: 第 2 段** 应用样式: 标题 1（**作用在第 2 段**）」
  - 实际：新段是**第 4 段**；被染上样式的是**第 3、4 段**；**第 2 段根本没动**。
- **根因**：[wps-com.ps1:412](../mcp/scripts/wps-com.ps1#L412) 的 `Find-WordParagraphOfText` 用
  `$full.IndexOf($text)`，取的是**第一处**匹配；调用点只有一处（:4977，insertText）。
- **为什么要和 P5 一起修**：**任何「用这个索引去设样式」的修法都会被它带偏**
  （重复文本时会把样式打到旧的那一段上）。所以 P5 的修法**不能**依赖这个索引。
- **建议修法**：不要按文本搜。插入完成后按**字符位置**数段落标记：
  `$textStart = 新段文本起点`（`new_paragraph` 时为 `$range.Start + 1`，否则 `$range.Start`），
  `$para = 1 + ($full.Substring(0, $textStart) 中 \r / \v 的个数)` —— 与文本内容无关，天然免疫重复。

---

## P8 · `scripts/run-tests.ps1` 的 `perFile` 只留下第一个文件（**已修 · FIXES 92**）

- **发现于**：验证 P1 时跑整轮回归（2026-10-04）。
- **现象**：跑**两个以上**测试文件时，每跑完一个文件就报一次
  `Item has already been added. Key in dictionary: 'file'  Key being added: 'file'`，
  最终 `test/summary.json` 的 `perFile` **只剩第一个文件那一行**（48 个文件丢 47 行）。
- **影响有限**：`files` / `assertions` / `pass` / `fail` / `badFiles` 是独立累加的，**仍然正确**；
  文档数字（`gen-numbers`）只读那几个计数，不读 `perFile`。所以是"机器可读明细残废"，不是数字错。
- **根因（第一版猜"脚本缓存"是错的，实测纠正为编码）**：该文件含 **215 个汉字却不带 UTF-8 BOM**，
  Windows PowerShell 5.1 把无 BOM 的文件按 **ANSI/GBK** 读 → 中文注释把紧随其后的那条语句
  `$summaryRows = @()` **整个吞掉** → 第一次 `$null += [ordered]@{...}` 让变量变成**字典** →
  第二次 `+=` 对字典加同一个键就抛那个错。三种报错形态（键重复 / ArrayList 版的 null 值 / 报错行号对不上）全部吻合。
  判据：`scripts/build-host-actions.ps1` 有 0 个非 ASCII 字符，同样不带 BOM 却一切正常。
- **修法（已落地）**：给 `scripts/run-tests.ps1` 补 **UTF-8 BOM**，行尾**保持 LF**（`scripts/*.ps1` 本来就是 LF，
  且 `.gitattributes` 写着 `* -text`「never rewrite line endings」），逐字符比对确认**内容一字未改** ——
  `git diff` 只有 1 行。并把它加进 `scripts/lint.mjs` 的字节约定清单防回归（该清单改成可逐文件声明"要不要 CRLF"）——
  这个 bug 能活下来，正是因为那份清单只列了 `mcp/scripts/wps-com.ps1` 与 `host/wps-com-host.ps1`。
- **验证**：`test/summary.json` 的 `perFile` 从 1 行恢复成 **48 行**，报错消失，整轮仍 `PASS=1085 FAIL=0`。

## ⚠️ 操作陷阱（本轮踩到，务必记牢）

**不要对"含超长行"的文件做"读入 → 拼接 → 写回"。** 本轮用这个方式修 lint 时，
把 `test/confirm-dialog.test.mjs` 第 20 行（一个内嵌 PowerShell 的 JS 字符串，**1573 字符**）截断成 2000 字符上限内的残片，
文件语法直接坏掉 —— 而它**只在整轮回归里**才暴露（`PASS=0 FAIL=0 exit=1`，单独跑反而看不出来）。
- **判据**：`node --check <file>`；以及"最长行 > 2000 字符"就**绝对不要**读改写。
- **正确做法**：用`edit`做**定点替换**（`old_string` 只取短行），不要整文件重写。

**同类陷阱（2026-10-04 第二次踩到，同样危险）：`read` 默认只返回 2000 行。**
对超过 2000 行的文件（如 `docs/FIXES.md`，3024 行）做"读入 → 拼接 → 写回"，会把文件**截断到最后 2000 行** ——
一次丢掉 1000 多行历史。当天写 FIXES 93 时就差点这么干，**是代码里的锚点断言先抛错才拦住**。
- **判据**：`read` 返回的 `lines.length` 与 `totalLines` 是否相等；不等就**绝对不能**用它的内容回写。
- **正确做法**：**一律用 `edit` 定点插入/替换**；只有新建文件才用 `write`。

## 修复方案汇总（建议，待用户拍板后执行）

**改动分三组，互不冲突：**

### 第一组 · 桥（`mcp/scripts/wps-com.ps1`）—— P5 + P7

改 `insertText`（:4934-4982）两处：

1. **样式只许落在刚插入的那一段**（P5）。`$range` 的起点是**上一段的段落标记**，
   直接 `$range.Style` 会连上一段一起染（段落标记决定整段样式）。
   在 `$p.new_paragraph` 且走了 ``\r`` 前缀分支时，把样式范围收窄一格：
   `$styleRange = $doc.Range([int]$range.Start + 1, [int]$range.End)`（start / end 两个分支同理），
   其余情况保持 `$range`。
2. **段落索引改按位置算**（P7），替换 `Find-WordParagraphOfText $doc ([string]$p.text)`。

> ⚠️ 桥是**真源**，改完必须重生成 `host/wps-actions.ps1`（`scripts/build-host-actions.ps1`），
> 且桥本身是**纯 CRLF、无 BOM** —— 不要用 `Get-Content -Raw` + `Set-Content` 改（会把中文写坏）。

### 第二组 · TS 工具层（`mcp/src/tools/word/content.ts`）—— P6

`import { resolveStyleName } from './style-names';` 后把 `style` 过一遍翻译再喂桥，
文案里同时给出「原输入」与「实际样式名」，与 `apply_style` 对齐。

### 第三组 · 测试与文档 —— P1 / P2 / P3 / P4

- P1：看门狗加**基线**（开跑前记下已存在的窗口句柄，只报**新增**的），并记下窗口所属进程便于事后判定。
  不建议只按进程名过滤（宁可多记不可漏记，避免把真弹框过滤掉导致门禁失效）。
- P2：`scripts/verify.mjs:51-52` 改成量数组载荷 `Buffer.byteLength(JSON.stringify(tools))`，与 tool-face 同源。
- P3：`docs/HANDOFF.md:301` 的 922 → **建议直接删掉这个数字**（它已漂过 782/918/928/922 四次），
  只保留「差额来自循环内断言」的说法；保守做法是改成 924。
- P4：`docs/HANDOFF.md:510-511` 那句因果写反的解释按 P1 的结论重写。

**修完必跑**（项目纪律，顺序不能颠倒）：
`build-host-actions.ps1` → `cd mcp; npx tsc` → `extract-spec.mjs` → `npx tsc` → `gen-tool-surface.mjs`
→ `gen-skill-tools.mjs` → `gen-tool-coverage.mjs` → `gen-numbers.mjs`，再跑
`lint` / `verify --static` / `param-contract` / `verify:package`。

**必须新增的回归断言**（现有测试只断言「插入的那段拿到了样式」，所以 P5 才漏得过去）：
1. 插入带样式的新段后，**断言此前已存在的每一段样式都未变**；
2. 断言新段**确实**拿到了样式；
3. 重复文本场景：断言被套样式的是**最后一段**、旧的那段未被改动，且回报的段落号与真实一致（覆盖 P7）。



---
name: wps-word
description: 用 WPS 文字工具读取与插入正文、按字符范围或段落样式改格式、页面设置与分节、生成目录与页码、填模板字段、批注与修订。
whenToUse: 任务针对文档、段落、标题、样式、字符格式、页眉页脚、目录、模板字段、批注修订或文档转换时用本技能。
---

# WPS 文字

共享的执行协议（目标确认、保存策略、验证与失败分流）见 `wps-office-next` 技能；
本节是文字特有的决策与流程。**应用技能可能被单独加载**，所以关键规则在本文件里也写了一遍。

<!-- GENERATED:advertised:start -->
本应用注册 **75** 个工具，standard 档直接广告 **30** 个；
完整工具表与按任务分组索引见同目录 `reference.md`（未广告的用 `wps_call`）。
<!-- GENERATED:advertised:end -->

## 1. 先分两件事：内容编辑 / 版式编辑

它们的作用范围与失败方式都不一样，混在一起做最容易出"改了别处"的错：

- **内容编辑**：插入、替换、删除、填表、模板字段。
- **版式编辑**：段落样式、字符格式、页面与分节、页眉页脚与页码、目录。

建议顺序：**先把内容定稿，再做版式**；或者明确知道自己在做哪一层时再动。改版式前先读结构，否则你不知道哪一段会被波及。

## 2. 三层格式，作用范围完全不同

| 层 | 工具 | 作用范围 |
|---|---|---|
| 字符格式 | `wps_word_set_font`（`fontName`/`fontSize`/`bold`/`color`…） | 精确落在 `range:{start,end}` 的字符上，或 `all` = 整篇 |
| 段落样式 | `wps_word_apply_style`（`styleName`） | **段落样式**：作用于与范围**相交的整段**，`range` 只能决定"从哪一段开始" |
| 节页面设置 | `wps_word_get_page_setup` / `wps_word_set_page_setup`（`section`） | 某一节；方向与纸张先应用、边距后应用 |

要点：

- `apply_style` **不是**精确字符格式工具。它会回报真实受影响的段落，连带染了多段时也会给 `warning` —— 看到就说明范围算偏了。
- `set_font` 只接受 `range:{start,end}` 或 `"all"`；**不传就是当前选区**，光标折叠时是空范围，什么都不会发生（工具会如实回报"0 个字符"并给 warning）。
- **不要先设标题再用 `all` 统一正文字体**：`all` 会连标题一起覆盖。需要全局基线时按下面的顺序做。

## 3. 字符坐标：来源、单位与失效

- **来源**：`wps_word_get_paragraphs` 每段返回 `{index, text, style, start, end}`。
- **单位**：Word 的 Range 字符偏移，**0 基**，`end` **不含**（即 `[start, end)`）。`index` 是 **1 基**的段落号。`text` 超过 200 字符会被截断——坐标仍可用于定位，但别把 `text` 当全文。
- **失效**：任何插入、删除、替换之后，**其后所有旧坐标都会平移**。旧坐标不可长期复用。
- **优先重新定位**：改完一处就读一次 `get_paragraphs`，再算下一处。
- 确实要按一批原坐标做替换时，只有**先证明这些范围互不重叠且顺序合适**，才考虑从后往前执行；否则重新定位。
- 改完检查**前后相邻段落**，确认格式或删除没有越界。
- 越界坐标会直接报错（`set_font` 要求 `0 ≤ start ≤ end ≤ 文档长度`），错误里会给出实际文档长度。

**不要用通用换行符切分代替 Word 的段落模型。** `get_document_text` 里段落之间显示为 `\v`；`\r` 是段落标记、`\v` 是手动换行/段落显示符、表格单元格结束标记又是另一种。按换行切分时不要假设它们语义相同，需要严格分段就 `get_paragraphs`。

## 4. 常用流程：改已有文档

1. 确认目标：`wps_word_get_active_document`（或先 `wps_word_get_open_documents` 看有几份打开）。
2. 读结构：`wps_word_get_paragraphs`（分段落看）或 `wps_word_get_document_text`（看全文）。
3. 改内容：`wps_word_insert_text` / `wps_word_replace_range` / `wps_word_find_replace`。
4. 改版式：按 §2 的三层选工具，范围要窄。
5. 回读核对：再取一次 `get_paragraphs` / `get_document_text`，必要时 `get_page_setup` 对账。
6. 按用户意图保存：`wps_common_save` 覆盖原文件，或 `wps_common_save_as` 另存副本。**用户没要求关闭就不要关闭。**

从零起草：`wps_word_create_document` → `wps_word_insert_text` 写内容 → `wps_common_save_as` 落盘（新文档没有路径时，工具不会替你猜，也不会丢弃内容）。

## 5. 排版配方

```
读标题与段落结构 → 确定样式映射 → 需要全局基线就先用 all 应用
→ 再对标题与特殊段落做局部覆盖 → 按节设纸张与边距 → 更新目录与页码
→ 回读核对（必要时导出检查）
```

- **顺序是「先全局、后局部」**：先 `set_font {range:"all"}` 定正文字体基线，再用 `get_paragraphs` 拿到的坐标对标题行单独设字体或 `apply_style` 覆盖。反过来就会把标题一起改掉。
- 只有**确定正文范围**时才做"统一正文字体"；否则按段落坐标一段段来。
- 样式名带空格：`标题 1`、`标题 2`，不是 `标题1`。
- 页面设置：`set_page_setup` 的边距单位是**磅**（整数，0-1584），键名 `topMargin/bottomMargin/leftMargin/rightMargin`，`orientation` 取 `portrait`/`landscape`。回报里始终给出四个边距 + 方向 + 纸张，可以据此核对"我要的上边距有没有跑到别的地方"。
- 多节文档（横竖混排）**必须指定 `section`** 并逐节核对；只改目标节，其他节不受影响。
- 页眉页脚、页码与"链接前节"要**按节**核对：`wps_word_get_header_footer {section}` 读该节的文本与域清单，用它确认页码到底进文件没有。

## 6. 目录与页码

- `wps_word_generate_toc {action:"insert"}` 新插一份目录（默认）；`{action:"update"}` **只刷新已有目录**，文档里根本没有目录时会明确报错并提示改用 `insert`。
- **不能只凭模板名称或经验判断"应该已经有目录"** —— 先看文档里是否真存在（更新失败就是没有）。
- `wps_word_insert_page_numbers` 的 `includeTotal:true` 会在页码域后再插一个 NUMPAGES 域，形成「第 X 页 / 共 Y 页」；返回里 `pageNumbers` 是页码域数、`totalPagesFields` 是总页数域个数。总页数插不进去会明确报错，此时页码已经插上，属**部分成功**。
- 这条路径需要 WPS 有活动窗口；没有窗口时会明确失败，而不是写错地方。

## 7. 模板配方

1. **先确认模板里的字段类型**：书签、内容控件、域、还是纯文本占位符。
2. 优先用已确认的字段：`wps_word_get_bookmarks` / `wps_word_get_content_controls` 读现状，再用对应的写工具。
3. 纯文本占位符才用 `wps_word_smart_fill_field {keyword, value, fillMode}`。**不要笼统认为它能保留所有格式** —— 它的匹配、替换与格式行为要按实际结果核对；也不要用 `find_replace` 去替字段名（会删掉关键字并破坏格式）。
4. 逐份生成时逐项核对：模板原件**未被修改**、输出路径唯一、字段没有遗漏、输入行与输出文件**一一对应**。
5. 失败时**只重试未完成的输出**，不要重跑已经验证过的文件（重复生成会覆盖或产生重复文件）。
6. 批量分份产出用 `wps_word_mail_merge`（数据文件是 UTF-8 CSV，第一行字段名）比手写循环更省事；它默认合并成一份，给了 `splitDir` 才每行一份。生成后仍要逐份抽查字段是否落对。

## 8. 已知限制与坑

- **破坏性范围替换**：`wps_word_replace_range` 越界直接报错；一次删掉超过 1 个段落标记必须显式传 `confirm: true`，否则拒绝执行。真要用它删正文前，先 `get_paragraphs` 核对坐标。
- `find_replace`：`findText` 必填；**不传 `replaceText` 就是纯查找**（只统计命中数、不改文档）。传了才替换，`replaceAll:false` 时只替换第一处。
- 多份文档同时打开时，读写正文都落在**活动文档**上，而活动文档跟窗口焦点走；Word 侧的歧义**不会有**自动提醒。动手前 `wps_word_get_active_document` 确认目标。
- 转换（`wps_convert_to_pdf` / `wps_convert_format`）**必须显式传 `appType:"word"`**，否则可能取到正在运行的 Excel。
- 不要向用户承诺可以撤销：COM 改动不一定进撤销栈。
- 修订与批注：`wps_word_get_revisions` 先看改了什么，`accept_revisions` / `reject_revisions` 是不可逆的，动手前说清范围。

## 9. 按需参考

| 想了解 | 读哪里 |
|---|---|
| 全部文字工具与按任务分组索引 | 同目录 `reference.md` |
| 范围 / 样式 / 分节的实测细则与坐标规则 | `references/ranges-styles-sections.md` |
| 排版已有报告的完整任务实例 | `recipes/format-existing-report.md` |
| 模板批量填写的完整任务实例 | `recipes/fill-template.md` |
| 共享执行协议 | `wps-office-next` 技能 |

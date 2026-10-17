# 文字：范围、样式、分节与页眉页脚的实测细则

只在你需要**按坐标改格式**、**处理多节文档**或**判断字段类型**时读本文件。

## 1. 字符坐标

| 项 | 值 |
|---|---|
| 来源 | `wps_word_get_paragraphs` 每段返回 `{index, text, style, start, end}` |
| 单位 | Word `Range.Start/End` 的字符偏移 |
| 基准 | **0 基**，`end` **不含**（区间是 `[start, end)`） |
| `index` | **1 基**的段落号 |
| `text` | 尾部 `\r`/`\n` 已去掉；超过 **200** 字符会截断（坐标仍准确） |

边界校验（`set_font`）：`0 ≤ start ≤ end ≤ 文档长度`；越界会直接报错并给出实际文档长度，可以用错误里的长度当参照。

**失效规则**：任何插入、删除、替换之后，其后所有坐标整体平移。旧坐标不要跨编辑复用。

**倒序执行的唯一前提**：先证明这些范围**互不重叠**且顺序合适（后面的改动不影响前面的坐标）。证明不了就一处一处重新定位。

## 2. 段落模型：`\r` / `\v` / 表格单元格标记

`wps_word_get_document_text` 里段落之间显示为 `\v`。它们不是同一种东西：

| 标记 | 含义 |
|---|---|
| `\r`（`[char]13`） | Word 的**段落标记**（真正的段落边界） |
| `\v`（`[char]11`） | 手动换行 / 段落显示符；`get_document_text` 用它展示段落分隔 |
| 表格行/单元格结束标记 | 单元格结构标记，不是正文段落 |

**不要用通用换行符切分来代替 Word 的段落模型。** 需要严格按段落处理就用 `get_paragraphs`；需要按行处理就把两个换行符都算上，并意识到表格里的内容也会被计入。

`insert_text` 的 `new_paragraph:true` 会在文本前插一个段落标记（这样文本独立成段），随后插入的文本范围**不含**那个标记——所以给新插入的段落套样式时会精确落在文本上，不会染到上一段。

## 3. `set_font` vs `apply_style`

| | `wps_word_set_font` | `wps_word_apply_style` |
|---|---|---|
| 性质 | **字符格式** | **段落样式** |
| 作用范围 | 精确落在 `range:{start,end}` 的字符上；`"all"` = 整篇 | 与范围**相交的整段**；`range` 只决定"从哪一段开始" |
| 不传 range | 当前选区（光标折叠 = 空范围 = 什么都不做，会回报"0 个字符" + warning） | 同左 |
| 回报 | 写入后读回，对不上走 warning | 回报真实受影响的段落；连带多段时有 warning |

实践含义：**要"改这几个字的字体"用 `set_font`；要"这一段是标题"用 `apply_style`。** 混用会得到"范围算偏了"的结果。

**全局基线 + 局部覆盖的顺序**：先 `set_font {range:"all"}` 定正文基线，再用段落坐标对标题/特殊段落 `apply_style` 或 `set_font` 覆盖。反过来（先标题再 all）会把标题一起改掉。

## 4. 分节与页面设置

- 一个文档可以有多个**节**，每节有自己的页面方向、纸张与边距。
- `wps_word_get_page_setup {section}` / `wps_word_set_page_setup {section, ...}`；不传 `section` 用第 1 节。
- 边距单位是**磅**（整数，0-1584），键名 `topMargin/bottomMargin/leftMargin/rightMargin`；`orientation` 取 `portrait`/`landscape`。
- 工具内部**固定先应用方向与纸张、后应用边距**（改方向会把已有边距整体旋转），回报里始终给出四个边距 + 方向 + 纸张，可以据此核对。
- **横向页的「纸张」可能是一个未命名代码，这不是错误**：实测 WPS 12.1 把「A4 横向」记成 `41（横向纸张代码）`，改回纵向又恢复 `A4`；**即使同时传 `paperSize:"A4"` + `orientation:"landscape"` 也一样**。
  遇到这种代码以**页面尺寸**（`pageWidth`/`pageHeight`，磅）为准判断纸张是否正确 —— 别因为它反复改纸张、反复核对，那会把一次正常的"改方向"拖成十几步。
- 横竖混排：改目标节之前先逐节 `get_page_setup` 读出基线，改完只回读目标节，并抽查相邻节没被改动。

节数怎么探：逐节调 `wps_word_get_page_setup {section:N}`，从 1 开始直到报错。

## 5. 页眉页脚、页码与"链接前节"

- `wps_word_get_header_footer {section}` 读该节的页眉页脚：文本（尾部段落标记已去掉）+ 域清单（`PAGE` / `NUMPAGES` / `OTHER`）。**核对"页码到底进文件没有"就靠它。**
- `wps_word_insert_header` / `insert_footer` 用 `{text, section}`；节号不存在会明确报错。
- `wps_word_insert_page_numbers` 的 `includeTotal:true` 会在页码域后再插一个 NUMPAGES 域。返回 `pageNumbers`（页码域数）与 `totalPagesFields`（总页数域数）。总页数插不进去会明确报错，此时页码已插上，按**部分成功**报告。
- 这条路径需要 WPS 有活动窗口；没有窗口会明确失败。
- 后续节可能"链接到前一节"：**逐节核对**，不要假设只有第 1 节有页眉页脚。

## 6. 目录

- `wps_word_generate_toc {action:"insert"}` 新插一份（默认）；`{action:"update"}` 只刷新**已有**目录。
- 文档里没有目录时 `update` 会明确报错并提示改用 `insert`——**这就是"有没有目录"的判据**，不要靠模板名字或记忆猜。
- `position` 取 `start` / `cursor`；`levels` 默认 3；`includePageNumbers` 默认 true。
- 目录内容来自**标题样式**：先用 `apply_style` 把标题段落套上 `标题 1/2/3`，再生成目录。套错样式会得到空目录（工具会以"调用后目录数仍为 0"明确失败）。

## 7. 模板字段类型

先读再写：

| 字段类型 | 读 | 写 |
|---|---|---|
| 书签 | `wps_word_get_bookmarks` | `wps_word_replace_bookmark_content`、`wps_word_insert_bookmark` |
| 内容控件 | `wps_word_get_content_controls` | `wps_word_add_content_control` |
| 域 | `wps_word_get_header_footer`（只覆盖页眉页脚域） | 对应工具 |
| 纯文本占位符 | 读全文 / `find_in_document` | `wps_word_smart_fill_field {keyword, value, fillMode}` |

`smart_fill_field` 的 `fillMode` 取 `auto/underline/afterColon/afterLabel/placeholder`。**它的匹配与替换行为要按实际结果核对**：不要笼统承诺"能保留所有格式"；也不要用 `find_replace` 去替字段名（会删掉关键字并破坏格式）。

## 8. 修订与批注

- `wps_word_get_revisions` 先看改了什么，并报告"修订跟踪"是否打开。
- `accept_revisions` / `reject_revisions` **不可逆**：动手前说明范围。
- `enable_track_changes` / `get_track_changes_status` 控制与查询跟踪状态。
- 批注：`insert_comment` / `get_comments` / `delete_comment`。

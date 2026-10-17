# 配方：给已有报告排版

适用：一份内容已经写完的 Word 报告，要统一字体、分层标题、加目录与页码、设置分节页面。
目标：**改版式，不改内容**；改完能逐项回读证明。

## 0. 先读基线（不要跳过）

```
wps_word_get_active_document          # 核对 name / path：确认改的是哪一份
wps_word_get_document_stats           # 页数 / 字数 / 段落数：作为"内容没被改"的对照
wps_word_get_paragraphs {startParagraph:1, endParagraph:50}
```

记下：段落总数、每个标题段落的 `{index, text, style, start, end}`、各节的页面设置（逐节 `wps_word_get_page_setup {section:N}`）。

> 内容做完了再排版。**先内容、后版式**——反过来每次改内容都要重做一遍坐标。

## 1. 确定样式映射

先把报告里的段落分成几类，各自对应到哪个样式或字符格式：

| 类别 | 处理 |
|---|---|
| 一级 / 二级 / 三级标题 | `apply_style {styleName:"标题 1"/"标题 2"/"标题 3", range:{start,end}}` |
| 正文 | 字符格式（`set_font`），不是段落样式 |
| 特殊段（引文、图注、表格标题） | 单独列出，逐个覆盖 |
| 页眉页脚 | 见第 4 步 |

样式名带空格：`标题 1`、`标题 2`，不是 `标题1`。

## 2. 先全局基线，再局部覆盖

顺序是**先全局、后局部**：

```
wps_word_set_font {fontName:"宋体", fontSize:12, range:"all"}      # 正文字体基线
wps_word_apply_style {styleName:"标题 1", range:{start:120, end:135}}   # 再逐段覆盖标题
```

反了就会把标题一起改掉。**只有确定正文范围时才做"统一正文字体"**；否则按 `get_paragraphs` 给出的坐标一段段来。

每改完一段就读一次 `get_paragraphs` 重新取坐标（编辑会让后面的坐标平移）。`apply_style` 回报的真实受影响段落数如果大于 1，说明范围算偏了，回上一步重来。

## 3. 目录与页码

```
wps_word_generate_toc {action:"update"}     # 已有目录：刷新
wps_word_generate_toc {action:"insert"}     # 没有目录：新插一份（update 报错就说明没有）
wps_word_insert_page_numbers {section:1, position:"footer", alignment:"right", includeTotal:true}
wps_word_get_header_footer {section:1}      # 回读：页码域 / 总页数域到底有没有进去
```

`update` 报错 = 文档里没有目录 → 改用 `insert`。这就是"有没有目录"的判据。

## 4. 分节页面（横竖混排）

```
wps_word_get_page_setup {section:1}                          # 逐节读基线
wps_word_set_page_setup {section:2, orientation:"landscape", paperSize:"A4"}
wps_word_get_page_setup {section:2}                          # 只回读目标节
```

- 边距单位**磅**（整数，0-1584）。工具固定"先方向纸张、后边距"，所以改完方向要**回读边距**确认没有跑位。
- 改完抽查**相邻节**没有变化；页眉页脚按节核对（可能有"链接前节"）。

## 5. 核对

| 核对项 | 判据 |
|---|---|
| 内容没变 | 段落数、字数与第 0 步记录一致（排版不该改内容） |
| 标题层级 | 按序回读每段的 `style`，与第 1 步的映射逐项对上 |
| 正文字体 | 抽查若干正文段的 `fontName` / `fontSize` |
| 标题字体 | 抽查标题段**没有**被基线覆盖 |
| 目录 | 目录存在、条目数与标题数一致；刷新后条目没有重复 |
| 页码 | `get_header_footer` 读到 PAGE（与 NUMPAGES，如果要求了） |
| 分节 | 目标节正确、相邻节未变 |

## 6. 保存与交付

- 覆盖原件：`wps_common_save`。
- 保留原件：先 `wps_common_save_as {outputPath}` 建副本，**确认后续改动落在副本上**，再排版。
- 用户没要求关闭就**不要关闭**。
- 能导出就 `wps_convert_to_pdf {appType:"word", outputPath:"..."}`；**只有真的看过渲染结果**才在报告里说"版式正常"，否则写"内容与属性已核对，视觉效果未验证"。

## 7. 验收标准

- [ ] 内容（段落数/字数）与排版前一致。
- [ ] 标题层级、正文字体、页面设置都有回读证据。
- [ ] 目录既不缺也不重复。
- [ ] 页码域真实存在（不是"插了但没进去"）。
- [ ] 目标文档身份在开始与结束时一致（没有改到另一份打开的文档）。
- [ ] 未验证项（视觉渲染）明确标注。

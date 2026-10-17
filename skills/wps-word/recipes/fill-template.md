# 配方：模板批量填写

适用：一份模板（带书签 / 内容控件 / 占位符）+ 一张数据表 → 逐行生成 N 份文档（或一份合订）。
目标：**模板原件不动、输出一一对应、字段不遗漏、失败只补做未完成的部分**。

## 1. 先确认字段类型（决定用哪个工具）

```
wps_word_open_document {path:"...\\模板.docx"}
wps_word_get_bookmarks            # 书签列表
wps_word_get_content_controls     # 内容控件列表
wps_word_get_document_text        # 看纯文本占位符长什么样
wps_word_get_paragraphs           # 占位符所在的段落坐标
```

| 模板里是什么 | 用什么写 |
|---|---|
| 书签 | `wps_word_replace_bookmark_content {…}` |
| 内容控件 | 对应写工具；`get_content_controls` 先确认可写字段 |
| 纯文本占位符（如 `{姓名}`） | `wps_word_smart_fill_field {keyword, value, fillMode}` |
| 表格单元格 | `wps_word_set_table_cell {table, row, column, text}` |

**不要**用 `find_replace` 替字段名：它会删掉关键字，并把占位符周边的格式一起带走。

`smart_fill_field` 的 `fillMode` 取 `auto/underline/afterColon/afterLabel/placeholder`。
**先用一行数据试填一次并回读**，确认它真的落在该落的位置、格式没坏，再批量。不要假设"它能保留所有格式"。

## 2. 两条批量路线

### 路线 A：`wps_word_mail_merge`（推荐，一步一份或一份合订）

```
wps_word_mail_merge {dataFile:"...\\名单.csv", splitDir:"...\\产出",
                     fileNameField:"通知书_{部门}_{姓名}_{工号}",
                     fields:["姓名","部门","工号"]}
```

- 数据文件是 **UTF-8 CSV**，第一行是字段名。
- 不传 `splitDir` → 合并成**一份**新文档（父文档不动）；传了 → **每行一份**存进该目录。
- 返回后核对：产出文件数 = 数据行数。

### 路线 B：逐行手工填（字段逻辑复杂、需要自定义命名时）

对每一行：

```
wps_word_open_document {path:"...\\模板.docx"}
wps_word_get_active_document                      # 核对打开的是模板副本
wps_word_smart_fill_field {keyword:"姓名", value:"张三"}
wps_common_save_as {outputPath:"...\\产出\\通知书_张三.docx"}
wps_word_get_document_text                        # 回读：字段都落上了吗
wps_word_close_document {save:false}              # 关掉这一份
```

**要点**：
- 每份都从**模板复制出来的新文档**开始，**绝不**在模板原件上直接 `save`。
- 先 `save_as` 到唯一路径，再写入；或先写入再 `save_as`——但一定确认保存的是**这一份**（用 `get_active_document` 的 `path` 核对）。
- 产出的命名要能反推输入行（用 `{部门}_{姓名}_{工号}` 这类模板）——否则失败重试时对不上账。

## 3. 失败处理

- **只重试未完成的输出。** 先列产出目录，与输入行逐行比对，找出缺哪几份；不要整体重跑（会覆盖已验证文件或产生重复）。
- 单行失败时记录：输入行的主键、失败的动作、原始错误。已生成的文件**不要删**（除非用户要求）。
- 文件数对不上时不要报"全部完成"。

## 4. 核对与验收

| 核对项 | 判据 |
|---|---|
| 模板原件 | 内容与开始时一致（字段仍是占位符）；mtime/内容抽查 |
| 输出数量 | 文件数 = 有效输入行数 |
| 一一对应 | 文件名 → 输入行主键能唯一对应；随机抽 3~5 份读回字段值与 CSV 一致 |
| 字段完整 | 没有残留 `{占位符}`；该有的字段都有值（空值也要是"空"而不是"没填"） |
| 格式 | 抽查字体/段落样式没被替换操作破坏 |
| 落盘 | 每份都独立可打开；PDF 与 docx 成对（如果要求了 PDF） |

## 5. 验收标准

- [ ] 模板原件未被修改。
- [ ] 输入行与输出文件一一对应，数量相等。
- [ ] 无遗留占位符；抽查值与来源一致。
- [ ] 失败行被单独列出并可重试，没有重跑已验证文件。
- [ ] 未验证项（未抽查到的文件、视觉渲染）明确标注。

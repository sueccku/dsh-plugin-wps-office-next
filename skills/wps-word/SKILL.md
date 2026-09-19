---
name: wps-word
description: 用 WPS 文字工具读取与插入正文、查找替换、设置字体与样式、页面设置、生成目录、填模板字段。
whenToUse: 任务针对文档、段落、标题、样式、批注、页眉页脚、目录或文档转换时。
---

# WPS 文字

## 直接广告的工具

wps_word_get_active_document、wps_word_get_document_text、wps_word_get_paragraphs、wps_word_insert_text、wps_word_find_replace、wps_word_apply_style、wps_word_set_font、wps_word_set_paragraph、wps_word_generate_toc、wps_word_smart_fill_field、wps_word_open_document、wps_word_create_document，以及通用的 wps_common_save、wps_convert_to_pdf。

其余工具（表格、图片、书签、批注、页眉页脚、分节符、行距、页面设置、修订与校对、关闭文档）通过 wps_call 使用，清单见同目录 reference.md，也可以直接 `wps_help {query:"关闭文档"}` 搜。

## 参数约定

- find_replace：find_text 必填；**不传 replace_text 就是纯查找**——只统计命中次数、不改动文档。传了 replace_text 才会替换，replace_all=false 时只替换第一处。
- insert_header / insert_footer 用 {text, section}，section 从 1 开始；节号不存在会明确报错。
- set_page_setup 的页边距单位是**磅**（整数，0-1584），键名是 marginTop / marginBottom / marginLeft / marginRight，orientation 取 portrait / landscape。成功后会回报实际生效值。
- insert_image 用 imagePath（不是 path）。
- set_line_spacing 用 lineSpacing（倍数）；set_paragraph 可同时设 alignment 与 lineSpacing。
- 从零起草（不是编辑现有文档）时先 `wps_word_create_document`，它会回报新文档名，再 insert_text 写内容，最后 wps_common_save_as 落盘。
- 关闭文档用 `wps_word_close_document`（save 默认 true）；它不在广告位，用 `wps_call {tool:"wps_word_close_document", args:{save:false}}`。对从未落盘的文档会自动改为不保存关闭并回报 warning，不会弹保存对话框。
- 多份文档同时打开时，读写正文的动作都作用在**活动文档**上，而活动文档跟着窗口焦点走；Word 侧本轮没有加自动提醒，动手前先 `wps_word_get_active_document` 确认目标，或先把多余文档 `close_document` 收掉。

## 常用流程

1. 先确定编辑目标：改现有文档用 wps_word_get_active_document 确认，再 get_document_text 或 get_paragraphs 读现状；从零起草用 wps_word_create_document 新建。
2. 建立大纲：先定位标题文本，用 apply_style 应用带空格的样式名（标题 1、标题 2），最后 generate_toc 生成目录。
3. 模板填写优先用 smart_fill_field，不要用 find_replace 替换字段名，后者会删掉关键字并破坏格式。
4. 改完用 get_paragraphs 或 get_document_text 复验，必要时 set_page_setup 回报的值也一并核对，再 wps_common_save。
5. 收尾：要归档到新路径用 wps_common_save_as；要关掉文档用 wps_word_close_document，别把没保存的文档留在用户机器上。

## 已知坑

- set_font 传 range 为 all 会连同标题字体一起改掉；正确顺序是先把标题应用样式，再统一正文字体，最后重新 apply_style 恢复标题。
- 样式名带空格，例如 标题 1，不是 标题1。
- 纯查找与替换是两个行为；用 find_replace 做"检查是否出现"时不要顺手传空的 replace_text，否则会被当成替换。
- 不要向用户承诺可以撤销：COM 改动不一定进撤销栈，重要文档先保存副本。

---
name: wps-word
description: 用 WPS 文字工具读取与插入正文、查找替换、设置字体与样式、生成目录、填模板字段。
whenToUse: 任务针对文档、段落、标题、样式、批注、页眉页脚、目录或文档转换时。
---

# WPS 文字

## 直接广告的工具

wps_word_get_active_document、wps_word_get_document_text、wps_word_insert_text、wps_word_find_replace、wps_word_apply_style、wps_word_set_font、wps_word_generate_toc，以及通用的 wps_common_save、wps_convert_to_pdf。

其余工具（段落读取、表格、图片、书签、批注、页眉页脚、分节符、行距、页面设置、模板智能填写、修订与校对）通过 wps_call 使用，清单见同目录 reference.md。

## 常用流程

1. wps_word_get_active_document 确认文档，再 get_document_text 或 get_paragraphs 读取现状。
2. 建立大纲：先定位标题文本，用 apply_style 应用带空格的样式名（标题 1、标题 2），最后 generate_toc 生成目录。
3. 模板填写优先用 smart_fill_field，不要用 find_replace 替换字段名，后者会删掉关键字并破坏格式。
4. 改完用 get_paragraphs 或 get_document_text 复验，再 wps_common_save。

## 已知坑

- set_line_spacing 在底层没有对应实现，会返回 Unknown action；改用 set_paragraph 设置行距。
- set_page_setup 的参数处理有缺陷，可能返回 undefined；改动页面设置后必须复验。
- set_font 传 range 为 all 会连同标题字体一起改掉；正确顺序是先把标题应用样式，再统一正文字体，最后重新 apply_style 恢复标题。
- 样式名带空格，例如 标题 1，不是 标题1。
- 不要向用户承诺可以撤销：COM 改动不一定进撤销栈，重要文档先保存副本。

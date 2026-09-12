---
name: wps-ppt
description: 用 WPS 演示工具增删幻灯片、设置标题与正文、添加文本框与形状、插入图片、导出幻灯片为图片。
whenToUse: 任务针对演示文稿、幻灯片、形状、文本框、动画、切换、母版或演示导出时。
---

# WPS 演示

## 直接广告的工具

wps_ppt_get_slide_count、wps_ppt_get_slide_info、wps_ppt_get_shapes、wps_ppt_add_slide、wps_ppt_set_slide_title、wps_ppt_set_slide_content、wps_ppt_add_textbox、wps_ppt_insert_ppt_image、wps_ppt_export_slide_as_image，以及通用的 wps_common_save、wps_convert_to_pdf。

演示是三个应用里工具最多的一个（115 个）。动画、切换、母版、配色、图表、流程图、KPI 卡片、时间线、3D 文字等全部通过 wps_call 使用，清单见同目录 reference.md。

## 参数约定

- 新增形状与文本框用 left、top、width、height，单位像素；没有 x、y 参数，传了会被忽略。
- 插入图片用 path（不是 imagePath），单位磅。
- 设置动画必须给 shapeIndex，先用 get_shapes 拿到形状索引；动画枚举是 fadeIn、flyIn、wipeIn、zoomIn、bounceIn、spinIn、fadeOut、flyOut。
- 设置切换的参数名是 transition（不是 transitionType），且没有 duration 参数。

## 常用流程

1. wps_ppt_get_slide_count 与 get_slide_info 确认文稿与页码。
2. 多份演示文稿同时打开时，先用 set_active_target 锁定目标文稿（通过 wps_call），否则 ActivePresentation 会漂移。
3. 加页用 add_slide，然后 set_slide_title 与 set_slide_content 填内容。
4. 需要精细排版时先 get_shapes 取形状索引，再对具体形状操作。
5. 导出前先 wps_common_save，导出用 export_slide_as_image（绝对路径）。

## 已知坑

- 示例里常见 x、y、shapeName、transitionType、imagePath 这类参数名，在本工具集里全部无效；只有参数名完全正确才会生效，写错名字不会报错。
- 配色相关工具有两套枚举（beautify 的内置配色与 apply_color_scheme 的 ocean/forest/sunset 等），不要混用。
- 定位到不存在的目标文稿时，部分底层实现会继续解引用空对象；锁定目标后要复验。

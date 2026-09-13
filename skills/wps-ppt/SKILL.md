---
name: wps-ppt
description: 用 WPS 演示工具增删幻灯片、设置标题与正文、添加文本框与形状、插入图片、动画与切换、导出幻灯片为图片。
whenToUse: 任务针对演示文稿、幻灯片、形状、文本框、动画、切换、母版或演示导出时。
---

# WPS 演示

## 直接广告的工具

wps_ppt_get_slide_count、wps_ppt_get_slide_info、wps_ppt_get_shapes、wps_ppt_set_shape_fill、wps_ppt_set_shape_text、wps_ppt_add_slide、wps_ppt_set_slide_title、wps_ppt_set_slide_content、wps_ppt_add_textbox、wps_ppt_insert_table、wps_ppt_insert_ppt_image、wps_ppt_export_slide_as_image、wps_ppt_open_presentation，以及通用的 wps_common_save、wps_convert_to_pdf。

演示是三个应用里工具最多的一个（115 个）。动画、切换、母版、配色、图表、流程图、KPI 卡片、时间线、3D 文字等全部通过 wps_call 使用，清单见同目录 reference.md。

## 参数约定

- 新增形状与文本框用 left、top、width、height，单位像素。没有 x、y 参数——传了不会被发现（可选参数拼错是静默忽略），只是完全不起作用，所以务必按 schema 的名字传。
- 插入图片用 path；相对路径会按当前工作目录解析，文件不存在会明确报 image file not found。
- **形状定位的三种写法**，语义不同，别混用：
  - shapeIndex / shapeName / name：任意形状的序号或名称；
  - imageIndex：**第几张图片**（只数图片）；
  - textboxIndex：**第几个文本框**（只数文本框）。
  用 get_shapes 拿到列表后再选合适的那个。
- 动画：set_animation / add_animation 的动画名支持 fadeIn、flyIn、wipeIn、zoomIn、bounceIn、spinIn、fadeOut、flyOut，也可以直接传 MsoAnimEffect 数值；trigger 取 onClick / withPrevious / afterPrevious。add_animation_preset 传 shapeIndex 时只给该形状加动画，不填则给整页每个形状都加。
- 切换：set_transition / set_slide_transition 的效果名支持 none、cut、fade、dissolve、push、wipe、split、reveal、cover、curtains（也可传 PpEntryEffect 数值），duration 单位为秒；apply_transition_to_all 会对所有幻灯片应用。
- 背景：set_slide_background / set_master_background 接受 background 对象，形如 {type:"solid",color:"#FF0000"}、{type:"gradient",colors:["#1a1a2e","#0f3460"]}、{type:"image",imagePath:"..."}。缺色或未知类型会明确报错。
- 幻灯片页码只支持显示/隐藏（show），没有起始编号参数。

## 常用流程

1. wps_ppt_get_slide_count 与 get_slide_info 确认文稿与页码。
2. 多份演示文稿同时打开时，先用 set_active_target 锁定目标文稿（通过 wps_call），否则 ActivePresentation 会漂移。
3. **新建文稿后先 add_slide**：WPS 的 Presentations.Add() 返回 null，且新建文稿通常是 0 页；此时 get_slide_count 可能仍报 0，不要据此判断文稿为空。
4. 加页用 add_slide，然后 set_slide_title 与 set_slide_content 填内容。
5. 需要精细排版时先 get_shapes 取形状索引，再对具体形状操作。
6. 导出前先 wps_common_save，导出用 export_slide_as_image（绝对路径）。

## 已知坑与不支持的能力

- 当前这版 WPS 里，**通过 COM 新建的文稿 Slides.Count 会一直报 0**，所以 get_slide_count 可能返回 0 而幻灯片确实存在。以 get_slide_info / get_shapes 的实际返回为准。
- 动画与切换此前整族不可用（底层要数值而工具传的是名字），现已修好；但名称仍必须在上面的列表内，写错会明确报 unknown animation / unknown transition 并列出可用名称。
- **环形图只画单一占比**：create_donut_chart 用 value（0-1 或百分数），不支持多段 data——WPS 的扇形只暴露一个 adjustment，赋值会挂住 COM 宿主。
- **组织架构图只生成默认骨架**：create_org_chart 不接受自定义节点（传自定义节点会 60 秒超时）。
- **图表不注入数据**：insert_ppt_chart / add_chart 只插入指定类型的空图表并可选标题，数据请在 WPS 中填写；data 参数已移除。
- **渐变角度/类型不可设置**：set_shape_gradient 只支持 stops（两个色标），没有 angle / type。
- 设置形状阴影、边框、渐变时传的是嵌套对象（shadow / border / gradient），对象里多写 action 不认识的字段会报错。
- 配色相关工具有两套枚举（beautify 的内置配色与 apply_color_scheme 的 ocean/forest/sunset 等），不要混用。
- 定位到不存在的目标文稿时，部分底层实现会继续解引用空对象；锁定目标后要复验。

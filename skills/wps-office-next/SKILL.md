---
name: wps-office-next
description: 通过 mcp__wps-office-next__* 工具操控 WPS 表格、文字与演示；负责路由、目标确认、写后校验与 wps_call 派发。
whenToUse: 用户要求读取、创建、编辑、排版、分析、导出 WPS 文档，或需要跨 WPS 应用搬运数据时。
---

# WPS Office（Windows / COM）

本技能覆盖 WPS 表格、文字、演示三大应用。工具名形如 mcp__wps-office-next__<工具名>。

## 先做三件事

1. 调用 wps_status 确认 WPS 已启动、当前活动应用与工具面档位。没连上就不要继续。
2. 读目标状态：表格用 wps_excel_get_sheet_list，文字用 wps_word_get_active_document，演示用 wps_ppt_get_slide_count 与 wps_ppt_get_slide_info。
3. 多个文档、工作簿或演示同时打开而用户没指明目标时，先问清楚，不要猜。

## 参数名：先查 schema，别照抄示例（重要）

**工具层与桥之间**的参数名现在会被校验：任何工具发给桥的多余键都会得到一个明确错误，并列出被拒的键与该 action 接受的键，例如

    unknown parameter(s) for 'evaluateFormula': celll | accepted: formula

这条保护针对的是工具自身的实现缺陷——以前"发了但没人读"等于什么都没做却回报成功。

**调用方传的参数名不受这条校验保护**，务必区分：

- **必填参数写错名**（例如用 rng 代替 range）会被 schema 校验挡下：`Missing required parameter: range`；
- **可选参数写错名不会被发现**：调用照常执行、静默使用默认值，随后可能因为别的原因失败（例如在没有打开文稿时得到 no presentation is open），错误信息与拼错的名字毫无关系。

所以：不确定参数名时先 wps_help {tool:"wps_ppt_set_animation"} 取完整 schema，再按 schema 里的名字调用，不要照搬其它工具集或旧示例的写法（比如用 imagePath 代替 filePath、用 transitionType 代替 transition）。少数工具为兼容保留了两套拼写（例如插入图片同时接受 filePath 与 path），但**以 schema 为准**最稳妥。

## 工具面

默认 standard 档直接广告 43 个工具，其余工具仍然完全可用：

- 直接广告的工具见各应用技能与同目录 reference.md。
- 未广告的工具：先 wps_help {app:"ppt"} 查目录，或 wps_help {query:"chart"} 搜索，再用 wps_help {tool:"wps_ppt_set_animation"} 取完整参数 schema，最后用 wps_call {tool, args} 执行。
- 多个连续操作可用 wps_batch 一次提交，最多 50 项。
- 未广告的工具也可以按全名直接调用，wps_help 只负责让你发现它们。

## 变更协议

按这个顺序做，不要跳步：

1. 先读取受影响的内容或结构，除非用户明确要新建文档。
2. 破坏性或大范围操作前，说明将影响的确切范围（区域、页、幻灯片、文件、对象）；范围含糊就先确认。
3. 大改动前先 wps_common_save，保留可回退版本。
4. 用最小的原子操作完成任务；表格优先用 wps_excel_write_range 批量写入，不要逐格写。
5. 变更后用读操作复验，并检查返回里的 success 字段；success 为 false 时把原始错误如实报告，不要当成成功。

## 保存、关闭与转换

- **关闭不会弹模态框**：三个 close 工具（表格/演示有对应工具，文字用 wps_call 调 closeDocument）都会关闭对话框保护。对**从未落盘**的文档，即使要求 save=true 也会改为不保存关闭，并在结果里返回 warning——此时要如实告诉用户文件没有写盘，必要时改用 save_as。
- **wps_common_save_as 的路径键是 filePath**，另存为失败或未写盘时不要报告成功。
- **转换要指定应用**：wps_convert_to_pdf / wps_convert_format 支持 app_type（excel/word/ppt）。不指定时按 Excel → Word → PPT 取第一个正在运行的文档——**Excel 常开着会让"把 Word 转成 PDF"导出工作簿**，所以转换前显式传 app_type。指定了却没有对应文档会明确报错，不会退回别的应用。
- openAfterExport=true 时导出后会自动打开文件，结果里回报 opened / failed。

## 单位与坐标

| 场景 | 单位 |
|---|---|
| 表格列宽 / 行高 | 字符宽度 / 磅 |
| 文字页边距 | 磅，**整数**，0-1584 |
| 演示 add_shape、add_textbox、set_shape_position | 像素，左上角为原点，向右下为正 |
| 演示 insert_ppt_image、字号 | 磅 |
| 颜色 | 十六进制字符串，如 #FF0000 |

同一概念在不同工具里单位可能不同，调用前用 wps_help {tool:"..."} 确认。

## 失败处理

- 看不到 mcp__wps-office-next__* 工具：说明 WPS MCP 桥未加载，提示检查插件安装并用 dsh --profile <name> --dump-config 验证。
- 工具在但 connected 为 false：请用户启动 WPS 12.1+ 并打开对应文档，不要反复重试。
- 返回 unknown parameter(s)：工具层发给桥的键有误，按错误里列出的 accepted 名字重试，不要原样重试。
- 返回 no presentation is open：当前没有打开的演示文稿，先 open_presentation 或 create_presentation 再 add_slide。
- 返回 Missing required parameter: X：调用方漏了必填参数。
- 结果里出现 `warnings`：本次操作**已经完成**，但其中某个「尽力而为」的步骤失败了。
  涉及用户关心的结果时要如实转述（例如「文字已插入，但样式未应用：<原因>」），
  不要因为 success 为真就当作完全成功。
- 返回 image file not found / sound file not found：路径不存在，确认绝对路径或相对仓库根目录的路径。
- 返回 slideIndex is out of range：页码确实越界（下界必查；上界在 WPS 计数可信时才查）。
- 返回 unknown transition / unknown animation：名称不在支持列表，错误里会给出可用名称，也可以直接传数值。
- 操作超时：常驻宿主会自动重启，先确认 WPS 是否卡在模态对话框上。

## 参考

同目录 reference.md 列出全部工具及其所属档位。

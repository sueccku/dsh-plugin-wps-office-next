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

## 工具面

默认 standard 档只广告 31 个工具，其余工具仍然完全可用：

- 直接广告的工具见各应用技能与同目录 reference.md。
- 未广告的工具：先 wps_help {app:"ppt"} 查目录，或 wps_help {query:"chart"} 搜索，再用 wps_help {tool:"wps_ppt_set_animation"} 取完整参数 schema，最后用 wps_call {tool, args} 执行。
- 多个连续操作可用 wps_batch 一次提交，最多 50 项。

未广告的工具也可以按全名直接调用，wps_help 只负责让你发现它们。

## 变更协议

按这个顺序做，不要跳步：

1. 先读取受影响的内容或结构，除非用户明确要新建文档。
2. 破坏性或大范围操作前，说明将影响的确切范围（区域、页、幻灯片、文件、对象）；范围含糊就先确认。
3. 大改动前先 wps_common_save，保留可回退版本。
4. 用最小的原子操作完成任务；表格优先用 wps_excel_write_range 批量写入，不要逐格写。
5. 变更后用读操作复验，并检查返回里的 success 字段；success 为 false 时把原始错误如实报告，不要当成成功。

## 单位与坐标

| 场景 | 单位 |
|---|---|
| 表格列宽 / 行高 | 字符宽度 / 磅 |
| 演示 add_shape、add_textbox、set_shape_position | 像素，左上角为原点，向右下为正 |
| 演示 insert_ppt_image、字号 | 磅 |
| 颜色 | 十六进制字符串，如 #FF0000 |

同一概念在不同工具里单位可能不同，调用前用 wps_help {tool:"..."} 确认。

## 失败处理

- 看不到 mcp__wps-office-next__* 工具：说明 WPS MCP 桥未加载，提示检查插件安装并用 dsh --profile <name> --dump-config 验证。
- 工具在但 connected 为 false：请用户启动 WPS 12.1+ 并打开对应文档，不要反复重试。
- 返回 Unknown action：该工具在工具层注册但底层 COM 未实现，属于已知缺陷；改用等价工具并如实说明。
- 操作超时：常驻宿主会自动重启，先确认 WPS 是否卡在模态对话框上。

## 参考

同目录 reference.md 列出全部工具及其所属档位。

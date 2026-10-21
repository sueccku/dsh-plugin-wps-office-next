# 能力范围与边界（P0，2026-10-19）

> 目的：让模型在动手前就知道"这个插件管什么、不管什么"，避免用 `wps_execute_method` 反复试探
> 不存在的能力。所有"不支持"都应先说明原因，再给替代路径；不要假装成功。

## 1. 支持范围

- **本机正在运行的 WPS Office（Windows x64 / COM）**：WPS 表格、WPS 文字、WPS 演示。
- 本地单个文件的读、写、排版、计算、导出、格式转换；同一时间只允许一个 DSH 会话操作 WPS。
- 打开/新建/切换/关闭文档；保存、另存（可自动创建输出目录）。
- Excel：数据读写、公式、格式、条件格式、数据验证、图表、透视表、页面设置、打印设置、图片/批注/超链接、保护。
- Word：正文/段落/样式、表格、页眉页脚/分节、目录、书签/交叉引用/脚注尾注、批注/修订、邮件合并、内容控件。
- PPT：幻灯片/版式/主题/母版、形状/文本框/图片/表格/图表、动画/切换、备注、页脚、导出图片。
- 格式转换：当前文档转 PDF / 另存为其他格式（`wps_convert_to_pdf` / `wps_convert_format`）。

## 2. 明确不支持（平台/产品边界）

| 能力 | 现状 | 正确响应 |
| --- | --- | --- |
| PDF 合并/拆分/转 Office/加水印 | 插件只有 `convert_to_pdf`（单向外导），没有 PDF 宿主 | 说明"插件不处理 PDF 输入"，建议用 WPS PDF 组件/外部工具；不要调 `convert_from_pdf` |
| 打印/打印预览 | 已决策不做（FIXES 1169-1171）；只做页面设置 + 导 PDF | 设置好打印参数，导出 PDF，由人打印；不要调 `PrintOut` |
| 云文档多人协作/共享冲突 | 本地桌面 COM 插件不覆盖云文档并发 | 说明边界；建议用 WPS 云文档客户端/服务端协作 |
| 手机端 WPS | 桌面 COM 不覆盖移动端 | 说明边界；建议在 PC 端完成 |
| WPS AI 生成/改写 | 不在插件工具面 | 说明边界；由用户使用 WPS AI 界面 |
| OLE/链接对象/活链接 | 无 OLE 工具；WPS 放映态也不自动刷新链接图表 | 一次性写入数据；Excel 图表可导 PNG 再插入 PPT |
| 水印 | WPS 页眉 `Shapes` 空壳，COM 不可用 | 说明 WPS 限制；需要水印请在 WPS 界面手工加或后期处理 |
| 文档属性/PDF 元数据 | `BuiltInDocumentProperties`/`CustomDocumentProperties` 是坏壳 | 说明 WPS 限制；在 WPS 界面手工改 |
| 媒体/音频/视频、SmartArt、3D | 已决策删除（D3） | 说明不支持；让用户在 WPS 界面插入 |
| 密码解密/忘记密码 | WPS 不支持撤销保护密码 | 明确拒绝；不要尝试爆破 |
| 宏/VBA | 插件不执行宏 | 说明边界；需要宏请在 WPS 内手工运行 |
| 切片器/方案管理器/模拟运算表 | WPS COM 限制（`Slicers.Count=0`、`Scenarios` 方法） | 说明限制；用透视字段/多表筛选/单变量求解替代 |

## 3. 安全默认值（P0 变更）

- `open_workbook` / `open_document` / `open_presentation` 默认 **只读**（`readOnly=true`）。要编辑必须显式传 `readOnly:false`。
- `close_workbook` / `close_document` / `close_presentation` 默认 **不保存**（`save=false`）。要保存必须显式传 `save:true`。
- `wps_batch` 默认 **遇到第一个失败就停**（`stopOnError=true`）；独立步骤要逐项继续时显式传 `false`。
- `wps_common_save_as` 默认 **自动创建输出目录**（`createDirs=true`）；要严格校验目录时传 `false`。
- `wps_excel_read_range` 默认最多返回 500 行；更大范围用 `startRow`/`maxRows` 分页。
- `wps_word_get_document_text` 支持 `start`/`end`（0 基字符索引），单次最多 10,000 字符；长文分页读。
- 破坏性/全局改动前：先 `wps_common_save_as` 副本，再改；不要承诺 Ctrl+Z 可撤销。

## 4. 发现与工具调用

- 先在 `wps_help` 查：`{query:"题注/文档比较/母版/限制编辑"}` 现在会做同义词扩展；`{app:"excel"}` 默认返回 200 条并支持 `offset/limit`。
- 同名不同义：`wps_help` 结果带 `app` 字段；"题注"可能命中 Word 样式名，"保护"可能命中 Excel 保护——按当前应用选择。
- 未广告工具用 `wps_call`；连续独立调用用 `wps_batch`；`wps_execute_method` 只转发桥 action，不是万能 COM 逃生舱。

## 5. 交付前的诚实报告

- 已执行（调了什么）、已核对（回读值）、未验证（视觉/渲染/PDF 页数）分开说。
- 不支持的场景先报边界，再给替代路径；不要用"已尽力"掩盖没做成。

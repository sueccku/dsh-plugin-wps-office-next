# 270 个工具 × 20 个真实场景：设计评估与改进建议

> 日期：2026-10-19（Asia/Shanghai）　基线：v0.9.0 检出（`D:\dsh\a`）
> 方法：**网上调研 20 个真实 WPS 用户场景 → 按模型可见面（95 个广告工具 + `wps_help/wps_call/wps_batch`）做纸面仿真 → 对工具名/参数/桥实现做 schema 与源码核验**。
> 状态：**未跑 WPS 真机、未跑仓库测试**（遵循 [AGENTS.md](../AGENTS.md) §1 的 R13）。本文是"设计层评估"，真机行为另列待验证项（§7）。
> 过程产物：`.tmp-eval/research-*.md`、`tool-inventory.md`、`tool-index.json`、`sim-*.md`、`verification.md`（均在本机、未入库）。
> 数字以独立核验后的口径为准；核验记录见 `.tmp-eval/verification.md`（97 条：81 confirmed / 14 incorrect / 2 unverifiable）。
> **后续（2026-10-19）**：在 20 场景基础上新增 60 个场景，综合 80 场景的 v2 报告见 [scenario-review-270-tools-v2.md](scenario-review-270-tools-v2.md)；本文的 20 条细粒度调用链仍有效。

---

## 0. 结论摘要

### 0.1 一句话结论

**270 个工具作为"WPS COM 能力底座"是合理且有深度的；作为"模型直接面对的用户任务面"不合理。**
问题不在数量，而在**形状**：工具按 COM 动词组织、按活动对象/选区寻址、把"任务"拆成原子调用；真实用户说的是"把总表做对""只改颜色别动字号""一人一份合同"，工具面接不住这些约束。

### 0.2 20 个场景的判定

判定口径：A 顺畅（≤1.5× 最小路径、几乎不用 help）／B 可行但绕／C 需拼凑（>3× 或绕路）／D 做不到（核心产出无工具）／E 风险（可与 A–D 叠加，单列）。

| 判定 | 数量 | 场景 |
| --- | ---: | --- |
| A 顺畅 | **0** | — |
| B 可行但绕 | **2** | E4 数据清洗、W4 修订批注 |
| C 需拼凑 | **13** | E1、E3、E5、E6、W1、W2、W3、W5、W6、P1、P2、P3、P4 |
| D 做不到 | **5** | E2 公式条件格式、E7 图表日期轴/空值、E8 云协作、X1 跨应用活链接、F1 批量打印 |
| E 风险（叠加） | **10** | E1、E3、E7、E8、F1、W1、W3、P1、P3、P4 |

- **A=0 不等于"全都不能用"**：真正的基础任务（E4、W4）是 B；13 个 C 说明"能做，但要绕 2.5–8 倍调用"；5 个 D 里 **2 个是刻意选的边界**（E8 云协作、F1 打印——打印已决策不做），**3 个是真实能力缺口**（E2、E7、X1）。
- 另有 **5 个场景含 D 级子项**（W1 多级编号/参考文献、W2 大纲级别、W3 一对多、W5 多行表头、P3 平滑切换）：整条任务能交付，但验收里的核心一项不达标。

### 0.3 成本画像（纸面仿真）

- 20 场景合计 **391 次顶层调用**，平均 **19.6 次/场景**；`wps_help` **105 次**，平均 **5.3 次/场景**；仿真里用到 **41 个隐藏业务工具**。
- 校验过的调用 **1,019 次（170 + 208 顶层 + 641 个 batch 子调用），0 个 schema 错误**——参数写法不是主要矛盾，**选对工具与补足能力**才是。
- 典型倍率（实际调用 / 任务级理想步数）：W4 2.8×、W1 4.3×、W5 5.3×、P4 7.7×、P3 8.0×；F1 的 50 文件循环在工具侧放大约 160 次执行。

### 0.4 五个关键发现

1. **工具面是"COM 动词表"，不是"用户任务面"**：93/270 个工具 `required=0`（34.4%），其中 58 个是写/删/生命周期类；大量写操作默认落在"当前选区/活动对象"上。真实任务从"一个文件路径"开始，模型必须先花 3–7 次调用做定位与探测。
2. **广告面按历史调用次数选，没按任务闭环选**：PPT 77 个只广告 14 个，P1–P4 四个"改已有稿"场景需要的 27 个工具全在隐藏面；Word 广告了 `get_revisions/get_comments`，却把 `accept_revisions/delete_comment/enable_track_changes` 三个动作藏起来。20 场景里 **36 个广告工具一次没用到**（另一面：41 个隐藏业务工具被用到）。
3. **`wps_help` 是名字查找器，不是任务导航器**：`wps_help {app}` 单次只显示 60 行；按真实注册序，**76 个工具在 60 行外（Excel 58 / PPT 17 / Word 1），其中 53 个隐藏、23 个广告，53/175 = 30.3% 的隐藏工具无法按应用目录发现**；9 个用户原词（多级编号/参考文献/悬挂缩进/跨页断行/相似格式/高亮/平滑/显示最终状态/引号）实测 0 命中。
4. **契约层有三类硬伤**：① schema 比桥窄——10 个高影响示例逐条成立（条件格式 type/值/色阶、数据验证提示、`word_insert_table.data`、PPT `applyToAll`、Excel `sheet` 等桥已实现、模型传不进去）；② 没有 `outputSchema`，读取返回"文本表格 / 前缀+JSON / 原始桥 JSON"三种风格，`wps_batch` 还会把每项结果截到 2,000 字符；③ 行列/字符基准混用（5 个 0-based 参数 vs 绝大多数 1-based 描述）、`name` 在 21 个工具里承载 8 种对象。
5. **危险默认值朝向"改源文件"**：`open_workbook.readOnly` 默认 `false`、三个 `close_*` 的 `save` 默认 `true`、`wps_batch.stopOnError` 默认 `false`；Excel/Word 没有 PPT 那样的目标锁（只有 `switch_workbook`/`switch_document` 改全局活动态）。20 场景里 10 条 E 风险，多数来自这三条默认值。

### 0.5 最该做的 6 件事（按性价比排序）

1. **修 `wps_help` 发现面**（去 60 行硬截断/加分页；隐藏工具优先；补 9 个用户词的同义词；加"场景→工具/配方"入口）。
2. **改安全默认值**（`open_*` 默认只读、`close_*` 默认不保存、依赖型 batch 默认 `stopOnError=true`、破坏性操作加 dry-run/预览）。
3. **把桥已有、schema 未暴露的高频键补进 schema**（10 个示例先做），并修正 3 处描述-实现不符（`set_background_gradient`、`set_master_background`、`add_master_element`）与 `get_document_text` 的假 `start/end`。
4. **把 20 个场景写成技能配方**（B 路径：起手体检 → 决策树 → 最小调用序列 → 验收清单 → 失败回退 → 明确"做不到"），不新增任务级桥工具（遵守 C 类边界）。
5. **广告面按任务闭环重排**（在 95→100 / 55,510→≤60,000 字节预算内换血，不扩容；按 R14 真机证据阈值）。
6. **把这 20 个场景固化成 `test/skilleval` 基准**，每次发版量：成功率、调用数、help 数、隐藏工具发现率、源文件是否被改、是否诚实报告。

### 0.6 评分（1–5，权重见 §4）

| 视角 | 分数 | 一句话 |
| --- | ---: | --- |
| 作为 COM 能力底座 | **4.1 / 5** | 覆盖深、错误诚实、分层克制；扣分在 schema 比桥窄、描述与实现不符 |
| 作为模型任务面 | **2.6 / 5** | 0 个 A、5 个 D、平均 19.6 次调用 / 5.3 次 help、无任务级组合与验证面 |
| 综合（当前形态） | **3.1 / 5** | 能用，但"能用"的代价是绕路、上下文与误改风险 |
| 做完 P0/P1 后的目标 | **3.8–4.0 / 5** | 不动 270 的规模，只改面与配方 |

---

## 1. 方法与边界

### 1.1 场景来源

- 从中文互联网调研 **22 个候选场景**（WPS 官方社区 bbs.wps.cn、WPS 学堂、正保会计网校问答、可打开的百家号/办公博客等），选 **20 个**进入仿真。
- 证据分级：每个场景至少 1 个 `web_fetch` 打开且直接支持的页面；知乎/B站/小红书/百度知道在本环境 404 或登录墙，均降级标注，未编造引文或热度。
- 独立核验：抽查 **16 个主来源 URL**，全部可打开且逐条支持对应场景；另有 1 个知乎链接 404（与调研自述一致，未作为来源使用）。**未发现编造引文或把打不开页面当 A 级来源**。
- 未入选：X2（PDF 导出质量，已由 E6/W2 的导出步骤间接覆盖）、F2（云协同，与 E8 同属边界，保留 E8 即可）。

20 个场景及来源见 §2 与附录 A。

### 1.2 仿真方法

- 站在"模型真实可见面"上：95 个广告工具 + `wps_status`/`wps_help`/`wps_call`/`wps_batch`。需要隐藏工具时，先写 `wps_help` 查询，再 `wps_call`/`wps_batch`。
- 每个场景写出：用户原话 → 起始文件 → 必须先探测什么 → 完整调用链（工具名 + JSON 参数 + 目的 + 广告/隐藏 + help 词 + 预期返回/风险）→ 步数与理想步数 → 摩擦点（主类 + 证据）→ A–E 判定 → 改进建议。
- 工具名/必填/类型/枚举/基准对 `spec/tool-definitions.json` 逐条校验；判"做不到"时同时读 `mcp/scripts/wps-com.ps1`/`host/wps-actions.ps1`，避免只凭描述下结论。
- 共校验 **1,019 次调用，0 schema 错误**。

### 1.3 边界与诚实的限制

1. **这是纸面仿真，不是真机执行**。COM 是否接受某个赋值、弹框、超时、性能，都可能改变结论；§7 列出必须真机验证的 12 项。
2. 场景里的规模是典型值（30–60 页 PPT、486 段论文、200 行名单、50 个文件）；规模越大，倍率只会更差。
3. "理想步数"按用户级子目标定义；判定对 ideal 的选择敏感。即使把理想步数放宽 50%，Word/PPT 组除 W4 外仍全部 >3×；Excel 组的 E1/E4/E6 会落到 2–3× 区间，C/B 的判定方向不变。
4. 仿真模型知道"有 `wps_help` 可以查"，但不知道全部 270 个工具——这正是真实模型的处境。
5. 未运行 `scripts/run-tests.ps1` / `e2e.mjs`；本文不声称任何真机通过率。

---

## 2. 20 个场景的结果

### 2.1 总表

| id | 真实任务（用户口径） | 证据 | 判定 | 顶层调用 / help | 核心卡点 |
| --- | --- | --- | --- | ---: | --- |
| E1 | 6 个部门月报合成一张总表，以后加数据能自动更新 | A | C+E | 24 / 1 | 无跨工作簿查询/刷新；活动工作簿无锁；`readOnly=false`+`close.save=true` |
| E2 | 签到表按 3–4 条公式规则给整行上色，空行不变 | A | D | 13 / 3 | 条件格式只支持 cellValue 比较，无公式规则/多规则/整行 |
| E3 | 两表按工号匹配，未匹配显示"未匹配" | A | C+E | 23 / 4 | 公式填充语义不明、多工作簿切换、文本/数字差异诊断 |
| E4 | 一格多个姓名拆成一列一行 + 去重 + 删空行 | A | B | 14 / 3 | 无"分列到行"；清洗工具职责交叉 |
| E5 | 按品类 × 月份做透视表并刷新 | A | C | 21 / 6 | 透视表日期分组/占比要绕；透视三件套全隐藏 |
| E6 | 每页重复表头、缩到一页宽、Excel 从第 3 页接页码 | A | C | 18 / 1 | 无 Excel 起始页码；PDF 不能合并；打印页数/预览已决策不做 |
| E7 | 折线图改日期轴、空值断开、别改公式 | A | D+E | 14 / 3 | 无图表列表；无 axisType/numberFormat/displayBlanksAs |
| E8 | 多人同表只改自己的行、保护密码、共享冲突 | A | D+E | 16 / 7 | 云协作/允许编辑区域超出本地插件；保护状态不可回读 |
| W1 | 论文/公文按模板排样式、页边距、参考文献、引号 | A | C+E | 34 / 9 | 无多级编号/题注/参考文献/悬挂缩进；无固定行距；引号批量替换易错 |
| W2 | 目录页码全 1 / 个别不对 / 正文误入目录 | A | C | 18 / 2 | 无大纲级别读/写、无"链接到前一条页眉"；起始页码只能从 1 重开 |
| W3 | 模板 + 200 行名单批量出合同，一人一份 | A | C+E | 12 / 0 | CSV-only、占位符语法窄、一对多不支持、批量 PDF 要 200×3 次 |
| W4 | 接受修订、清批注、关修订、交干净版 | A | B | 17 / 8 | 清高亮/显示最终状态无工具；其余动作可核对 |
| W5 | 表格跨页每页有表头、简历不被挤到下一页 | A | C | 16 / 4 | 多行表头只写第 1 行；无"允许跨页断行"；溢出只能靠猜 |
| W6 | 同类文字格式批量统一/格式刷 | A | C | 14 / 4 | 无"选择格式相似"；格式刷仅 Excel 有 copy_format |
| P1 | 有内容的 PPT 套公司模板，字体/LOGO/页脚不跑偏 | B | C+E | 28 / 13 | 主题/母版/页脚/统一字体/对齐全隐藏；LOGO 只能逐页猜形状 |
| P2 | 6 张图排一行、等距、整排相对页面居中 | A | C | 22 / 9 | 对齐/分布/组合隐藏；`align` 无相对页面参数；索引语义易错 |
| P3 | 全篇统一动画与"平滑"切换，播放不闪 | A | C+E | 24 / 11 | 无 animation-to-all；"平滑"不在 EntryEffect 映射；preset 会给所有形状加动画 |
| P4 | 只改文本框字色不动字号/占位符/表格；全文替换 | A | C+E | 23 / 8 | 无对象类型/页码范围过滤；无字体属性回读；250 次逐形状改色 |
| X1 | Excel 数据/图表进 PPT，保留可刷新链接 | A | D | 30 / 6 | 无粘贴链接/更新链接；只能一次性写死；单位/数据结构不统一 |
| F1 | 文件夹 50 个文档批量打印 + 批量导出 PDF | A | D+E | 10 / 3 | 打印已决策不做；批量导出要 150 次循环、4 个 batch，无失败清单 |

> 关于 F1：如果按"插件目标范围内"剔除已决策不做的打印，F1 降为 **C+E**（批量导出可编排但极贵、默认 save 有回写风险）；保留 D 是为了忠实反映用户完整诉求。
> 关于 E8/X1：它们是**能力边界**（本地桌面插件不做云协同/活链接），不是"工具写错"；但工具面必须能给模型一个结构化的"做不到 + 替代路径"，而不是让模型去裸试。

### 2.2 判定分布说明了什么

- **0 个 A**：即使最基础的 E4/W4 也要 14–17 次调用、3–8 次 help。按严格 rubric，没有一条真实任务能"顺手完成"。
- **13 个 C**：不是功能没有，而是"要绕"——探测 3–7 次、隐藏工具 2–10 个、验证 0–4 次。这是"270 个原子工具 + 没有任务面"的直接结果。
- **5 个 D + 5 个含 D 子项**：缺口集中在 **"域/结构"层**（公式条件格式、图表坐标轴、多级编号、题注、参考文献、大纲级别、多行表头、平滑切换）与 **跨对象/跨文件层**（活链接、批量打印、云协作）。这些恰好是中文办公场景里"验收最硬"的部分。

### 2.3 调用成本画像

- **定位与探测吃掉大量预算**：`wps_status`、`open_*`、`get_*_list`、`get_*_info`、`read_range` 在 20 场景里反复出现；E1 的 24 次调用里只有 4 次是真正的业务写（`write_range`×2、`remove_duplicates`、`create_list_object`）。
- **隐藏工具发现成本是"每次任务"的固定税**：用到隐藏工具的场景，平均每个隐藏工具要 2 次调用（1 次 help + 1 次 call）；P1 用了 13 次 help，P3 11 次，P2 9 次。
- **"改已有稿"比"从零生成"更贵**：P1–P4 全部是改别人/旧版的稿子，广告面 14 个 PPT 工具里 10 个"新建/写"类一次没用，真正需要的 27 个全在隐藏面。
- **批量任务线性放大**：F1 的 50 文件 ×3 次 = 150 次工具执行，`wps_batch` 上限 50 项，需要 4 批；P4 的 250 次改色靠 6 个 batch 压成 23 次顶层调用，但工具侧仍是 250 次 COM 调用。
- **验证预算几乎为零**：10 个场景只有 0–4 次验证调用；PPT 没有字体属性读回，Word 没有渲染/预览，图表没有属性读回，PDF 没有页数核验。

### 2.4 五个代表性场景（读这五条就懂全局）

#### E1 多文件汇总：能合，但不能"自动更新"
- 6 份源表 + 1 个模板：先 `status`/`get_open_workbooks`/`get_sheet_list`/`get_sheet_info`/`read_range`，再 `open`/`write_range`/`remove_duplicates`/`create_list_object`/`save_as`/`close`，24 次调用。
- **"新增数据自动更新"没有工具**：`create_pivot_table` 只接单表区域，`consolidate` 只合同簿区域，`refresh_*` 只刷透视表/外部链接；只能写静态值或 ListObject 等手工追加。
- **两处 E 风险来自默认值**：`open_workbook` 的 `readOnly` 默认 false；`close_workbook` 的 `save` 默认 true。验收要求"源文件一个字节不变"，模型必须每一步都记得显式传 `true/false`。
- 判定 C+E；理想任务级 2 步，现有工具面最小 16 步。

#### W1 长文档排版：最贵的不是排版，是"坐标与域"
- 486 段、100+ 页：`get_paragraphs` 50 段/次 → 10 次读取；`get_document_text` 的 `start/end` 是假参数（桥只取 `Content.Text` 再截前 10,000 字符），帮不上长文。
- 标题套样式要 50+8 次 `apply_style`；引号统一没有"成对替换"，只能 86 次 `replace_range`（等长、可审计）；`find_replace` 一把梭会把 86 个直引号全变成开引号（E 风险的来源）。
- **多级编号、题注、参考文献自动编号、悬挂缩进没有任何工具**，`wps_help` 对这些词 0 命中 → W1 的 D 子项。
- 判定 C+E；任务级理想 8 步，实际 34 次顶层 + 159 次 batch 子调用。

#### W3 批量合同：`mail_merge` 是全仓库最好的"任务级工具"，但边界太窄
- 一次 `mail_merge` 就能生成 200 份 docx（`dataFile` + `splitDir` + `fileNameField`），这是"一个工具完成一个任务"的正面样板。
- 但它**只吃 UTF-8 CSV**（真实名单是 xlsx/.et），占位符只认 `⟨⟩/{{}}/«»`（用户模板常写【】），**一对多明细不支持**；批量 PDF 仍要 200 次 open/convert/close。
- 判定 C+E；说明"工具不需要多，需要的是**输入面宽 + 输出可核对**"。

#### P4 只改字色：属性粒度与验证面的双重缺失
- `set_font_color` 只有 `slideIndex/shapeIndex/color`，没有"只改文本框""只改第 1–50 页""占位符/表格跳过"的过滤参数；50 页 × 5 文本框 = **250 次调用**（6 个 batch，子调用 300）。
- 更致命的是**没有读字体属性的工具**（`get_textboxes/get_shapes/get_slide_info` 都只给文本与几何），改完无法证明"字号没动、表格没动"。
- 判定 C+E；这是"setter 视角 vs 用户约束视角"的典型。

#### X1 跨应用活链接：能力边界必须显式说清
- 三页静态 PPT 可做：读 Excel → 锁 PPT 目标 → `add_slide`×3 → `insert_ppt_chart` + `set_ppt_chart_data` → 存盘，30 次调用。
- **"明天现场改 Excel、PPT 跟着刷新"做不到**：没有粘贴链接/OLE 链接/更新链接工具；WPS 社区管理员也明确答复"放映时画面不动无法自动刷新图表"（来源见附录 A-X1）。
- 判定 D；正确做法不是在桥里加链接（C 类：跨应用工作流不进桥），而是在技能里写清"一次性写入 + Excel 图表导 PNG"的替代路径，并让工具返回结构化的 `unsupported` 原因。

---

## 3. 270 工具面审计（核验后）

### 3.1 总量与分层

| 项 | 值 |
| --- | --- |
| 注册工具 | **270** = Excel 118 / Word 61 / PPT 77 / 通用 7 / 转换 2 / 门面 5 |
| 直接广告 | **95** = 4 门面 + 91 应用（Excel 46 / Word 27 / PPT 14 / 通用 2 / 转换 2） |
| 隐藏 | **175** = Excel 72 / Word 34 / PPT 63 / 通用 5 / 门面 1 |
| 广告面载荷 | **55,510 字节 / ≈15,860 tokens**；D1 上限 100 工具 / 60,000 字节（余 5 位 / 4,490 字节） |
| 全量 schema | 165,422 字节（270 工具） |

分层本身是对的：95 个广告工具约占 15.9k tokens/请求，若全量广告则 165KB（≈47k tokens），不可接受。**问题不在"藏了 175 个"，而在藏了之后没有好的发现面**。

### 3.2 结构性问题 1：任务面缺失，工具在"活动对象"上做文章

- **93/270（34.4%）工具 `required=0`**；其中按 effect 口径 **58 个是写/删/生命周期类**。典型：`set_sheet_page_setup`（16 参数/0 必填）、`set_shape_effect`（17/0）、`set_table_format`（14/0）、`update_pivot_table`（12/0）、`add_animation`（10/0）、`update_chart`（10/0）。
- 逻辑必填未标：`set_shape_position` 把 `left/top` 标必填，却把真正需要的形状选择器 `shapeIndex/shapeName` 留成可选。
- **三套目标模型**：
  - Excel：84 个工具带 `sheet`，但 **13 个操作型工具没有 `sheet`**（`find_replace`、`sort_range`、`set_conditional_format`、`freeze_panes`、`lock_cells`…）；桥侧其实读 `sheet`，schema 没暴露。
  - Word：**2/61** 用 `name` 定位文档（close/switch），其余 59 个一律作用活动文档；0 个 `documentId`。
  - PPT：56 个用 `slideIndex`，**5 个接受 `presentationName`**；但桥 action 里有 **70 个**读 `presentationName`，其中 **65 个 schema 没暴露**——只有 PPT 有 `set_active_target`，Excel/Word 没有对应物。
- **后果**：多文件/多表任务靠 `switch_*` 改全局活动态；`wps_batch` 顺序执行 + 默认 `stopOnError=false`，前置切换失败后后续步骤会落到错误目标上。这是 20 场景里 10 条 E 风险的主要来源。

### 3.3 结构性问题 2：广告面按"调用次数"选，没按"任务闭环"选

- 过去的重平衡依据是"历史轨迹里被 `wps_call` 绕进来的次数"（FIXES 90 / R14），这在"已知要调什么"的前提下有效；但面对**新任务类型**会失灵。
- 20 场景的实证：
  - PPT 14 个广告工具里，**10 个"新建/写"类在 P1–P4 一次没用**（这 4 个场景全是"改已有稿"）；真正需要的 **27 个隐藏工具**（主题/页脚/统一字体/对齐/分布/切换/动画/改色/目标锁）一个都不在广告面。
  - Word 侧：`get_revisions`/`get_comments` 广告了，`accept_revisions`/`delete_comment`/`enable_track_changes` 三个动作没广告——"读得到、改不了"的典型不对称。
  - Excel 侧：`get_pivot_tables`/`refresh_pivot_tables`/`diagnose_formula`/`lock_cells` 全隐藏，而 `wps_help {app:"excel"}` 的 60 行截断又让其中一部分连目录都看不到。
- **36 个广告工具在 20 场景里一次没用到**（E4 类基础任务只用 `open/read/write/close`；这不能证明它们无用，但能说明"广告面覆盖的是上一批任务的形状"）。

### 3.4 结构性问题 3：发现面被 60 行截断 + 中文词表不足

- 代码事实：`mcp/src/server/mcp-server.ts:379-383` 的 `pool.slice(0, 60)`，无 `offset/limit`。
- 真实注册序（不是 spec 的字母序）下：**76 个工具在 60 行外（Excel 58 / PPT 17 / Word 1），53 个隐藏 / 23 个广告；53/175 = 30.3% 的隐藏工具无法通过 `wps_help {app}` 被发现**。关键位置：`get_sheet_settings`(102)、`set_sheet_print_titles`(104)、`get_pivot_tables`(110)、`refresh_pivot_tables`(111)、`consolidate`(91)、`lock_cells`(72)、`unprotect_sheet`(71)。
- 中文匹配用 CJK bigram 兜底，方向是对的；但 9 个用户原词 0 命中：**多级编号、参考文献、悬挂缩进、跨页断行、相似格式、高亮、平滑、显示最终状态、引号**。"格式刷"返回 26 条无关工具。
- 技能 `reference.md` 能补（仿真里也确实靠它找到了部分工具），但按 rubric 每次仍要 2 次 help；且 reference 是"按工具分组"，不是"按用户任务分组"。

### 3.5 结构性问题 4：契约不一致，模型容易"选错/传错"

**（a）范围/地址**：79 个地址类参数，四种表达——Excel A1 字符串、Word 0-based 字符偏移对象、PPT `slideIndex + shapeIndex`、`create_chart.position` 未声明子属性的对象。
**（b）基准**：5 个 0-based 参数（`create_sheet.position`、`copy_sheet.position`、`move_sheet.position`、`replace_range.startPos`、`set_line_spacing.paragraphIndex`）vs 绝大多数 1-based 描述；`endPos` 与 `get_document_text.start/end` 未写基准；`position` 在 10 个工具里至少 4 种语义。
**（c）选择器**：`name` 出现在 21 个工具，承载 8 种对象；`shapeIndex`/`shapeName`/`name`/`textboxIndex`/`imageIndex` 五套形状选择器；`get_shapes` 的 index 与 `delete_textbox.textboxIndex` 不是同一索引空间。
**（d）返回**：270 个工具 **0 个 `outputSchema`**；`read_range` 描述说"返回二维数组"，实际返回文本表格；`get_cell_info` 是"前缀+JSON"；`wps_call` 是原始桥 JSON；`wps_batch` 每项结果截 2,000 字符（错误字段不截）。
**（e）schema 比桥窄**：按不同统计口径，约 68–121 个工具存在"桥读、schema 不收"的键；**聚合数不可复现，但 10 个高影响示例逐条成立**：

| 工具 | 桥支持、schema 未暴露 | 用户可见后果 |
| --- | --- | --- |
| `wps_excel_set_conditional_format` | `type`/`value`/`operator`/`colorScaleType` | 只能 cellValue；色阶/数据条做不了 |
| `wps_excel_set_data_validation` | `inputTitle`/`inputMessage`/`errorTitle`/`errorMessage`/`showDropdown` | 下拉提示、出错警告设不了 |
| `wps_excel_set_cell_style` | `backgroundColor`/`bold`/`border`/`fontName`/`fontSize`… | 只能套预定义样式名 |
| `wps_excel_find_replace` | `range`/`sheet`/`matchWholeWord` | 只能替换活动表 |
| `wps_word_insert_table` | `data` | 只能插空表再逐格写 |
| `wps_word_set_text_color` | `range` | 只能改当前选区 |
| `wps_ppt_set_background_color` | `applyToAll`/`presentationName` | 不能一次全册 |
| `wps_ppt_delete_shape` | `name`/`presentationName` | 不能按名删/删非活动文稿 |
| `wps_ppt_insert_table`/`insert_ppt_chart` | `width`/`height` | 插入后不能定尺寸 |
| `wps_word_insert_header`/`footer` | `alignment`/`includePageNumber` | 页眉不能对齐/自动页码 |
| 65 个 PPT 工具（桥 70 个读 `presentationName`） | `presentationName` | 不能逐调用指定文稿 |

**（f）描述-实现不符（3 处已核实）**：
- `wps_ppt_set_background_gradient`：描述说 `gradient:{type,angle,colors,stops}`，桥只读 `color1/color2`，按描述传会报 unknown parameter。
- `wps_ppt_set_master_background`：描述写 `background.path`，桥读 `imagePath`。
- `wps_ppt_add_master_element`：描述说支持 image/path，实现只有 textbox/shape 分支，参数白名单也没有 `path`。
- 另有 `wps_word_get_document_text` 的 `start/end` 是假参数（桥忽略，恒截 10,000 字符）。
- `wps_execute_method` 不是"任意 COM 方法"：它转发的是**桥 action 名**（`wps-client.ts:231-241`），不能当万能逃生舱。

### 3.6 结构性问题 5：能力缺口集中在"域/结构"层

按 20 场景的验收需要，缺口可归为三类（已剔除"打印/媒体"等已决策不做项）：

| 类别 | 缺口 | 影响的场景 |
| --- | --- | --- |
| Excel 分析/可视化 | 公式条件格式（xlExpression）、色阶/数据条（桥有、schema 无）、图表日期轴/空值/轴数字格式、图表列表、Excel 起始页码、插入分页符、多键排序/多条件筛选、跨表/跨簿复制、图片/形状对象管理、透视图、图标集 | E2、E5、E6、E7 |
| Word 结构/域 | 多级编号、题注、参考文献自动编号、悬挂缩进、固定行距、样式定义修改、大纲级别读/写、链接到前一条页眉、多行表头重复、允许跨页断行、格式相似选择、清高亮、显示最终状态、图片环绕/裁剪/删除、表格一次写入 data | W1、W2、W4、W5、W6 |
| PPT 结构/批量 | 平滑/morph 切换映射、逐段动画/动画到全部页、动画/字体属性回读、按对象类型/页码范围改色、对齐相对页面、取消组合、节/隐藏页、表格增删行列/合并单元格、母版加图片/LOGO、表格数据批量写、图表删除/类型切换 | P1、P2、P3、P4 |

> 边界说明：打印/打印预览（FIXES 1169-1171）、水印/文档属性/切片器/方案管理器（WPS 自身不支持）、媒体/SmartArt/3D（D3 已删）**不算缺口**；它们应在技能里写成"明确不做 + 替代路径"。
> 另：`wps_ppt_beautify` 仍在注册面（隐藏），不要对外宣称"美化族 0 残留"；是否刻意保留需 owner 确认。

### 3.7 结构性问题 6：验证面几乎空白

- 改完能回读的只有少数：`get_sheet_info`、`read_range`、`get_conditional_formats`（只序号+类型）、`get_slide_info`（形状几何/文本）、`get_revisions`/`get_comments`。
- 读不到的：PPT 字体属性（颜色/字号/字体名）、占位符类型、Word 渲染/打印预览、图表坐标轴/空值设置、保护状态、PDF 页数。
- 因此大量验收只能"导出图片让人看"（`export_slide_as_image` 1 页 1 次，30 页 = 30 次）——模型可能看不到图片，只能说"未验证视觉效果"。
- 后果：模型只能报告"我调用了什么"，很难证明"我没动什么"（P4 "字号别动"、E1 "源文件别动"、W1 "公式别改"）。

### 3.8 结构性问题 7：安全默认值朝向"会改源文件"

| 默认值 | 现状 | 风险 |
| --- | --- | --- |
| `open_workbook.readOnly` | 默认 `false` | 只读源数据时，漏传就打开可写 |
| `close_workbook`/`close_document`/`close_presentation.save` | 默认 `true` | 批量循环 close 会把 50 个源文件逐个回写 |
| `wps_batch.stopOnError` | 默认 `false` | 前置失败后继续在错误目标上执行 |
| 破坏性操作 | 无 dry-run/预览（部分工具有 `impact`） | 批量改格式/删行前无法预演 |
| 撤销 | 不承诺（README 已写） | 误改只能靠副本，不能 Ctrl+Z |

20 场景的 E 风险里，E1/E3/F1 的根因就是这三条默认值，而不是模型"乱来"。

### 3.9 重叠与冗余（16 组，择要）

| 组 | 工具 | 判断 |
| --- | --- | --- |
| Excel 自动调整 | `auto_fit` / `auto_fit_columns` / `auto_fit_rows` | 该合并为 `auto_fit(mode)`；三个都是 required=0 |
| Excel 读单元格 | `get_cell_value`(row/col) / `get_cell_info`(A1) / `read_range`(区域) | 同一"读一格"三种地址；`get_cell_value` 该并入 `get_cell_info` |
| Excel 格式 | `set_cell_format`（15 参数，含 `numberFormat`）/ `set_number_format` / `clear_formats` vs `clear_range(type=formats)` | 同义入口，建议收敛 |
| Excel 行列 | `hide_rows`/`show_rows`/`hide_column`/`show_columns`/`delete_rows`/`delete_columns` | 参数形态各异；`hide_column` 还要 `hide` 参数 |
| PPT 形状样式 | `set_shape_fill` / `set_shape_style` / `set_shape_effect`（17 参数） | 同一"填充/边框"三套表达，该合并 |
| PPT 背景 | `set_slide_background`（超集）/ `set_background_color` / `_gradient` / `_image` | 三个专用是语法糖，该并入 |
| PPT 删除 | `delete_shape` / `delete_textbox` / `delete_ppt_image` | 三套索引空间，应先 `get_shapes` 再统一删 |
| PPT 文本 | `set_shape_text` / `set_textbox_text` / 占位符三工具 | 文本框就是形状；前两个该合并 |
| Word 页眉页脚 | `insert_header` / `insert_footer` / `insert_page_numbers` / `get_header_footer` | 语义不同，保留；但 schema 比桥窄 |
| Word 文本写入 | `insert_text` / `replace_range` / `find_replace` / `smart_fill_field` / `replace_bookmark_content` | 意图不同，保留；需要决策树 |

> 合并建议必须服从两条已定约束：**D2 Word 长尾全要（不删能力）**；**D4 别名债为 0**（若要保留旧名，需走明确的弃用周期，不要悄悄恢复别名表）。

---

## 4. 设计合理性评分

| 维度 | 权重 | 分数 | 依据 |
| --- | ---: | ---: | --- |
| 任务覆盖/可完成性 | 25% | 3.0 | 20 场景 15 条可交付（B/C），5 条 D（2 条边界）；5 条含 D 子项 |
| 调用效率/粒度 | 15% | 2.0 | 平均 19.6 次/场景；2.8–8× 理想步数；无任务级组合 |
| 可发现性 | 15% | 2.0 | 5.3 次 help/场景；60 行截断 53 个隐藏工具；9 个用户词 0 命中 |
| 参数/契约一致性 | 15% | 2.0 | 93 个 required=0；1/0 基混用；`name` 8 义；schema 比桥窄 |
| 目标选择与安全 | 10% | 2.0 | Excel/Word 无目标锁；readOnly/save 默认危险；batch 默认继续 |
| 能力深度与边界诚实 | 10% | 4.5 | COM 覆盖深；warnings/shortfalls/部分失败诚实；少数描述-实现不符 |
| 上下文成本 | 5% | 4.0 | 95 工具 15.9k tokens 可接受；隐藏层思路正确 |
| 可验证/可回读 | 5% | 2.0 | 无 outputSchema；PPT/图表/Word 渲染读不回 |
| **加权（作为任务面）** | 100% | **2.6** | 各项加权：3.0×25% + 2.0×15% + 2.0×15% + 2.0×15% + 2.0×10% + 4.5×10% + 4.0×5% + 2.0×5% = 2.60 |
| **作为 COM 能力底座（另一视角）** | — | **4.1** | 覆盖、测试矩阵、错误契约、预算纪律 |
| **综合当前形态** | — | **3.1** | 能用，但代价高 |

> 评分不是精确科学；它表达的是"同样的 270 个工具，换一种组织方式可以显著提高分数，而不必增加数量"。

---

## 5. 改进建议

> 约束：遵守 [HANDOFF.md](HANDOFF.md) §3.2 的 D1（广告面 ≤100 / ≤60,000 字节）、D2（Word 长尾全要）、D3（PPT 已删媒体/SmartArt/讲义/3D/美化族）、C 类（跨应用工作流/美化不进桥，走技能配方）、打印已决策不做（FIXES 1169-1171）、R14（隐藏→广告需真机轨迹证据）。

### P0：先修"说了不算 / 说了做不到"（1–2 天，不动 270 的规模）

**P0-1 修 `wps_help` 发现面**
- 去掉 60 行硬截断，或加 `offset/limit`；默认"隐藏工具优先"排序；保留 `truncated` 时给下一步提示。
- 补 9 个用户词的同义词/场景别名（多级编号、参考文献、悬挂缩进、跨页断行、相似格式、高亮、平滑、显示最终状态、引号）。
- 加"场景→工具/配方"入口：`wps_help {query:"批量合同"}` 返回 `mail_merge` + 配方名 + 前置条件。
- 验收：41 个仿真用到的隐藏业务工具 100% 可通过 app/query 找到；20 场景平均 help ≤2 次；9 个零命中词全部命中。

**P0-2 改安全默认值**
- `open_workbook.readOnly` 默认 `true`（需要写时显式 `false`）；`close_*` 的 `save` 默认 `false`。
- `wps_batch.stopOnError` 默认 `true`（或加 `dependsOnPrevious`），并在结果里给"因前置失败而跳过"的逐项状态。
- 对破坏性/批量工具加 `dryRun` 或 `preview`：返回将影响的区域/文件/行数，不落盘。
- 验收：20 场景默认路径下源文件零改动；E 风险 10 → ≤3。

**P0-3 把桥已有、schema 未暴露的高频键补进 schema（不新增广告工具）**
- 首批 10 个：条件格式 `type/value/operator/colorScaleType`、数据验证提示/警告/下拉、`set_cell_style` 自定义键、`word_insert_table.data`、`find_replace.range/sheet`、PPT `applyToAll`、PPT/Excel 的 `presentationName`/`sheet`、`word_insert_header/footer` 的 `alignment/includePageNumber`。
- 同步修 3 处描述-实现不符与 `get_document_text` 的假参数（要么实现 `start/end`，要么删掉并明说"只回前 1 万字符"）。
- 验收：10 个高影响示例全部可传；`arg-shape` 不再拒绝桥能读的键；**补 schema 会增加广告面字节（若该工具在广告面），必须同步 `gen-numbers` 复测 ≤60,000，必要时从描述/参数里腾出预算**。

**P0-4 写 20 个技能配方 + 边界清单**
- 每个配方：触发词、起手体检（`status`/`get_open_*`/`get_*_info`）、决策树（选哪个工具、为什么）、最小调用序列、验收清单、失败回退、**禁止捷径**（如 `find_replace` 做引号配对）。
- 每应用 SKILL.md 增加"明确做不到 + 替代路径"（打印、云协作、活链接、媒体等）。
- 验收：新模型按配方完成 E4/W4 不超过配方步数 1.5×；D 场景输出"不支持的根因 + 替代方案"，不再裸试 `execute_method`。

### P1：广告面换血 + 接口小改（下一个版本，1–2 周）

**P1-1 广告面按任务闭环重排（不扩容，只换血）**
- 候选晋升（按 20 场景出现频率/任务闭环）：
  - PPT：`set_active_target`、`set_slide_theme`、`set_slide_footer`、`unify_font`、`set_font_color`、`align_shapes`、`distribute_shapes`、`add_animation`、`set_slide_transition`/`apply_transition_to_all`、`get_open_presentations`。
  - Word：`accept_revisions`、`delete_comment`、`enable_track_changes`、`get_header_footer`、`set_table_format`、`find_in_document`。
  - Excel：`get_pivot_tables`、`refresh_pivot_tables`、`diagnose_formula`、`lock_cells`、`set_row_height`、`text_to_columns`、`set_print_area`。
- 候选换出（20 场景零使用且非任务闭环必需，需按 R14 再取证）：`evaluate_formula`、`goal_seek`、`export_range_as_image`、`clear_pivot_table`、`add_list_row`、`get_named_ranges`、`set_chart_labels`、`sort_range`、`unmerge_cells` 等。
- 硬约束：≤100 工具 / ≤60,000 字节；每次换血同步 `gen-numbers` 并复测；不接受"先加后删"的临时超限。
- 验收：20 场景 help 总数下降 ≥30%；PPT 改稿类场景的隐藏工具发现次数下降 ≥50%。

**P1-2 补/改高杠杆接口**（新增工具可以隐藏，不占广告位；每项先做 WPS COM 可行性实测）
- Excel：`read_range` 加 `maxRows/startRow/truncated`；`get_charts`；`update_chart` 加 `axis:{categoryType,numberFormat,displayBlanksAs}`；`set_sheet_page_setup` 加 `firstPageNumber`；`insert_page_break`；`sort_range` 多键；`copy_range` 跨表/跨簿。
- Word：`get_document_text` 分页/修假参数；`set_paragraph` 加 `lineSpacingRule:"exact"` + `hangingIndent`；`update_style`（改样式定义：字体/段落/编号）；多级编号/题注/参考文献（若 COM 可行）；`set_table_format.headerRepeat` 支持多行；`insert_table.data`。
- PPT：`apply_animation_to_all`；`set_table_data`（二维批量写，替代 N×M 次 `set_table_cell`）；`align/distribute` 加 `relativeTo:"slide"`；`get_text_format`（字体属性回读）；`set_font_color` 加 `scope:{slides,objectTypes}`；`replace_ppt_text` 加 `scope`；平滑切换映射（若 WPS 常量存在）。
- 通用：Excel/Word 加 `set_active_target`（对齐 PPT 已有实现），或统一为 `wps_common_activate_target`。
- 验收：E6/X1 的静态部分从 C/D 降到 B；P4 250 次 → ≤10 次；W1 34 次 → ≤15 次；每项有真机测试与 schema 契约。

**P1-3 把 20 场景做成基准**
- 在 `test/skilleval/` 增加 20 个场景的 fixture、prompt、确定性验收（基于裸 COM/openpyxl/python-pptx 判定，不读插件自述）。
- 每轮指标：成功率、顶层调用、help、隐藏工具发现率、源文件改动、无依据成功声明。
- 用它作为下次广告面换血的唯一证据来源（替代"历史调用次数"单一口径）。

### P2：契约卫生与长期维护（随版本）

1. **统一基准与选择器**：全部 1-based（或全部 0-based）并在描述里写死；形状族统一 `shapeIndex|shapeName`；禁止用 `string` 承载序号；`name` 按对象域改名（`workbookName`/`sheetName`/`presentationName`/`shapeName`/`bookmarkName`）。
2. **返回契约**：数据读取统一 `{data, rows, cols, truncated}`；`wps_batch` 对超 2,000 字符的结果标 `truncated:true` 并提示单独读取；加 `outputSchema`（至少对 30 个高频读取工具）。
3. **effect 元数据**：7 个写工具（`add_list_row`、`resize_list_object`、`set_list_object_totals`、`unlist_list_object`、`update_list_object`、两个 `find_replace`）错标 `read`，修正（仅元数据，不影响运行时）。
4. **重复工具收敛**：按 §3.9 的合并清单分批做；Word 长尾不删能力；保留旧名必须走显式弃用周期（D4 的别名债不能靠"悄悄加回"绕过）。
5. **命名卫生**：6 个 snake_case 参数（`target_cell`、`include_header`、`color_scheme`、`beautify_all`、`new_paragraph`、`start_offset`）、10 个 `_ppt_` 冗余名、6 处 snake_case/枚举只写在描述里的问题。

### 5.x 明确不要做（避免把建议带偏）

| 不要做 | 原因 |
| --- | --- |
| 为了"减少工具数"删桥 action | 隐藏工具不占上下文；删 action 会留下"没有工具驱动"的台账（FIXES 48 教训） |
| 加 `wps_common_print` / 打印预览 | 已决策不做（物理副作用/模态框）；用"设置好 + 导 PDF 由人打印" |
| 加跨应用活链接/更新链接进桥 | C 类跨应用工作流不进桥；WPS 放映态本身不刷新 |
| 加文件级"批量导出 PDF"大工具进桥 | 属跨应用编排，走技能配方 + `wps_batch`；避免再开一个"任务级巨工具" |
| 加媒体/SmartArt/讲义/3D 能力 | D3 已删；不要回潮 |
| 无证据地扩广告面（+N 个） | D1 上限 100，当前 95；R14 要求真机轨迹证据 |
| 把"审计/排版美化"类做成桥工具 | C 类；放技能配方（如 `wps_ppt_audit_layout` 应是 recipe，不是 tool） |

---

## 6. 20 场景 → 配方清单（B 路径的落点）

| id | 建议配方名 | 关键工具 | 验收点 |
| --- | --- | --- | --- |
| E1 | `merge-workbooks` | `get_open_workbooks`/`read_range`/`write_range`/`remove_duplicates`/`create_list_object`/`save_as` | 列对齐规则、源文件只读、静态 vs 可刷新说明 |
| E2 | `conditional-format-rules` | `get_conditional_formats`/`set_conditional_format`（补 type/公式后） | 规则类型、整行作用域、空行不触发 |
| E3 | `lookup-and-reconcile` | `set_formula`/`copy_range`/`get_formula`/`diagnose_formula` | 文本-数字差异、绝对引用、未匹配清单 |
| E4 | `clean-messy-list` | `read_range`/`text_to_columns`/`remove_duplicates`/`clean_data`/`delete_rows` | 拆列到行 vs 到列、保留首次、空行 |
| E5 | `pivot-summary` | `get_pivot_tables`/`create_pivot_table`/`refresh_pivot_tables`/`update_pivot_table` | 日期分组、占比、刷新 |
| E6 | `print-ready-sheet` | `get_sheet_settings`/`set_sheet_page_setup`/`set_sheet_print_titles`/`set_sheet_header_footer`/`convert_to_pdf` | 一页宽、标题行、起始页码（缺，需说明） |
| E7 | `chart-diagnose` | `read_range`/`get_cell_info`/`update_chart`/`export_chart_as_image` | 日期轴/空值（缺，替代：改数据源/导图） |
| E8 | `protect-shared-sheet` | `protect_sheet`/`lock_cells`/`unprotect_sheet` | 能做什么/不能做什么（云协作边界） |
| W1 | `format-long-doc` | `get_paragraphs`/`set_font`/`set_paragraph`/`apply_style`/`insert_page_numbers` | 样式/编号/参考文献/引号（缺项列清单） |
| W2 | `toc-and-page-numbers` | `get_paragraphs`/`generate_toc`/`insert_page_numbers`/`insert_section_break` | 分节、罗马/阿拉伯、起始值限制 |
| W3 | `mail-merge-batch` | `mail_merge` + CSV 生成 + `convert_to_pdf` | CSV/占位符/一人一份/一对多边界 |
| W4 | `clean-revisions` | `get_revisions`/`accept_revisions`/`get_comments`/`delete_comment`/`enable_track_changes` | 清高亮/最终状态（缺，需说明） |
| W5 | `table-page-layout` | `get_tables`/`set_table_format`/`set_paragraph` | 表头重复（单行）、跨页断行（缺） |
| W6 | `unify-similar-format` | `get_paragraphs`/`find_in_document`/`set_font`/`apply_style` | 按文本模式圈定、逐段写、验证 |
| P1 | `apply-deck-template` | `set_active_target`/`set_slide_theme`/`unify_font`/`set_slide_footer`/`get_shapes` | 套模板后逐页体检、LOGO 定位风险 |
| P2 | `arrange-slide-shapes` | `get_shapes`/`align_shapes`/`distribute_shapes`/`group_shapes`/`set_shape_position` | 相对页面居中要手工算；≥3 对象 |
| P3 | `unify-animations` | `set_active_target`/`add_animation`/`apply_transition_to_all`/`get_animations` | 平滑不可用、preset 会给所有形状加动画 |
| P4 | `change-font-color-only` | `get_shapes`/`set_font_color`/`replace_ppt_text`/`get_textboxes` | 对象类型过滤、回读（缺，需导出图片） |
| X1 | `excel-to-deck-static` | `read_range`/`insert_ppt_chart`/`set_ppt_chart_data`/`export_slide_as_image` | 一次性写入、不可刷新、单位统一 |
| F1 | `batch-export-pdf` | `open_*`/`convert_to_pdf`/`close_*` + `wps_batch` | `appType` 必传、`save:false`、四批/失败清单；打印不做 |

> 配方应放在技能 `recipes/` 下，而不是桥里。现有 `skills/*/recipes/` 已有 5 个，可扩展成 20 个，并在每个 SKILL.md 的"按需参考"里加索引。

---

## 7. 下一步验证（真机 + 基准）

**必须真机验证的 12 项**（纸面无法定论；建议抽 6 项先跑，每项 1–2 轮，遵守单实例租约）：

1. `close_*.save` 默认 true 的实际行为：只读打开的文件关闭时是否回写/弹框/报错。
2. `open_workbook.readOnly` 默认 false 的实际后果（多文件批量场景）。
3. `set_formula` 多格 range 广播后的相对引用与性能（500 行）。
4. `copy_range` 是否按目标调整相对引用。
5. `convert_to_pdf` 对 Excel 的导出页数/是否含全部 sheet；PDF 合并是否只能人工。
6. `get_document_text` 的 `start/end` 是否所有入口都忽略（源码确认忽略）。
7. `set_table_format.headerRepeat` 对 2 行表头的实际效果（源码只写第 1 行）。
8. `align_shapes`/`distribute_shapes`/`group_shapes` 传整数索引给 `Shapes.Range` 是否可行。
9. `add_master_element` 的实际报错（参数白名单 vs 类型分支）；`unify_font` 对表格/组合/图表文字覆盖。
10. "平滑/morph" 的 `EntryEffect` 常量是否存在；`set_slide_footer` 在无页脚占位版式上是否可见。
11. `get_shapes.type` 数值与 MSO 常量映射；`set_font_color` 对混排 run 的效果。
12. WPS 版本差异：日期轴/空值选项、`Adjustments` 空壳等。

**基准化**：把 20 场景固化成 `test/skilleval`，至少量 6 个指标（成功率/调用数/help/隐藏发现/源文件改动/诚实报告）。下一次广告面换血时，用这份基准的真机轨迹做 R14 证据，而不是只看历史调用次数。

---

## 附录 A：20 个场景的来源（抽样核验 16/16）

> 以下均为调研时 `web_fetch` 打开并核对的页面；知乎/B站/小红书/百度知道因 404/登录墙未作为主来源。完整清单与打开记录见过程产物 `research-excel-word.md` / `research-ppt-cross.md`。

| id | 主来源 |
| --- | --- |
| E1 | <https://bbs.wps.cn/topic/94580>、<https://bbs.wps.cn/topic/83869> |
| E2 | <http://bbs.wps.cn/topic/55896>、<https://bbs.wps.cn/topic/80663> |
| E3 | <https://bbs.wps.cn/topic/47807>、<https://bbs.wps.cn/topic/48623>、<https://www.wps.cn/article/vlookup-function-params-error-fix.html> |
| E4 | <https://bbs.wps.cn/topic/99559>、<https://bbs.wps.cn/topic/85171>、<https://bbs.wps.cn/topic/52760> |
| E5 | <https://bbs.wps.cn/topic/70245>、<https://bbs.wps.cn/topic/94580> |
| E6 | <https://bbs.wps.cn/topic/48334>、<https://bbs.wps.cn/topic/80161> |
| E7 | <https://bbs.wps.cn/topic/81936>、<https://bbs.wps.cn/topic/47346> |
| E8 | <https://bbs.wps.cn/topic/46379>、<https://bbs.wps.cn/topic/84166>、<https://bbs.wps.cn/topic/46366> |
| W1 | <https://bbs.wps.cn/topic/85526>、<https://bbs.wps.cn/topic/85441>、<https://bbs.wps.cn/topic/85351>、<https://bbs.wps.cn/topic/90405> |
| W2 | <https://bbs.wps.cn/topic/87731>、<https://bbs.wps.cn/topic/78173>、<https://bbs.wps.cn/topic/88183> |
| W3 | <http://bbs.wps.cn/topic/90404>、<https://bbs.wps.cn/topic/98583>、<http://bbs.wps.cn/topic/85263>、<http://bbs.wps.cn/topic/86939> |
| W4 | <https://bbs.wps.cn/topic/70087>、<http://bbs.wps.cn/topic/20922> |
| W5 | <https://bbs.wps.cn/topic/18055>、<https://bbs.wps.cn/topic/99573>、<https://bbs.wps.cn/topic/84107> |
| W6 | <https://bbs.wps.cn/topic/80831>、<https://bbs.wps.cn/topic/43381>、<https://bbs.wps.cn/topic/84901> |
| P1 | <https://bbs.wps.cn/topic/90006>、<https://www.wps.cn/learning/room/l/all/7231> |
| P2 | <https://bbs.wps.cn/topic/68866> |
| P3 | <https://bbs.wps.cn/topic/91262>、<https://bbs.wps.cn/topic/21313>、<https://bbs.wps.cn/topic/52997> |
| P4 | <https://baijiahao.baidu.com/s?id=1875311339880195322>、<https://bbs.wps.cn/topic/83099>、<https://www.wps.cn/learning/search/show/tab/question/keyword/ppt%20%E5%85%83%E7%B4%A0%E6%9B%BF%E6%8D%A2.html> |
| X1 | <https://bbs.wps.cn/topic/85901>、<https://bbs.wps.cn/topic/34641>、<https://bbs.wps.cn/topic/37566> |
| F1 | <https://www.chinaacc.com/wenda/detail/nofollow/xt/10302990>、<https://www.wps.cn/learning/course/detail/id/332090.html>、<https://bbs.wps.cn/topics/tag/2992?sort=hot> |

---

## 附录 B：工具使用与"未使用"清单（20 场景口径）

- **隐藏业务工具用了 41 个**（另有 2 个门面/逃生舱）。高频：`set_font_color`(250)、`set_shape_position`(31)、`add_animation`(31)、`get_header_footer`(8)、`set_active_target`(4)、`accept_revisions`/`delete_comment`/`enable_track_changes` 各 1、`set_slide_theme`/`unify_font`/`set_slide_footer` 各 1。
- **36 个广告工具在 20 场景里一次没用到**（剔除 F1 的规模放大后更明显）：`add_list_row`、`auto_fit`、`clear_formats`、`clear_pivot_table`、`clear_range`、`copy_format`、`create_chart`、`delete_rows`、`evaluate_formula`、`export_range_as_image`、`find_replace`、`freeze_panes`、`get_list_objects`、`get_named_ranges`、`goal_seek`、`set_cell_format`、`set_chart_labels`、`set_column_width`、`set_number_format`、`set_sheet_appearance`、`sort_range`、`switch_sheet`、`unmerge_cells`、`insert_ppt_image`、`insert_table`、`set_shape_fill`、`set_shape_text`、`set_slide_title`、`set_table_format`、`create_document`、`get_active_document`、`get_notes`、`get_table_data`、`insert_text`、`set_table_cell`、`smart_fill_field`。
- 读法：**"未使用"不等于"该删"**——其中很多是另一条正确路径（例如 `smart_fill_field` vs `mail_merge`、`set_cell_format` vs 更细的 setter）。它只说明"当前广告面与这 20 条真实任务的形状不匹配"，换血时需结合多场景与真机轨迹，不要按单场景缺席就换出。

---

## 附录 C：关键证据索引

| 证据 | 位置 |
| --- | --- |
| `wps_help {app}` 60 行截断 | `mcp/src/server/mcp-server.ts:379-383` |
| `wps_batch` 50 项 / 2,000 字符截断 | `mcp/src/server/mcp-server.ts:451-452`、`:473` |
| 未知参数拒绝（schema 比桥窄的直接原因） | `mcp/src/server/arg-shape.ts`、`tool-registry.ts:307-330` |
| `read_range` 返回文本表格 | `mcp/src/tools/excel/data.ts:88-108` |
| `open_workbook.readOnly` 默认 false | `mcp/scripts/wps-com.ps1:5400`、schema 描述 |
| `close_*` save 默认 true | `wps-com.ps1:5434-5437`/`5585-5588`/`6572-6574` |
| `get_document_text` 忽略 start/end、截 10,000 | `wps-com.ps1:5535-5542`、`host/wps-actions.ps1:5791-5797` |
| `set_formula` 多格广播同一公式 | `wps-com.ps1:2166-2191` |
| `convert_to_pdf` 整本导出 + appType 回退顺序 | `wps-com.ps1:1732-1763` |
| 条件格式只有 cellValue（无 xlExpression） | `wps-com.ps1:2432-2454` |
| `headerRepeat` 只写第 1 行 | `wps-com.ps1:4195` |
| PPT 切换映射无 morph | `wps-com.ps1:919-932` |
| `align_shapes` 的参照系（第一个对象/平均值） | `wps-com.ps1:7413-7452` |
| `add_master_element` 只支持 textbox/shape | `wps-com.ps1:7629-7658`、`host/wps-actions.ps1` 的 `ActionParamKeys.addMasterElement` |
| `wps_execute_method` 转发桥 action 名 | `mcp/src/client/wps-client.ts:231-241` |
| 打印/打印预览已决策不做 | `docs/FIXES.md`（1169-1171） |
| D1/D2/D3/C 类边界 | `docs/HANDOFF.md` §3.2、§8.5 |

---

## 附录 D：未入选的候选场景

| 候选 | 为什么没进 20 个 |
| --- | --- |
| X2 PPT/Word→PDF 交付（字体/清晰度/排版） | 已在 E6/W2 的导出步骤覆盖；单独列会重复"导出质量不可验证" |
| F2 多文件云协同（多人+手机端） | 与 E8 同属"本地插件边界"，保留 E8 即可；F2 更适合放在"能力边界"章节 |
| E8 云协作（已入选但仅作边界） | 用于划界，不用于要求插件实现 |
| 其他 WPS 社区高频问题（邮件合并 .et 数据源、PDF 合并、批量改名） | 与 W3/F1 重叠，保留代表场景 |

---

### 一句话收尾

**不要问"270 个是不是太多"，要问"这 270 个里，有多少个是模型在真实任务里找得到、传得对、验得了的"。**
这份评估给出的答案是：底座够深，但任务面不够像"用户"；修发现面、修默认值、补 schema、写配方、按任务闭环换广告面——这五件事不需要把 270 变成 100，就能把 0 个 A 变成一批 A。

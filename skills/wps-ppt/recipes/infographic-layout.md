# 配方：信息图（用基础形状组合）

KPI 卡片、时间线、流程图、组织架构图、仪表盘、环形图、进度条、迷你图表、标题装饰、页码指示器、配色统一
**都没有封装工具**，用基础形状搭。这样位置、配色、文案都能按场景调，而不是被写死的参数限制。

## 0. 通用做法

1. **先读画布**：`wps_ppt_get_slide_info {slideIndex:1}` → `slideWidth/slideHeight`（**磅**）。
   下面所有坐标都以磅为单位；按像素填会大 1/3。
2. **先把坐标全算出来**（写成一张 left/top/width/height 表），再逐个 `add_shape` / `add_textbox`。
3. **落完回读**：`wps_ppt_get_shapes {slideIndex}` 核对位置、尺寸与数量。
4. **导出图片核对**：`wps_ppt_export_slide_as_image` → 看有没有重叠/溢出/跑出画布。读不到图就明确说"视觉效果未验证"。

常用基础工具：`wps_ppt_add_shape`、`wps_ppt_set_shape_fill`、`wps_ppt_set_shape_text`、`wps_ppt_set_shape_position`、`wps_ppt_add_textbox`、`wps_ppt_get_shapes`、`wps_ppt_set_shape_z_order`。

记 `W = slideWidth`、`H = slideHeight`、`m = W * 0.06`（边距）、`U = W - 2m`（可用宽度）。

## 1. KPI 卡片行

- 卡宽 `c = (U - 2 * gap) / 3`，`gap = W * 0.03`；卡高 `h = H * 0.22`；`top = H * 0.38`。
- 第 i 张（0 起）：`left = m + i * (c + gap)`。
- 每张一个 `rectangle`，`set_shape_text` 写「指标名\n数值」（字号 14 / 32）。
- 用不同 `fillColor` 区分正负（如 `#28A745` / `#DC3545`）。
- 文字容量：数值字号 32 → 单行高约 45 磅；卡片高 `H*0.22`（16:9 下约 89 磅）能放下两行。

## 2. 时间线

- 轴：一条细长 `rectangle`，`left = m`，`top = H * 0.5`，`width = U`，`height = 4`，深色填充。
- 节点：每个事件一个 `oval`（直径 `W * 0.03`），`top` 与轴对齐（`top = H * 0.5 - 直径/2`），`left` 沿轴等距。
- 说明：节点下方 `add_textbox` 写日期与标题，宽度 `U / 节点数`，字号 12。
- 验收：节点在轴上居中；文字不与相邻节点文字重叠。

## 3. 流程图

- 每步一个 `rectangle`（`text` 直接写步骤名），尺寸 `c = U / 4`、`h = H * 0.12`，同 `top`。
- 步骤之间用 `arrow`（`type:"arrow"`）连接，箭头长度约 `W * 0.02`、高 `2`。
- 分支：并排放两个 `rectangle`（各 `c / 1.2`），从上游各引一根 `arrow`。
- 验收：箭头首尾相接（不重叠节点）、文字不溢出。

## 4. 组织架构图

- 根节点：`rectangle` 居中，`left = m + (U - c) / 2`，`top = H * 0.18`，`c = W * 0.22`，`h = H * 0.1`。
- 子节点：同一 `top = H * 0.45` 上等距排开，每个宽 `c2 = U / n - gap`。
- 连接：用无填充 `rectangle`（细线）或 `arrow` 从根节点下沿连到每个子节点上沿。
- 验收：连线不穿过节点文字；左右不越界。

## 5. 进度条

- 轨道：灰色 `rectangle`，`left = m`，`top = H * 0.6`，`width = U * 0.7`，`height = H * 0.04`。
- 填充：同位置同高，**宽度 = 轨道宽 × 完成比例**，强调色。
- 说明：轨道右侧 `add_textbox` 写百分比与释义。
- 注意：填充与轨道 `left`/`top` 必须完全一致，否则会出现"错位"观感。

## 6. 仪表盘 / 环形图

- 外圈：`oval`，直径 `d = H * 0.3`，居中。
- 内圈：较小的白色 `oval` 盖住中心（`d * 0.6`，`left/top` 各偏移 `d * 0.2`）。
- 中心：`add_textbox` 写百分比。
- **多段占比不要用扇形**：WPS 的扇形只暴露一个 adjustment，改角度会挂住 COM 宿主。用**横向堆叠条 + 图例**表达分布。
- 验收：内圈真的盖住外圈中心（`z_order` 正确）、数字居中。

## 7. 迷你图表（sparkline 风格）

每个指标一组文本框：数值（字号 24）、标签（字号 10）、趋势箭头 `↑` / `↓`（字号 14）。
不要用真正的图表对象——一页十几个图表对象既慢又难排版。

## 8. 标题装饰

标题下方一根细 `rectangle`：`left = m`，`top = H * 0.26`，`width = W * 0.3`，`height = 5`，强调色。
做成侧边条就把宽度收到 5、高度拉到 `H * 0.2`。

## 9. 页码指示器

右下角 `add_textbox`：`left = W - m - W * 0.12`，`top = H - H * 0.1`，`width = W * 0.12`，`height = H * 0.05`，
内容「当前页 / 总页数」（页数取自 `wps_ppt_get_slide_count`）。
**注意**：这是静态文本，插页/删页后要重新写。

## 10. 配色统一

先 `wps_ppt_get_shapes` 拿列表，再对每个形状 `wps_ppt_set_shape_fill`（标题深色、正文浅色）。
比一次性套主题更可控，也不会把不该改的形状一起改掉。
配色用固定一组（深色标题 + 中性正文 + 一个强调色），最多三种色。

也可以直接用 `wps_ppt_beautify`（`color_scheme` 取 `business/tech/creative/minimal`）；它改动范围更大，用之前先想清楚要不要人工控制。

## 11. 网格 / 自动排版

按 `rows × cols` 算好每个单元的 `left/top/width/height`（`cw = (U - (cols-1)*gap)/cols`、`ch = (Ua - (rows-1)*gap)/rows`），逐个 `set_shape_position`。
先写入一张坐标表再执行，比边算边画更容易核对。

## 12. 验收清单

- [ ] 所有形状的 left/top/width/height 都落在画布内（与 `slideWidth/slideHeight` 比）。
- [ ] `get_shapes` 回读的数量、位置与坐标表一致。
- [ ] 导出图片核对：无文字溢出、无对象重叠、边距一致、图片比例正确。
- [ ] 读不到图片时，报告里写"内容与属性已核对，视觉效果尚未验证"。

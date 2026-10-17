---
name: wps-ppt
description: 用 WPS 演示工具增删与重排幻灯片、选择版式、填占位符与文本框、添加形状与图片、组织版式、设置动画与切换、导出幻灯片为图片。
whenToUse: 任务针对演示文稿、幻灯片、页序、版式占位符、形状与文本框、图片表格、动画切换、母版或幻灯片导出时用本技能。
---

# WPS 演示

共享的执行协议（目标确认、保存策略、验证与失败分流）见 `wps-office-next` 技能；
本节是演示特有的决策与流程。**应用技能可能被单独加载**，所以关键规则在本文件里也写了一遍。

<!-- GENERATED:advertised:start -->
本应用注册 **77** 个工具，standard 档直接广告 **14** 个；
完整工具表与按任务分组索引见同目录 `reference.md`（未广告的用 `wps_call`）。
<!-- GENERATED:advertised:end -->

## 1. 先定画布，再定结构，最后版式

顺序反了会返工：

1. **定目标文稿**：同时打开多份时显式传 `presentationName`（`wps_ppt_get_open_presentations` 看清单）。
2. **读画布尺寸**：`wps_ppt_get_slide_info` 会回报 `slideWidth` / `slideHeight`（**磅**）与 `slideWidthPx` / `slideHeightPx`。
   不确定画布就套固定坐标，一定会偏。
3. **定内容结构**：几页、每页要什么（标题 / 正文 / 图 / 表 / 信息图）。
4. **定版式**：按结构选 `add_slide {layout}`，**加页时就把版式定下来**。

## 2. 单位与画布（混合单位，务必分清）

| 对象 | 单位 | 说明 |
|---|---|---|
| 形状几何：`add_shape` / `add_textbox` / `set_shape_position` / `insert_table` / `insert_ppt_image` 的 `left/top/width/height` | **磅**（point） | 直接交给 COM 的 `Shape.Left` 等属性 |
| `wps_ppt_set_slide_size` 的 `width` / `height` | **像素** | 桥内部 `×0.75` 换成磅。常用：`1280×720` 像素 = 标准宽屏 16:9（= 960×540 磅）；`960×720` 像素 = 4:3（= 720×540 磅） |
| `wps_ppt_export_slide_as_image` 的 `width` / `height` | **像素** | 默认 1280×720 |
| 字号（`fontSize`） | **磅** | |
| 动画 `duration` | **秒** | |

换算：`1 像素 = 0.75 磅`（96 dpi），即 `磅 = 像素 × 0.75`。**96dpi 下 1 英寸 = 96 像素 = 72 磅**。

也就是说：**`set_slide_size` 传像素，形状坐标传磅**。一块 1280×720 像素（标准 16:9）的画布，形状可用区域是 960×540 磅。
按"像素"给形状排版的坐标会大出 1/3，直接跑到画布外。详细换算表与版式模板见 `references/layout-and-shapes.md`。

## 3. 版式与占位符（不是每个版式都有占位符）

`add_slide` 与 `set_slide_layout` 收同一套键名。实测对应关系：

| layout 键 | 母版版式 | 页上的占位符（ppPlaceholderType） |
|---|---|---|
| `title` | 标题页 | 标题(3) + **副标题(4)** |
| `titleContent` | 标题和内容 | 标题(1) + **正文(2)** |
| `twoColumn` | 两栏内容 | 标题(1) + 表格占位符(12) |
| `comparison` | 比较 | 标题(1) + 正文(2) + 图表占位符(8) |
| `titleOnly` | 仅标题 | 标题(1) |
| `blank` | 空白 | 无 |

- `set_slide_title` 需要标题占位符（`title` / `titleContent` / `titleOnly` / `twoColumn` / `comparison` 都有）；在 `blank` 页上**明确报错** → 那一页改用 `add_textbox`。
- `set_slide_subtitle` 只在 `title` 版式可用；`set_slide_content` 只在带正文占位符的版式（如 `titleContent`）可用。用错版式会得到带"下一步"的明确失败。
- **先定版式、再填内容**；不要填完内容再换版式。
- 换版式不是万能药：对已有空白页执行 `set_slide_layout` 时 WPS 自己可能直接拒（`E_FAIL`）。要某个版式，最稳的是**加页时**就指定 `layout`。

## 4. 页序与形状索引会变

- `add_slide` 把新页追加到**末尾**；`move_slide` / `delete_slide` 之后后面的页号整体前移。
- 形状 / 文本框 / 图片的序号在插入、删除、重排之后都会变；每次操作前重新 `get_shapes` / `get_textboxes`。
- 定位形状有三种写法，语义不同：`shapeIndex` / `shapeName` / `name`（任意形状）、`imageIndex`（只数图片）、`textboxIndex`（只数文本框）。

**改过页序就必须回读终检**：`get_slide_count` 取页数 → 逐页 `get_slide_title`（或 `get_slide_info`）**按序回读标题** → 与目标顺序逐页比对 → 回报里写清"第 X 页是什么"。
每一步都会回 `success`，**看起来对不代表顺序对**。最省事的自查是"我报的页号 → 回读到的标题"能否一一对上。

页码参数统一 **1 起**，上界 = 当前页数；越界（含 0、非整数）会被桥**在动手之前**拦下。

## 5. 编辑已有页 vs 从零创建页

- **从零创建**：`create_presentation` → `add_slide {layout}` → 填标题/正文 → 形状与图片 → 保存。
  新建文稿通常是 **0 页**，必须先 `add_slide`；`Presentations.Add()` 返回 null，**不要解引用它的返回值**，用页数与形状的读操作判断状态。
- **编辑已有页**：先 `get_slide_info` 看这一页有什么（版式、形状、文本），再决定改哪个对象；不要盲目按序号写。
- 局部改文本用 `replace_ppt_text` / `set_textbox_text`；整段换内容用 `set_slide_content`（需要正文占位符）。

## 6. 排版与信息图

KPI 卡片、时间线、流程图、组织架构图、仪表盘、进度条这类信息图**没有封装工具**，用基础形状组合搭：`add_shape` + `set_shape_fill` + `set_shape_text` + `set_shape_position` + `add_textbox`。

做法：**先按画布算坐标 → 落形状 → 回读 `get_shapes` 校验位置与尺寸**。完整配方（含每种的坐标算法、文字容量、验收点）见 `recipes/infographic-layout.md`。

几条约定的边界：
- WPS 的扇形只暴露一个 adjustment，多段占比不要用扇形（改角度会挂住 COM 宿主）——用横向堆叠条 + 图例。
- 形状渐变只支持两个色标（`set_shape_effect` 的 `gradientColor1/2`），没有角度与类型参数。
- 形状效果（阴影 / 边框 / 渐变 / 透明度）用扁平参数（`shadow*` / `border*` / `gradientColor*` / `transparency`）；图片样式用嵌套 `style` 对象、背景用 `background` 对象。对象里多写 action 不认识的字段会报错；**顶层**参数名写错也会被明确拒绝。

## 7. 动画、切换与放映

- `wps_ppt_add_animation`：给 `preset` 就按整页预设给页上每个形状依次加动画；给了 `shapeIndex` / `shapeName` 就只给该形状；给 `effect` 加入场/退场，`effectKind:"emphasis"` 加强调。`trigger` 取 `onClick` / `withPrevious` / `afterPrevious`。也可以直接传 MsoAnimEffect 数值。其余动画操作（`remove_animation` / `get_animations` / `set_animation_order`）用 `wps_call`。
- 切换：`wps_ppt_set_slide_transition` 的效果名支持 `none/cut/fade/dissolve/push/wipe/split/reveal/cover/curtains`（也可传 PpEntryEffect 数值），`duration` 单位是**秒**；`apply_transition_to_all` 应用到全部；`remove_slide_transition` 移除。
- 名称写错会明确报 `unknown animation` / `unknown transition` 并列出可用名称。

## 8. 核对、保存与视觉检查

- **属性回读 ≠ 排版正确**：位置、字号、颜色都对，不代表没有文字溢出、对象重叠或裁切。
- 能力允许时先 `wps_common_save`，再 `export_slide_as_image`（绝对路径，必要时加大 `width` / `height`），然后用图像读取能力看：标题完整、文字不溢出、对象不重叠、边距一致、图片比例正确。
- **读不到导出的图片时，明确说"内容与属性已核对，视觉效果尚未验证"**，不要把"导出成功"写成"排版无误"。
- 收尾：要归档到新路径用 `wps_common_save_as`；**用户没要求关闭就不要关闭**演示文稿。

## 9. 已知限制与坑

- 目标文稿不存在时，部分底层实现会继续解引用空对象；**锁定目标后要复验**（用 `get_slide_info` 读回文稿名与页内容）。
- 某些 WPS 构建上，**新建文稿的 `Slides.Count` 可能滞后**（报 0 而页确实存在）。以 `get_slide_info` / `get_shapes` 的实际返回为准，并把构建版本记下来。
- 图表要分两步：`insert_ppt_chart` 只按 `type` 插入空图表（可给 `title`/`left`/`top`）；数据用 `set_ppt_chart_data {slideIndex, chartIndex, data}`，其中 `data` 是 `{categories:[...], series:[{name, values}]}`（**不是**顶层 `categories`/`series`，`chartIndex` 数的是第几个图表而不是形状索引）；样式用 `set_ppt_chart_style`。
- 图片路径参数一律是 **`imagePath`**（不是 `path`）；表格单元格用 `{tableIndex, row, col, text}`（`col` 不是 `column`）。
- `beautify` 的配色方案（`color_scheme`）取 `business/tech/creative/minimal`，不要与其它工具的颜色参数混用。
- 幻灯片页码只支持显示/隐藏，没有起始编号参数。
- 转换（`wps_convert_to_pdf` / `wps_convert_format`）**必须显式传 `appType:"ppt"`**。

## 10. 按需参考

| 想了解 | 读哪里 |
|---|---|
| 全部演示工具与按任务分组索引 | 同目录 `reference.md` |
| 画布 / 单位换算 / 形状与版式细则 | `references/layout-and-shapes.md` |
| 从零做一份演示的完整任务实例 | `recipes/create-deck.md` |
| 信息图配方（坐标算法、文字容量、验收点） | `recipes/infographic-layout.md` |
| 共享执行协议 | `wps-office-next` 技能 |

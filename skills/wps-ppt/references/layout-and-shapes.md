# 演示：画布、单位、版式与形状索引

只在你需要**算坐标**、**选版式**或**定位形状**时读本文件。

## 1. 先读画布，再算坐标

```
wps_ppt_get_slide_info {slideIndex:1}
→ presentationName / presentationPath / slideIndex
→ slideWidth / slideHeight          （磅）
→ slideWidthPx / slideHeightPx       （像素，= 磅 ÷ 0.75）
→ layout / shapes / shapeCount
```

**禁止在不知道画布尺寸时直接套固定坐标。** 不同演示文稿的页面尺寸可能不同（4:3、16:9、16:10、自定义）。

## 2. 单位换算表

`1 英寸 = 96 像素 = 72 磅`，所以 **`磅 = 像素 × 0.75`**、**`像素 = 磅 ÷ 0.75`**。

| 参数 | 单位 | 说明 |
|---|---|---|
| `add_shape` / `add_textbox` 的 `left/top/width/height` | **磅** | 直接写 COM 的 `Shape.Left/Top/Width/Height` |
| `set_shape_position` 的 `left/top/width/height` | **磅** | 同上；支持小数 |
| `insert_table` / `insert_ppt_image` 的 `left/top/width/height` | **磅** | 同上 |
| `set_table_format` 的 `left/top` | **磅** | |
| `set_slide_size` 的 `width/height` | **像素** | 桥内部 `×0.75` 转磅 |
| `export_slide_as_image` 的 `width/height` | **像素** | 默认 1280×720 |
| `fontSize` | **磅** | |
| `add_animation` 之外的 `duration`（切换） | **秒** | |

常见画布对照（**像素**是 `set_slide_size` 的输入，**磅**是形状坐标所在的坐标系）：

| 画布（像素，用于 `set_slide_size`） | 比例 | 形状可用区域（磅） |
|---|---|---|
| 960 × 720 | 4:3（标准 4:3） | 720 × 540 |
| 1280 × 720 | 16:9（标准宽屏） | 960 × 540 |
| 1280 × 800 | 16:10 | 960 × 600 |
| 1123 × 794 | A4 横版 | 842 × 595 |
| 794 × 1123 | A4 竖版 | 595 × 842 |

> **别把磅当像素传**：标准 16:9 的磅值是 960 × 540，而它的像素值是 1280 × 720。
> 按 `960 × 540` 传进 `set_slide_size` 会得到一个 720 × 405 磅的演示文稿 —— 比例对，但比标准小一圈。
> 实测默认文稿：`get_slide_info` 回报 `960 × 540 磅（1280 × 720 像素）`。

排版的正确姿势：拿到 `slideWidthPx/slideHeightPx` → 用**磅**画布（`slideWidth/slideHeight`）算位置 →
例如"左右各留 5% 边距"就是 `left = slideWidth * 0.05`、可用宽度 `= slideWidth * 0.9`。
按像素数值直接填给形状，会大出 1/3 并跑出画布。

## 3. 版式与占位符

| layout 键 | 母版版式 | 页上的占位符（ppPlaceholderType） |
|---|---|---|
| `title` | 标题页 | 标题(3) + 副标题(4) |
| `titleContent` | 标题和内容 | 标题(1) + 正文(2) |
| `twoColumn` | 两栏内容 | 标题(1) + 表格占位符(12) |
| `comparison` | 比较 | 标题(1) + 正文(2) + 图表占位符(8) |
| `titleOnly` | 仅标题 | 标题(1) |
| `blank` | 空白 | 无 |

- `set_slide_title`：`title` / `titleContent` / `titleOnly` / `twoColumn` / `comparison` 可用；`blank` 上明确报错 → 用 `add_textbox`。
- `set_slide_subtitle`：只 `title`。
- `set_slide_content`：只带正文占位符的版式（`titleContent`、`comparison`）。
- **先定版式再填内容**。换成版式对已有空白页可能直接 `E_FAIL`；要版式就在 `add_slide` 时给。

## 4. 形状定位的三种写法

| 写法 | 含义 |
|---|---|
| `shapeIndex` / `shapeName` / `name` | 任意形状的序号或名称 |
| `imageIndex` | **第几张图片**（只数图片） |
| `textboxIndex` | **第几个文本框**（只数文本框） |

先 `wps_ppt_get_shapes`（或 `get_textboxes`）拿列表，再选合适的那个。序号在插入、删除、重排之后会变，**每次操作前重新读**。

## 5. 版式计算模板（以磅为单位）

设画布 `W = slideWidth`、`H = slideHeight`，左右边距 `m = W * 0.06`，可用宽度 `U = W - 2m`。

| 版式 | 算法 |
|---|---|
| 单栏正文 | `left = m`，`top = H * 0.28`，`width = U`，`height = H * 0.55` |
| 两栏 | 栏宽 `c = (U - W * 0.04) / 2`；左栏 `left = m`，右栏 `left = m + c + W * 0.04` |
| 三卡片行 | 卡宽 `c = (U - 2 * gap) / 3`（`gap = W * 0.03`）；第 i 张 `left = m + i * (c + gap)` |
| 标题 + 下划装饰条 | 标题占位符在 `top = H * 0.08`；装饰条 `left = m`，`top = H * 0.26`，`width = W * 0.3`，`height = 4`（磅） |
| 页脚 / 页码指示 | `top = H - H * 0.08`，`left = W - m - W * 0.12`，`width = W * 0.12`，`height = H * 0.05` |

## 6. 文字容量（经验值，用于避免溢出）

中文字号与所需高度（磅）的粗略关系：单行高度 ≈ `fontSize * 1.4`。

| 字号 | 单行高 ≈ | 一行最多容纳 | 备注 |
|---|---|---|---|
| 40 | 56 | 宽度 ÷ 40 个字 | 大标题 |
| 28 | 39 | 宽度 ÷ 28 | 页标题 |
| 18 | 25 | 宽度 ÷ 18 | 正文 |
| 12 | 17 | 宽度 ÷ 12 | 注释 / 页脚 |

文本框默认**不自动缩放**，文字超出会溢出框外。给文本框留 20% 余量，或缩小字号。
**属性正确不代表没有溢出**——最终要靠导出图片人工核对。

## 7. 形状效果与颜色

- 颜色一律 `#RRGGBB` 十六进制字符串。
- 形状填充 `set_shape_fill {color}`；文字颜色 `set_font_color`。
- 阴影 / 边框 / 渐变 / 透明度用**扁平参数**（`shadow*` / `border*` / `gradientColor1`、`gradientColor2` / `transparency`）。
- 渐变**只支持两个色标**，没有角度与类型参数。
- 图片样式用嵌套 `style` 对象；幻灯片与母版背景用 `background` 对象（`{type:"solid",color:...}` / `{type:"gradient",colors:[...]}` / `{type:"image",imagePath:...}`）。缺色或未知类型会明确报错。
- `z_order`：`set_shape_z_order` 控制叠放次序；做"内圈盖住外圈"这类效果时需要它。

## 8. 导出与视觉检查

```
wps_common_save
wps_ppt_export_slide_as_image {slideIndex:1, outputPath:"...\\slide1.png", width:1920, height:1080}
```

导出尺寸是**像素**，加大会更利于看细节。导出后用图像读取能力核对：标题完整、文字不溢出、对象不重叠、边距一致、图片比例正确。
**读不到图片就明确说"视觉效果尚未验证"**，不要把"导出成功"当成"排版无误"。

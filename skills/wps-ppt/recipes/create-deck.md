# 配方：从零做一份演示

适用：用户给了主题/要点，要一份演示文稿；或给了数据要一页一页汇报。
目标：**页序正确、版式选对、内容装得下、导出可核对**。

## 1. 定位目标与画布

```
wps_ppt_get_open_presentations        # 多份打开时先看清
wps_ppt_create_presentation           # 新建（也可以 open_presentation {path}）
wps_ppt_get_slide_info {slideIndex:1} # presentationName / presentationPath / 画布尺寸
```

- 新建文稿通常 **0 页**，而且 `Presentations.Add()` 返回 null——不要解引用返回值，用读操作判断状态。
- 明确输出路径时，**尽早** `wps_common_save_as {outputPath}` 定下落盘位置（未落盘的新文档被要求"保存并关闭"时，工具会因为没有路径而明确报错）。
- 画布单位见 `references/layout-and-shapes.md` §2：**形状坐标是磅**。

## 2. 定页面尺寸（可选）

```
wps_ppt_set_slide_size {width:1280, height:720}    # 参数是像素
wps_ppt_get_slide_info {slideIndex:1}              # 回读 slideWidth/slideHeight（磅）
```

不传就用文稿默认尺寸；**一定要回读**，后面所有坐标都基于它。

## 3. 规划每页（先写计划再落形状）

对每一页先定：版式键、标题文案、正文要点、是否需要图/表/信息图。

```
wps_ppt_add_slide {layout:"title"}                  # 封面
wps_ppt_add_slide {layout:"titleContent"}           # 内容页
wps_ppt_add_slide {layout:"blank"}                  # 自定义排版页
```

**加页时就给 `layout`**（换版式对已有空白页可能 E_FAIL）。

然后逐页填：

```
wps_ppt_set_slide_title {slideIndex:2, title:"总体结论"}
wps_ppt_set_slide_content {slideIndex:2, content:"要点一\n要点二"}
```

- 副标题只在 `title` 版式可用；正文只在带正文占位符的版式可用（见 `references/layout-and-shapes.md` §3）。
- `blank` 页没有占位符：`set_slide_title` 会**明确报错**，改用 `add_textbox`。

## 4. 图片 / 表格 / 图表

```
wps_ppt_insert_ppt_image {slideIndex:2, imagePath:"...\\图.png", left:60, top:120, width:300, height:200}
wps_ppt_insert_table {slideIndex:3, rows:5, cols:4, left:60, top:140}
wps_ppt_set_table_cell {slideIndex:3, tableIndex:1, row:1, col:1, text:"指标"}
wps_ppt_insert_ppt_chart {slideIndex:4, type:"column_clustered", title:"趋势"}
wps_ppt_set_ppt_chart_data {slideIndex:4, chartIndex:1, data:{categories:["1月","2月"], series:[{name:"收入", values:[10,20]}]}}
```

- 图片参数是 **`imagePath`**（不是 `path`）；相对路径按当前工作目录解析，文件不存在会明确报 `image file not found`。
- 图表的**数据是一个 `data` 对象**（`{categories, series}`），而**不是**顶层 `categories` / `series`：`insert_ppt_chart` 只插空图，`set_ppt_chart_data` 按 `chartIndex`（第几个图表，不是形状索引）写数据，样式用 `set_ppt_chart_style`。
- 表格单元格是 `{tableIndex, row, col, text}`（`tableIndex` 是第几张表，`col` 不是 `column`）。
- 坐标是**磅**。图片要按比例缩放（只给一个边长时另一边取 `-1` 让它自适应）。

## 5. 页序：改过就必须回读

`add_slide` 把新页**追加到末尾**。要在中间插页，先加完再 `move_slide`，或按最终顺序依次添加。

```
wps_ppt_get_slide_count
wps_ppt_get_slide_title {slideIndex:1}   # 逐页回读
...
```

把"我报的页号 → 回读到的标题"逐页与目标顺序对齐。**每一步都会回 success，但顺序可能不对**；连续 `move_slide` 最容易把两页调换。

## 6. 页脚 / 页码 / 切换（按需）

```
wps_ppt_set_slide_footer {showSlideNumber:true, footerText:"内部资料", showFooter:true}
wps_ppt_set_slide_transition {slideIndex:1, transition:"fade", duration:0.5}
wps_ppt_apply_transition_to_all {transition:"fade"}
```

## 7. 核对与交付

| 核对项 | 判据 |
|---|---|
| 页数与页序 | 按序回读标题，与计划逐页一致 |
| 每页内容 | 标题/正文回读，无占位符残留文本 |
| 版式 | `get_slide_info` 的 `layout` 与计划一致 |
| 目标身份 | `presentationName` / `presentationPath` 与开始一致 |
| 落盘 | `wps_common_save_as` 后文件存在且可独立打开 |
| 视觉 | 导出图片后核对；读不到图就写"视觉效果未验证" |

## 8. 验收标准

- [ ] 页序回读证据与目标顺序逐页对上。
- [ ] 每页版式正确；没有把内容填进不存在的占位符。
- [ ] 图片/表格/图表位置在画布内，不重叠。
- [ ] 文件已落盘；用户没要求关闭就不关。
- [ ] 未验证项（视觉渲染、未看到的那几页）明确标注。

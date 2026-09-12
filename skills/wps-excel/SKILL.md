---
name: wps-excel
description: 用 WPS 表格工具读写区域、设置格式与公式、建图表与透视表、导出与转换。
whenToUse: 任务针对表格、工作簿、单元格区域、公式、图表、透视表或数据清洗时。
---

# WPS 表格

## 直接广告的工具

wps_excel_get_sheet_list、wps_excel_read_range、wps_excel_write_range、wps_excel_set_cell_format、wps_excel_set_formula、wps_excel_create_chart、wps_excel_create_pivot_table、wps_excel_open_workbook，以及通用的 wps_common_save、wps_convert_to_pdf。

其余 70 多个表格工具（排序、筛选、条件格式、批注、冻结窗格、命名范围、数据验证、转置、分列、分类汇总、导出图片等）通过 wps_call 使用，清单见同目录 reference.md。

## 参数约定

- 单元格定位用 {sheet, row, col}，行列都从 1 开始。
- 区域用 A1 记法字符串，如 A1:C10；不传 sheet 时用当前活动工作表。
- read_range 的 include_header 为 true 时把首行当表头返回。
- 批量写入用 wps_excel_write_range，数据是二维数组，从起始单元格向右下填充。

## 常用流程

1. wps_excel_get_open_workbooks 或 wps_excel_get_sheet_list 确认目标工作簿与工作表。
2. 写入前先 read_range 看现有内容，避免覆盖用户数据。
3. 写入用 write_range 一次提交整块数据。
4. 公式用 set_formula，单元格地址必须是有效的 A1 记法。
5. 图表与透视表的源区域必须包含表头行、数据连续无空行。
6. 写回后用 read_range 复验，再 wps_common_save。

## 已知坑

- 省略 sheet 时底层曾出现取 0 号工作表的缺陷；不确定时显式传 sheet 名。
- 全列引用如 A:A 在公式里代价高，尽量写成 A1:A1000 这类有界区域。
- generate_formula 需要 description 参数，不传会校验失败。
- 导出图片依赖剪贴板，并发或无人值守时可能失败，失败信息要如实上报。

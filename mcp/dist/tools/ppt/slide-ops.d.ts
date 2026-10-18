/**
 * Input: 幻灯片操作工具参数
 * Output: 幻灯片操作结果
 * Pos: PPT 幻灯片操作工具实现。一旦我被修改，请更新我的头部注释，并在 docs/FIXES.md 记账。
 *
 * 幻灯片操作Tools - 幻灯片生命周期（删除/移动/查询/切换/布局/备注/尺寸）
 * 与形状、文本框的创建（形状的读取/修改族在 shape-basic.ts / textbox.ts）
 *
 * 包含（15 个）：
 * - wps_ppt_delete_slide: 删除指定幻灯片
 * - wps_ppt_move_slide: 移动幻灯片到指定位置
 * - wps_ppt_get_slide_count: 获取幻灯片总数
 * - wps_ppt_get_slide_info: 获取指定幻灯片的详细信息
 * - wps_ppt_switch_slide: 切换到指定幻灯片
 * - wps_ppt_set_slide_layout: 设置幻灯片版式布局
 * - wps_ppt_get_slide_notes: 获取幻灯片备注内容
 * - wps_ppt_set_slide_notes: 设置幻灯片备注
 * - wps_ppt_add_shape: 添加形状
 * - wps_ppt_set_shape_style: 设置形状样式
 * - wps_ppt_add_textbox: 添加文本框
 * - wps_ppt_set_slide_title: 设置幻灯片标题
 * - wps_ppt_set_shape_text: 设置形状文字
 * - wps_ppt_set_slide_size: 设置幻灯片尺寸
 * - wps_ppt_set_shape_fill: 设置形状填充颜色
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
export declare const deleteSlideDefinition: ToolDefinition;
export declare const deleteSlideHandler: ToolHandler;
export declare const moveSlideDefinition: ToolDefinition;
export declare const moveSlideHandler: ToolHandler;
export declare const getSlideCountDefinition: ToolDefinition;
export declare const getSlideCountHandler: ToolHandler;
export declare const getSlideInfoDefinition: ToolDefinition;
export declare const getSlideInfoHandler: ToolHandler;
export declare const switchSlideDefinition: ToolDefinition;
export declare const switchSlideHandler: ToolHandler;
export declare const setSlideLayoutDefinition: ToolDefinition;
export declare const setSlideLayoutHandler: ToolHandler;
export declare const getSlideNotesDefinition: ToolDefinition;
export declare const getSlideNotesHandler: ToolHandler;
export declare const setSlideNotesDefinition: ToolDefinition;
export declare const setSlideNotesHandler: ToolHandler;
export declare const addShapeDefinition: ToolDefinition;
export declare const addShapeHandler: ToolHandler;
export declare const setShapeStyleDefinition: ToolDefinition;
export declare const setShapeStyleHandler: ToolHandler;
export declare const addTextboxDefinition: ToolDefinition;
export declare const addTextboxHandler: ToolHandler;
export declare const setSlideTitleDefinition: ToolDefinition;
export declare const setSlideTitleHandler: ToolHandler;
export declare const setShapeTextDefinition: ToolDefinition;
export declare const setShapeTextHandler: ToolHandler;
export declare const setSlideSizeDefinition: ToolDefinition;
export declare const setSlideSizeHandler: ToolHandler;
export declare const setShapeFillDefinition: ToolDefinition;
export declare const setShapeFillHandler: ToolHandler;
export declare const slideOpsTools: RegisteredTool[];
export default slideOpsTools;
//# sourceMappingURL=slide-ops.d.ts.map
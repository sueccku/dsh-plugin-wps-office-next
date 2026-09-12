/**
 * Input: 幻灯片操作工具参数
 * Output: 幻灯片操作结果
 * Pos: PPT 幻灯片操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 幻灯片操作Tools - 删除/复制/移动/查询/切换/布局/备注
 *
 * 包含：
 * - wps_ppt_delete_slide: 删除指定幻灯片
 * - wps_ppt_duplicate_slide: 复制幻灯片
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
 * - wps_ppt_insert_image: 插入图片
 * - wps_ppt_set_shape_text: 设置形状文字
 * - wps_ppt_set_animation: 设置元素动画
 * - wps_ppt_set_background: 设置幻灯片背景
 * - wps_ppt_set_slide_size: 设置幻灯片尺寸
 * - wps_ppt_set_transition: 设置幻灯片切换效果
 * - wps_ppt_add_chart: 在幻灯片中插入图表
 * - wps_ppt_set_shape_fill: 设置形状填充颜色
 * - wps_ppt_add_speaker_notes: 添加演讲者备注
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
export declare const deleteSlideDefinition: ToolDefinition;
export declare const deleteSlideHandler: ToolHandler;
export declare const duplicateSlideDefinition: ToolDefinition;
export declare const duplicateSlideHandler: ToolHandler;
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
export declare const insertImageDefinition: ToolDefinition;
export declare const insertImageHandler: ToolHandler;
export declare const setShapeTextDefinition: ToolDefinition;
export declare const setShapeTextHandler: ToolHandler;
export declare const setAnimationDefinition: ToolDefinition;
export declare const setAnimationHandler: ToolHandler;
export declare const setBackgroundDefinition: ToolDefinition;
export declare const setBackgroundHandler: ToolHandler;
export declare const setSlideSizeDefinition: ToolDefinition;
export declare const setSlideSizeHandler: ToolHandler;
export declare const setTransitionDefinition: ToolDefinition;
export declare const setTransitionHandler: ToolHandler;
export declare const addChartDefinition: ToolDefinition;
export declare const addChartHandler: ToolHandler;
export declare const setShapeFillDefinition: ToolDefinition;
export declare const setShapeFillHandler: ToolHandler;
export declare const addSpeakerNotesDefinition: ToolDefinition;
export declare const addSpeakerNotesHandler: ToolHandler;
export declare const slideOpsTools: RegisteredTool[];
export default slideOpsTools;
//# sourceMappingURL=slide-ops.d.ts.map
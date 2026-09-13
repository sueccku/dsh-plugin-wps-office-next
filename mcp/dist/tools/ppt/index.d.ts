/**
 * Input: PPT 工具定义
 * Output: PPT 工具注册数组
 * Pos: PPT Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT Tools入口 - PPT工具汇总模块
 *
 * 整合所有PPT相关的Tools
 * 包含：
 * - 幻灯片Tools: add_slide, beautify, unify_font
 * - 幻灯片操作Tools: delete_slide, duplicate_slide, move_slide, get_slide_count,
 *   get_slide_info, switch_slide, set_slide_layout, get_slide_notes, set_slide_notes,
 *   add_shape, set_shape_style, add_textbox, set_slide_title, insert_image, set_shape_text,
 *   set_animation, set_background, set_slide_size
 * - 演示文稿管理Tools: create_presentation, open_presentation, close_presentation, get_open_presentations, switch_presentation, copy_slide
 * - 背景与页面信息Tools: set_slide_background, set_background_color, set_background_image,
 *   set_slide_number, set_ppt_footer, set_ppt_date_time, duplicate_shape, set_shape_z_order
 * - 形状基础Tools: delete_shape, get_shapes, set_shape_position, set_shape_shadow,
 *   set_shape_gradient, set_shape_border, set_shape_transparency, align_shapes,
 *   distribute_shapes, group_shapes
 * - 图片Tools: insert_ppt_image, delete_ppt_image, set_image_style, export_slide_as_image
 * - 动画与切换Tools: add_animation, remove_animation, get_animations, set_animation_order,
 *   add_animation_preset, add_emphasis_animation, set_slide_transition, remove_slide_transition, apply_transition_to_all
 * - 图表与流程图Tools: insert_ppt_chart, set_ppt_chart_data, set_ppt_chart_style,
 * - 杂项Tools: get_slide_master, set_master_background, add_master_element,
 *   set_3d_rotation, set_3d_depth, set_3d_material,
 *   add_ppt_hyperlink, remove_ppt_hyperlink, find_ppt_text, replace_ppt_text, start_slide_show
 * - 表格Tools: insert_table, set_table_cell, get_table_cell, set_table_style, set_table_cell_style, set_table_row_style
 */
import { RegisteredTool } from '../../types/tools';
/**
 * 所有PPT相关的Tools
 * 包含：
 * - 幻灯片Tools: add_slide, beautify, unify_font
 * - 幻灯片操作Tools: delete_slide, duplicate_slide, move_slide, get_slide_count,
 *   get_slide_info, switch_slide, set_slide_layout, get_slide_notes, set_slide_notes
 * - 演示文稿管理Tools: create_presentation, open_presentation, close_presentation, get_open_presentations, switch_presentation
 * - 表格Tools: insert_table, set_table_cell, get_table_cell, set_table_style, set_table_cell_style, set_table_row_style
 * - 高级美化Tools: apply_color_scheme, auto_beautify_slide, beautify_all_slides, create_kpi_cards, create_styled_table, add_title_decoration, add_page_indicator, set_background_gradient
 */
export declare const pptTools: RegisteredTool[];
export { slideTools } from './slide';
export { slideOpsTools } from './slide-ops';
export { presentationTools } from './presentation';
export { textboxTools } from './textbox';
export { backgroundTools } from './background';
export { shapeBasicTools } from './shape-basic';
export { imageTools } from './image';
export { animationTools } from './animation';
export { chartFlowTools } from './chart-flow';
export { miscTools } from './misc';
export { tableTools } from './table';
export { beautifyAdvancedTools } from './beautify-advanced';
export { addSlideDefinition, addSlideHandler, beautifyDefinition, beautifyHandler, unifyFontDefinition, unifyFontHandler, setFontColorDefinition, setFontColorHandler, } from './slide';
export { deleteSlideDefinition, deleteSlideHandler, moveSlideDefinition, moveSlideHandler, getSlideCountDefinition, getSlideCountHandler, getSlideInfoDefinition, getSlideInfoHandler, switchSlideDefinition, switchSlideHandler, setSlideLayoutDefinition, setSlideLayoutHandler, getSlideNotesDefinition, getSlideNotesHandler, setSlideNotesDefinition, setSlideNotesHandler, addShapeDefinition, addShapeHandler, setShapeStyleDefinition, setShapeStyleHandler, addTextboxDefinition, addTextboxHandler, setSlideTitleDefinition, setSlideTitleHandler, setShapeTextDefinition, setShapeTextHandler, setSlideSizeDefinition, setSlideSizeHandler, setShapeFillDefinition, setShapeFillHandler, } from './slide-ops';
export { createPresentationDefinition, createPresentationHandler, openPresentationDefinition, openPresentationHandler, closePresentationDefinition, closePresentationHandler, getOpenPresentationsDefinition, getOpenPresentationsHandler, switchPresentationDefinition, switchPresentationHandler, copySlideDefinition, copySlideHandler, insertSlidesFromFileDefinition, insertSlidesFromFileHandler, setActiveTargetDefinition, setActiveTargetHandler, setSlideThemeDefinition, setSlideThemeHandler, } from './presentation';
export { deleteTextboxDefinition, deleteTextboxHandler, getTextboxesDefinition, getTextboxesHandler, setTextboxTextDefinition, setTextboxTextHandler, setTextboxStyleDefinition, setTextboxStyleHandler, getSlideTitleDefinition, getSlideTitleHandler, setSlideSubtitleDefinition, setSlideSubtitleHandler, setSlideContentDefinition, setSlideContentHandler, } from './textbox';
export { setSlideBackgroundDefinition, setSlideBackgroundHandler, setBackgroundColorDefinition, setBackgroundColorHandler, setBackgroundImageDefinition, setBackgroundImageHandler, setSlideNumberDefinition, setSlideNumberHandler, setPptFooterDefinition, setPptFooterHandler, setPptDateTimeDefinition, setPptDateTimeHandler, duplicateShapeDefinition, duplicateShapeHandler, setShapeZOrderDefinition, setShapeZOrderHandler, } from './background';
export { deleteShapeDefinition, deleteShapeHandler, getShapesDefinition, getShapesHandler, setShapePositionDefinition, setShapePositionHandler, setShapeShadowDefinition, setShapeShadowHandler, setShapeGradientDefinition, setShapeGradientHandler, setShapeBorderDefinition, setShapeBorderHandler, setShapeTransparencyDefinition, setShapeTransparencyHandler, alignShapesDefinition, alignShapesHandler, distributeShapesDefinition, distributeShapesHandler, groupShapesDefinition, groupShapesHandler, } from './shape-basic';
export { insertPptImageDefinition, insertPptImageHandler, deletePptImageDefinition, deletePptImageHandler, setImageStyleDefinition, setImageStyleHandler, exportSlideAsImageDefinition, exportSlideAsImageHandler, replacePptImageDefinition, replacePptImageHandler, } from './image';
export { addAnimationDefinition, addAnimationHandler, removeAnimationDefinition, removeAnimationHandler, getAnimationsDefinition, getAnimationsHandler, setAnimationOrderDefinition, setAnimationOrderHandler, addAnimationPresetDefinition, addAnimationPresetHandler, addEmphasisAnimationDefinition, addEmphasisAnimationHandler, setSlideTransitionDefinition, setSlideTransitionHandler, removeSlideTransitionDefinition, removeSlideTransitionHandler, applyTransitionToAllDefinition, applyTransitionToAllHandler, } from './animation';
export { insertPptChartDefinition, insertPptChartHandler, setPptChartDataDefinition, setPptChartDataHandler, setPptChartStyleDefinition, setPptChartStyleHandler, } from './chart-flow';
export { getSlideMasterDefinition, getSlideMasterHandler, setMasterBackgroundDefinition, setMasterBackgroundHandler, addMasterElementDefinition, addMasterElementHandler, set3DRotationDefinition, set3DRotationHandler, set3DDepthDefinition, set3DDepthHandler, set3DMaterialDefinition, set3DMaterialHandler, addPptHyperlinkDefinition, addPptHyperlinkHandler, removePptHyperlinkDefinition, removePptHyperlinkHandler, findPptTextDefinition, findPptTextHandler, replacePptTextDefinition, replacePptTextHandler, startSlideShowDefinition, startSlideShowHandler, } from './misc';
export { insertPptTableDefinition, insertPptTableHandler, setPptTableCellDefinition, setPptTableCellHandler, getPptTableCellDefinition, getPptTableCellHandler, setPptTableStyleDefinition, setPptTableStyleHandler, setPptTableCellStyleDefinition, setPptTableCellStyleHandler, setPptTableRowStyleDefinition, setPptTableRowStyleHandler, } from './table';
export { setBackgroundGradientDefinition, setBackgroundGradientHandler, } from './beautify-advanced';
export default pptTools;
//# sourceMappingURL=index.d.ts.map
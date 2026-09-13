"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.insertImageDefinition = exports.setSlideTitleHandler = exports.setSlideTitleDefinition = exports.addTextboxHandler = exports.addTextboxDefinition = exports.setShapeStyleHandler = exports.setShapeStyleDefinition = exports.addShapeHandler = exports.addShapeDefinition = exports.setSlideNotesHandler = exports.setSlideNotesDefinition = exports.getSlideNotesHandler = exports.getSlideNotesDefinition = exports.setSlideLayoutHandler = exports.setSlideLayoutDefinition = exports.switchSlideHandler = exports.switchSlideDefinition = exports.getSlideInfoHandler = exports.getSlideInfoDefinition = exports.getSlideCountHandler = exports.getSlideCountDefinition = exports.moveSlideHandler = exports.moveSlideDefinition = exports.duplicateSlideHandler = exports.duplicateSlideDefinition = exports.deleteSlideHandler = exports.deleteSlideDefinition = exports.setFontColorHandler = exports.setFontColorDefinition = exports.alignObjectsHandler = exports.alignObjectsDefinition = exports.unifyFontHandler = exports.unifyFontDefinition = exports.beautifyHandler = exports.beautifyDefinition = exports.addSlideHandler = exports.addSlideDefinition = exports.beautifyAdvancedTools = exports.tableTools = exports.miscTools = exports.chartFlowTools = exports.animationTools = exports.imageTools = exports.shapeBasicTools = exports.backgroundTools = exports.textboxTools = exports.presentationTools = exports.slideOpsTools = exports.slideTools = exports.pptTools = void 0;
exports.setSlideBackgroundDefinition = exports.setSlideContentHandler = exports.setSlideContentDefinition = exports.setSlideSubtitleHandler = exports.setSlideSubtitleDefinition = exports.getSlideTitleHandler = exports.getSlideTitleDefinition = exports.setTextboxStyleHandler = exports.setTextboxStyleDefinition = exports.setTextboxTextHandler = exports.setTextboxTextDefinition = exports.getTextboxesHandler = exports.getTextboxesDefinition = exports.deleteTextboxHandler = exports.deleteTextboxDefinition = exports.setActiveTargetHandler = exports.setActiveTargetDefinition = exports.insertSlidesFromFileHandler = exports.insertSlidesFromFileDefinition = exports.insertSlideImageHandler = exports.insertSlideImageDefinition = exports.copySlideHandler = exports.copySlideDefinition = exports.switchPresentationHandler = exports.switchPresentationDefinition = exports.getOpenPresentationsHandler = exports.getOpenPresentationsDefinition = exports.closePresentationHandler = exports.closePresentationDefinition = exports.openPresentationHandler = exports.openPresentationDefinition = exports.createPresentationHandler = exports.createPresentationDefinition = exports.setTransitionHandler = exports.setTransitionDefinition = exports.setShapeFillHandler = exports.setShapeFillDefinition = exports.addSpeakerNotesHandler = exports.addSpeakerNotesDefinition = exports.addChartHandler = exports.addChartDefinition = exports.setSlideSizeHandler = exports.setSlideSizeDefinition = exports.setBackgroundHandler = exports.setBackgroundDefinition = exports.setAnimationHandler = exports.setAnimationDefinition = exports.setShapeTextHandler = exports.setShapeTextDefinition = exports.insertImageHandler = void 0;
exports.getAnimationsDefinition = exports.removeAnimationHandler = exports.removeAnimationDefinition = exports.addAnimationHandler = exports.addAnimationDefinition = exports.replacePptImageHandler = exports.replacePptImageDefinition = exports.exportSlideAsImageHandler = exports.exportSlideAsImageDefinition = exports.setImageStyleHandler = exports.setImageStyleDefinition = exports.deletePptImageHandler = exports.deletePptImageDefinition = exports.insertPptImageHandler = exports.insertPptImageDefinition = exports.groupShapesHandler = exports.groupShapesDefinition = exports.distributeShapesHandler = exports.distributeShapesDefinition = exports.alignShapesHandler = exports.alignShapesDefinition = exports.setShapeTransparencyHandler = exports.setShapeTransparencyDefinition = exports.setShapeBorderHandler = exports.setShapeBorderDefinition = exports.setShapeGradientHandler = exports.setShapeGradientDefinition = exports.setShapeShadowHandler = exports.setShapeShadowDefinition = exports.setShapePositionHandler = exports.setShapePositionDefinition = exports.getShapesHandler = exports.getShapesDefinition = exports.deleteShapeHandler = exports.deleteShapeDefinition = exports.setShapeZOrderHandler = exports.setShapeZOrderDefinition = exports.duplicateShapeHandler = exports.duplicateShapeDefinition = exports.setPptDateTimeHandler = exports.setPptDateTimeDefinition = exports.setPptFooterHandler = exports.setPptFooterDefinition = exports.setSlideNumberHandler = exports.setSlideNumberDefinition = exports.setBackgroundImageHandler = exports.setBackgroundImageDefinition = exports.setBackgroundColorHandler = exports.setBackgroundColorDefinition = exports.setSlideBackgroundHandler = void 0;
exports.setPptTableCellStyleDefinition = exports.setPptTableStyleHandler = exports.setPptTableStyleDefinition = exports.getPptTableCellHandler = exports.getPptTableCellDefinition = exports.setPptTableCellHandler = exports.setPptTableCellDefinition = exports.insertPptTableHandler = exports.insertPptTableDefinition = exports.startSlideShowHandler = exports.startSlideShowDefinition = exports.replacePptTextHandler = exports.replacePptTextDefinition = exports.findPptTextHandler = exports.findPptTextDefinition = exports.removePptHyperlinkHandler = exports.removePptHyperlinkDefinition = exports.addPptHyperlinkHandler = exports.addPptHyperlinkDefinition = exports.set3DMaterialHandler = exports.set3DMaterialDefinition = exports.set3DDepthHandler = exports.set3DDepthDefinition = exports.set3DRotationHandler = exports.set3DRotationDefinition = exports.addMasterElementHandler = exports.addMasterElementDefinition = exports.setMasterBackgroundHandler = exports.setMasterBackgroundDefinition = exports.getSlideMasterHandler = exports.getSlideMasterDefinition = exports.setPptChartStyleHandler = exports.setPptChartStyleDefinition = exports.setPptChartDataHandler = exports.setPptChartDataDefinition = exports.insertPptChartHandler = exports.insertPptChartDefinition = exports.applyTransitionToAllHandler = exports.applyTransitionToAllDefinition = exports.removeSlideTransitionHandler = exports.removeSlideTransitionDefinition = exports.setSlideTransitionHandler = exports.setSlideTransitionDefinition = exports.addEmphasisAnimationHandler = exports.addEmphasisAnimationDefinition = exports.addAnimationPresetHandler = exports.addAnimationPresetDefinition = exports.setAnimationOrderHandler = exports.setAnimationOrderDefinition = exports.getAnimationsHandler = void 0;
exports.setBackgroundGradientHandler = exports.setBackgroundGradientDefinition = exports.setPptTableRowStyleHandler = exports.setPptTableRowStyleDefinition = exports.setPptTableCellStyleHandler = void 0;
const slide_1 = require("./slide");
const slide_ops_1 = require("./slide-ops");
const presentation_1 = require("./presentation");
const textbox_1 = require("./textbox");
const background_1 = require("./background");
const shape_basic_1 = require("./shape-basic");
const image_1 = require("./image");
const animation_1 = require("./animation");
const chart_flow_1 = require("./chart-flow");
const misc_1 = require("./misc");
const table_1 = require("./table");
const beautify_advanced_1 = require("./beautify-advanced");
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
exports.pptTools = [
    ...slide_1.slideTools,
    ...slide_ops_1.slideOpsTools,
    ...presentation_1.presentationTools,
    ...textbox_1.textboxTools,
    ...background_1.backgroundTools,
    ...shape_basic_1.shapeBasicTools,
    ...image_1.imageTools,
    ...animation_1.animationTools,
    ...chart_flow_1.chartFlowTools,
    ...misc_1.miscTools,
    ...table_1.tableTools,
    ...beautify_advanced_1.beautifyAdvancedTools,
];
// 分别导出，方便按需使用
var slide_2 = require("./slide");
Object.defineProperty(exports, "slideTools", { enumerable: true, get: function () { return slide_2.slideTools; } });
var slide_ops_2 = require("./slide-ops");
Object.defineProperty(exports, "slideOpsTools", { enumerable: true, get: function () { return slide_ops_2.slideOpsTools; } });
var presentation_2 = require("./presentation");
Object.defineProperty(exports, "presentationTools", { enumerable: true, get: function () { return presentation_2.presentationTools; } });
var textbox_2 = require("./textbox");
Object.defineProperty(exports, "textboxTools", { enumerable: true, get: function () { return textbox_2.textboxTools; } });
var background_2 = require("./background");
Object.defineProperty(exports, "backgroundTools", { enumerable: true, get: function () { return background_2.backgroundTools; } });
var shape_basic_2 = require("./shape-basic");
Object.defineProperty(exports, "shapeBasicTools", { enumerable: true, get: function () { return shape_basic_2.shapeBasicTools; } });
var image_2 = require("./image");
Object.defineProperty(exports, "imageTools", { enumerable: true, get: function () { return image_2.imageTools; } });
var animation_2 = require("./animation");
Object.defineProperty(exports, "animationTools", { enumerable: true, get: function () { return animation_2.animationTools; } });
var chart_flow_2 = require("./chart-flow");
Object.defineProperty(exports, "chartFlowTools", { enumerable: true, get: function () { return chart_flow_2.chartFlowTools; } });
var misc_2 = require("./misc");
Object.defineProperty(exports, "miscTools", { enumerable: true, get: function () { return misc_2.miscTools; } });
var table_2 = require("./table");
Object.defineProperty(exports, "tableTools", { enumerable: true, get: function () { return table_2.tableTools; } });
var beautify_advanced_2 = require("./beautify-advanced");
Object.defineProperty(exports, "beautifyAdvancedTools", { enumerable: true, get: function () { return beautify_advanced_2.beautifyAdvancedTools; } });
// 导出单独的定义和处理器，方便测试
var slide_3 = require("./slide");
Object.defineProperty(exports, "addSlideDefinition", { enumerable: true, get: function () { return slide_3.addSlideDefinition; } });
Object.defineProperty(exports, "addSlideHandler", { enumerable: true, get: function () { return slide_3.addSlideHandler; } });
Object.defineProperty(exports, "beautifyDefinition", { enumerable: true, get: function () { return slide_3.beautifyDefinition; } });
Object.defineProperty(exports, "beautifyHandler", { enumerable: true, get: function () { return slide_3.beautifyHandler; } });
Object.defineProperty(exports, "unifyFontDefinition", { enumerable: true, get: function () { return slide_3.unifyFontDefinition; } });
Object.defineProperty(exports, "unifyFontHandler", { enumerable: true, get: function () { return slide_3.unifyFontHandler; } });
Object.defineProperty(exports, "alignObjectsDefinition", { enumerable: true, get: function () { return slide_3.alignObjectsDefinition; } });
Object.defineProperty(exports, "alignObjectsHandler", { enumerable: true, get: function () { return slide_3.alignObjectsHandler; } });
Object.defineProperty(exports, "setFontColorDefinition", { enumerable: true, get: function () { return slide_3.setFontColorDefinition; } });
Object.defineProperty(exports, "setFontColorHandler", { enumerable: true, get: function () { return slide_3.setFontColorHandler; } });
var slide_ops_3 = require("./slide-ops");
Object.defineProperty(exports, "deleteSlideDefinition", { enumerable: true, get: function () { return slide_ops_3.deleteSlideDefinition; } });
Object.defineProperty(exports, "deleteSlideHandler", { enumerable: true, get: function () { return slide_ops_3.deleteSlideHandler; } });
Object.defineProperty(exports, "duplicateSlideDefinition", { enumerable: true, get: function () { return slide_ops_3.duplicateSlideDefinition; } });
Object.defineProperty(exports, "duplicateSlideHandler", { enumerable: true, get: function () { return slide_ops_3.duplicateSlideHandler; } });
Object.defineProperty(exports, "moveSlideDefinition", { enumerable: true, get: function () { return slide_ops_3.moveSlideDefinition; } });
Object.defineProperty(exports, "moveSlideHandler", { enumerable: true, get: function () { return slide_ops_3.moveSlideHandler; } });
Object.defineProperty(exports, "getSlideCountDefinition", { enumerable: true, get: function () { return slide_ops_3.getSlideCountDefinition; } });
Object.defineProperty(exports, "getSlideCountHandler", { enumerable: true, get: function () { return slide_ops_3.getSlideCountHandler; } });
Object.defineProperty(exports, "getSlideInfoDefinition", { enumerable: true, get: function () { return slide_ops_3.getSlideInfoDefinition; } });
Object.defineProperty(exports, "getSlideInfoHandler", { enumerable: true, get: function () { return slide_ops_3.getSlideInfoHandler; } });
Object.defineProperty(exports, "switchSlideDefinition", { enumerable: true, get: function () { return slide_ops_3.switchSlideDefinition; } });
Object.defineProperty(exports, "switchSlideHandler", { enumerable: true, get: function () { return slide_ops_3.switchSlideHandler; } });
Object.defineProperty(exports, "setSlideLayoutDefinition", { enumerable: true, get: function () { return slide_ops_3.setSlideLayoutDefinition; } });
Object.defineProperty(exports, "setSlideLayoutHandler", { enumerable: true, get: function () { return slide_ops_3.setSlideLayoutHandler; } });
Object.defineProperty(exports, "getSlideNotesDefinition", { enumerable: true, get: function () { return slide_ops_3.getSlideNotesDefinition; } });
Object.defineProperty(exports, "getSlideNotesHandler", { enumerable: true, get: function () { return slide_ops_3.getSlideNotesHandler; } });
Object.defineProperty(exports, "setSlideNotesDefinition", { enumerable: true, get: function () { return slide_ops_3.setSlideNotesDefinition; } });
Object.defineProperty(exports, "setSlideNotesHandler", { enumerable: true, get: function () { return slide_ops_3.setSlideNotesHandler; } });
Object.defineProperty(exports, "addShapeDefinition", { enumerable: true, get: function () { return slide_ops_3.addShapeDefinition; } });
Object.defineProperty(exports, "addShapeHandler", { enumerable: true, get: function () { return slide_ops_3.addShapeHandler; } });
Object.defineProperty(exports, "setShapeStyleDefinition", { enumerable: true, get: function () { return slide_ops_3.setShapeStyleDefinition; } });
Object.defineProperty(exports, "setShapeStyleHandler", { enumerable: true, get: function () { return slide_ops_3.setShapeStyleHandler; } });
Object.defineProperty(exports, "addTextboxDefinition", { enumerable: true, get: function () { return slide_ops_3.addTextboxDefinition; } });
Object.defineProperty(exports, "addTextboxHandler", { enumerable: true, get: function () { return slide_ops_3.addTextboxHandler; } });
Object.defineProperty(exports, "setSlideTitleDefinition", { enumerable: true, get: function () { return slide_ops_3.setSlideTitleDefinition; } });
Object.defineProperty(exports, "setSlideTitleHandler", { enumerable: true, get: function () { return slide_ops_3.setSlideTitleHandler; } });
Object.defineProperty(exports, "insertImageDefinition", { enumerable: true, get: function () { return slide_ops_3.insertImageDefinition; } });
Object.defineProperty(exports, "insertImageHandler", { enumerable: true, get: function () { return slide_ops_3.insertImageHandler; } });
Object.defineProperty(exports, "setShapeTextDefinition", { enumerable: true, get: function () { return slide_ops_3.setShapeTextDefinition; } });
Object.defineProperty(exports, "setShapeTextHandler", { enumerable: true, get: function () { return slide_ops_3.setShapeTextHandler; } });
Object.defineProperty(exports, "setAnimationDefinition", { enumerable: true, get: function () { return slide_ops_3.setAnimationDefinition; } });
Object.defineProperty(exports, "setAnimationHandler", { enumerable: true, get: function () { return slide_ops_3.setAnimationHandler; } });
Object.defineProperty(exports, "setBackgroundDefinition", { enumerable: true, get: function () { return slide_ops_3.setBackgroundDefinition; } });
Object.defineProperty(exports, "setBackgroundHandler", { enumerable: true, get: function () { return slide_ops_3.setBackgroundHandler; } });
Object.defineProperty(exports, "setSlideSizeDefinition", { enumerable: true, get: function () { return slide_ops_3.setSlideSizeDefinition; } });
Object.defineProperty(exports, "setSlideSizeHandler", { enumerable: true, get: function () { return slide_ops_3.setSlideSizeHandler; } });
Object.defineProperty(exports, "addChartDefinition", { enumerable: true, get: function () { return slide_ops_3.addChartDefinition; } });
Object.defineProperty(exports, "addChartHandler", { enumerable: true, get: function () { return slide_ops_3.addChartHandler; } });
Object.defineProperty(exports, "addSpeakerNotesDefinition", { enumerable: true, get: function () { return slide_ops_3.addSpeakerNotesDefinition; } });
Object.defineProperty(exports, "addSpeakerNotesHandler", { enumerable: true, get: function () { return slide_ops_3.addSpeakerNotesHandler; } });
Object.defineProperty(exports, "setShapeFillDefinition", { enumerable: true, get: function () { return slide_ops_3.setShapeFillDefinition; } });
Object.defineProperty(exports, "setShapeFillHandler", { enumerable: true, get: function () { return slide_ops_3.setShapeFillHandler; } });
Object.defineProperty(exports, "setTransitionDefinition", { enumerable: true, get: function () { return slide_ops_3.setTransitionDefinition; } });
Object.defineProperty(exports, "setTransitionHandler", { enumerable: true, get: function () { return slide_ops_3.setTransitionHandler; } });
var presentation_3 = require("./presentation");
Object.defineProperty(exports, "createPresentationDefinition", { enumerable: true, get: function () { return presentation_3.createPresentationDefinition; } });
Object.defineProperty(exports, "createPresentationHandler", { enumerable: true, get: function () { return presentation_3.createPresentationHandler; } });
Object.defineProperty(exports, "openPresentationDefinition", { enumerable: true, get: function () { return presentation_3.openPresentationDefinition; } });
Object.defineProperty(exports, "openPresentationHandler", { enumerable: true, get: function () { return presentation_3.openPresentationHandler; } });
Object.defineProperty(exports, "closePresentationDefinition", { enumerable: true, get: function () { return presentation_3.closePresentationDefinition; } });
Object.defineProperty(exports, "closePresentationHandler", { enumerable: true, get: function () { return presentation_3.closePresentationHandler; } });
Object.defineProperty(exports, "getOpenPresentationsDefinition", { enumerable: true, get: function () { return presentation_3.getOpenPresentationsDefinition; } });
Object.defineProperty(exports, "getOpenPresentationsHandler", { enumerable: true, get: function () { return presentation_3.getOpenPresentationsHandler; } });
Object.defineProperty(exports, "switchPresentationDefinition", { enumerable: true, get: function () { return presentation_3.switchPresentationDefinition; } });
Object.defineProperty(exports, "switchPresentationHandler", { enumerable: true, get: function () { return presentation_3.switchPresentationHandler; } });
Object.defineProperty(exports, "copySlideDefinition", { enumerable: true, get: function () { return presentation_3.copySlideDefinition; } });
Object.defineProperty(exports, "copySlideHandler", { enumerable: true, get: function () { return presentation_3.copySlideHandler; } });
Object.defineProperty(exports, "insertSlideImageDefinition", { enumerable: true, get: function () { return presentation_3.insertSlideImageDefinition; } });
Object.defineProperty(exports, "insertSlideImageHandler", { enumerable: true, get: function () { return presentation_3.insertSlideImageHandler; } });
Object.defineProperty(exports, "insertSlidesFromFileDefinition", { enumerable: true, get: function () { return presentation_3.insertSlidesFromFileDefinition; } });
Object.defineProperty(exports, "insertSlidesFromFileHandler", { enumerable: true, get: function () { return presentation_3.insertSlidesFromFileHandler; } });
Object.defineProperty(exports, "setActiveTargetDefinition", { enumerable: true, get: function () { return presentation_3.setActiveTargetDefinition; } });
Object.defineProperty(exports, "setActiveTargetHandler", { enumerable: true, get: function () { return presentation_3.setActiveTargetHandler; } });
var textbox_3 = require("./textbox");
Object.defineProperty(exports, "deleteTextboxDefinition", { enumerable: true, get: function () { return textbox_3.deleteTextboxDefinition; } });
Object.defineProperty(exports, "deleteTextboxHandler", { enumerable: true, get: function () { return textbox_3.deleteTextboxHandler; } });
Object.defineProperty(exports, "getTextboxesDefinition", { enumerable: true, get: function () { return textbox_3.getTextboxesDefinition; } });
Object.defineProperty(exports, "getTextboxesHandler", { enumerable: true, get: function () { return textbox_3.getTextboxesHandler; } });
Object.defineProperty(exports, "setTextboxTextDefinition", { enumerable: true, get: function () { return textbox_3.setTextboxTextDefinition; } });
Object.defineProperty(exports, "setTextboxTextHandler", { enumerable: true, get: function () { return textbox_3.setTextboxTextHandler; } });
Object.defineProperty(exports, "setTextboxStyleDefinition", { enumerable: true, get: function () { return textbox_3.setTextboxStyleDefinition; } });
Object.defineProperty(exports, "setTextboxStyleHandler", { enumerable: true, get: function () { return textbox_3.setTextboxStyleHandler; } });
Object.defineProperty(exports, "getSlideTitleDefinition", { enumerable: true, get: function () { return textbox_3.getSlideTitleDefinition; } });
Object.defineProperty(exports, "getSlideTitleHandler", { enumerable: true, get: function () { return textbox_3.getSlideTitleHandler; } });
Object.defineProperty(exports, "setSlideSubtitleDefinition", { enumerable: true, get: function () { return textbox_3.setSlideSubtitleDefinition; } });
Object.defineProperty(exports, "setSlideSubtitleHandler", { enumerable: true, get: function () { return textbox_3.setSlideSubtitleHandler; } });
Object.defineProperty(exports, "setSlideContentDefinition", { enumerable: true, get: function () { return textbox_3.setSlideContentDefinition; } });
Object.defineProperty(exports, "setSlideContentHandler", { enumerable: true, get: function () { return textbox_3.setSlideContentHandler; } });
var background_3 = require("./background");
Object.defineProperty(exports, "setSlideBackgroundDefinition", { enumerable: true, get: function () { return background_3.setSlideBackgroundDefinition; } });
Object.defineProperty(exports, "setSlideBackgroundHandler", { enumerable: true, get: function () { return background_3.setSlideBackgroundHandler; } });
Object.defineProperty(exports, "setBackgroundColorDefinition", { enumerable: true, get: function () { return background_3.setBackgroundColorDefinition; } });
Object.defineProperty(exports, "setBackgroundColorHandler", { enumerable: true, get: function () { return background_3.setBackgroundColorHandler; } });
Object.defineProperty(exports, "setBackgroundImageDefinition", { enumerable: true, get: function () { return background_3.setBackgroundImageDefinition; } });
Object.defineProperty(exports, "setBackgroundImageHandler", { enumerable: true, get: function () { return background_3.setBackgroundImageHandler; } });
Object.defineProperty(exports, "setSlideNumberDefinition", { enumerable: true, get: function () { return background_3.setSlideNumberDefinition; } });
Object.defineProperty(exports, "setSlideNumberHandler", { enumerable: true, get: function () { return background_3.setSlideNumberHandler; } });
Object.defineProperty(exports, "setPptFooterDefinition", { enumerable: true, get: function () { return background_3.setPptFooterDefinition; } });
Object.defineProperty(exports, "setPptFooterHandler", { enumerable: true, get: function () { return background_3.setPptFooterHandler; } });
Object.defineProperty(exports, "setPptDateTimeDefinition", { enumerable: true, get: function () { return background_3.setPptDateTimeDefinition; } });
Object.defineProperty(exports, "setPptDateTimeHandler", { enumerable: true, get: function () { return background_3.setPptDateTimeHandler; } });
Object.defineProperty(exports, "duplicateShapeDefinition", { enumerable: true, get: function () { return background_3.duplicateShapeDefinition; } });
Object.defineProperty(exports, "duplicateShapeHandler", { enumerable: true, get: function () { return background_3.duplicateShapeHandler; } });
Object.defineProperty(exports, "setShapeZOrderDefinition", { enumerable: true, get: function () { return background_3.setShapeZOrderDefinition; } });
Object.defineProperty(exports, "setShapeZOrderHandler", { enumerable: true, get: function () { return background_3.setShapeZOrderHandler; } });
var shape_basic_3 = require("./shape-basic");
Object.defineProperty(exports, "deleteShapeDefinition", { enumerable: true, get: function () { return shape_basic_3.deleteShapeDefinition; } });
Object.defineProperty(exports, "deleteShapeHandler", { enumerable: true, get: function () { return shape_basic_3.deleteShapeHandler; } });
Object.defineProperty(exports, "getShapesDefinition", { enumerable: true, get: function () { return shape_basic_3.getShapesDefinition; } });
Object.defineProperty(exports, "getShapesHandler", { enumerable: true, get: function () { return shape_basic_3.getShapesHandler; } });
Object.defineProperty(exports, "setShapePositionDefinition", { enumerable: true, get: function () { return shape_basic_3.setShapePositionDefinition; } });
Object.defineProperty(exports, "setShapePositionHandler", { enumerable: true, get: function () { return shape_basic_3.setShapePositionHandler; } });
Object.defineProperty(exports, "setShapeShadowDefinition", { enumerable: true, get: function () { return shape_basic_3.setShapeShadowDefinition; } });
Object.defineProperty(exports, "setShapeShadowHandler", { enumerable: true, get: function () { return shape_basic_3.setShapeShadowHandler; } });
Object.defineProperty(exports, "setShapeGradientDefinition", { enumerable: true, get: function () { return shape_basic_3.setShapeGradientDefinition; } });
Object.defineProperty(exports, "setShapeGradientHandler", { enumerable: true, get: function () { return shape_basic_3.setShapeGradientHandler; } });
Object.defineProperty(exports, "setShapeBorderDefinition", { enumerable: true, get: function () { return shape_basic_3.setShapeBorderDefinition; } });
Object.defineProperty(exports, "setShapeBorderHandler", { enumerable: true, get: function () { return shape_basic_3.setShapeBorderHandler; } });
Object.defineProperty(exports, "setShapeTransparencyDefinition", { enumerable: true, get: function () { return shape_basic_3.setShapeTransparencyDefinition; } });
Object.defineProperty(exports, "setShapeTransparencyHandler", { enumerable: true, get: function () { return shape_basic_3.setShapeTransparencyHandler; } });
Object.defineProperty(exports, "alignShapesDefinition", { enumerable: true, get: function () { return shape_basic_3.alignShapesDefinition; } });
Object.defineProperty(exports, "alignShapesHandler", { enumerable: true, get: function () { return shape_basic_3.alignShapesHandler; } });
Object.defineProperty(exports, "distributeShapesDefinition", { enumerable: true, get: function () { return shape_basic_3.distributeShapesDefinition; } });
Object.defineProperty(exports, "distributeShapesHandler", { enumerable: true, get: function () { return shape_basic_3.distributeShapesHandler; } });
Object.defineProperty(exports, "groupShapesDefinition", { enumerable: true, get: function () { return shape_basic_3.groupShapesDefinition; } });
Object.defineProperty(exports, "groupShapesHandler", { enumerable: true, get: function () { return shape_basic_3.groupShapesHandler; } });
var image_3 = require("./image");
Object.defineProperty(exports, "insertPptImageDefinition", { enumerable: true, get: function () { return image_3.insertPptImageDefinition; } });
Object.defineProperty(exports, "insertPptImageHandler", { enumerable: true, get: function () { return image_3.insertPptImageHandler; } });
Object.defineProperty(exports, "deletePptImageDefinition", { enumerable: true, get: function () { return image_3.deletePptImageDefinition; } });
Object.defineProperty(exports, "deletePptImageHandler", { enumerable: true, get: function () { return image_3.deletePptImageHandler; } });
Object.defineProperty(exports, "setImageStyleDefinition", { enumerable: true, get: function () { return image_3.setImageStyleDefinition; } });
Object.defineProperty(exports, "setImageStyleHandler", { enumerable: true, get: function () { return image_3.setImageStyleHandler; } });
Object.defineProperty(exports, "exportSlideAsImageDefinition", { enumerable: true, get: function () { return image_3.exportSlideAsImageDefinition; } });
Object.defineProperty(exports, "exportSlideAsImageHandler", { enumerable: true, get: function () { return image_3.exportSlideAsImageHandler; } });
Object.defineProperty(exports, "replacePptImageDefinition", { enumerable: true, get: function () { return image_3.replacePptImageDefinition; } });
Object.defineProperty(exports, "replacePptImageHandler", { enumerable: true, get: function () { return image_3.replacePptImageHandler; } });
var animation_3 = require("./animation");
Object.defineProperty(exports, "addAnimationDefinition", { enumerable: true, get: function () { return animation_3.addAnimationDefinition; } });
Object.defineProperty(exports, "addAnimationHandler", { enumerable: true, get: function () { return animation_3.addAnimationHandler; } });
Object.defineProperty(exports, "removeAnimationDefinition", { enumerable: true, get: function () { return animation_3.removeAnimationDefinition; } });
Object.defineProperty(exports, "removeAnimationHandler", { enumerable: true, get: function () { return animation_3.removeAnimationHandler; } });
Object.defineProperty(exports, "getAnimationsDefinition", { enumerable: true, get: function () { return animation_3.getAnimationsDefinition; } });
Object.defineProperty(exports, "getAnimationsHandler", { enumerable: true, get: function () { return animation_3.getAnimationsHandler; } });
Object.defineProperty(exports, "setAnimationOrderDefinition", { enumerable: true, get: function () { return animation_3.setAnimationOrderDefinition; } });
Object.defineProperty(exports, "setAnimationOrderHandler", { enumerable: true, get: function () { return animation_3.setAnimationOrderHandler; } });
Object.defineProperty(exports, "addAnimationPresetDefinition", { enumerable: true, get: function () { return animation_3.addAnimationPresetDefinition; } });
Object.defineProperty(exports, "addAnimationPresetHandler", { enumerable: true, get: function () { return animation_3.addAnimationPresetHandler; } });
Object.defineProperty(exports, "addEmphasisAnimationDefinition", { enumerable: true, get: function () { return animation_3.addEmphasisAnimationDefinition; } });
Object.defineProperty(exports, "addEmphasisAnimationHandler", { enumerable: true, get: function () { return animation_3.addEmphasisAnimationHandler; } });
Object.defineProperty(exports, "setSlideTransitionDefinition", { enumerable: true, get: function () { return animation_3.setSlideTransitionDefinition; } });
Object.defineProperty(exports, "setSlideTransitionHandler", { enumerable: true, get: function () { return animation_3.setSlideTransitionHandler; } });
Object.defineProperty(exports, "removeSlideTransitionDefinition", { enumerable: true, get: function () { return animation_3.removeSlideTransitionDefinition; } });
Object.defineProperty(exports, "removeSlideTransitionHandler", { enumerable: true, get: function () { return animation_3.removeSlideTransitionHandler; } });
Object.defineProperty(exports, "applyTransitionToAllDefinition", { enumerable: true, get: function () { return animation_3.applyTransitionToAllDefinition; } });
Object.defineProperty(exports, "applyTransitionToAllHandler", { enumerable: true, get: function () { return animation_3.applyTransitionToAllHandler; } });
var chart_flow_3 = require("./chart-flow");
Object.defineProperty(exports, "insertPptChartDefinition", { enumerable: true, get: function () { return chart_flow_3.insertPptChartDefinition; } });
Object.defineProperty(exports, "insertPptChartHandler", { enumerable: true, get: function () { return chart_flow_3.insertPptChartHandler; } });
Object.defineProperty(exports, "setPptChartDataDefinition", { enumerable: true, get: function () { return chart_flow_3.setPptChartDataDefinition; } });
Object.defineProperty(exports, "setPptChartDataHandler", { enumerable: true, get: function () { return chart_flow_3.setPptChartDataHandler; } });
Object.defineProperty(exports, "setPptChartStyleDefinition", { enumerable: true, get: function () { return chart_flow_3.setPptChartStyleDefinition; } });
Object.defineProperty(exports, "setPptChartStyleHandler", { enumerable: true, get: function () { return chart_flow_3.setPptChartStyleHandler; } });
var misc_3 = require("./misc");
Object.defineProperty(exports, "getSlideMasterDefinition", { enumerable: true, get: function () { return misc_3.getSlideMasterDefinition; } });
Object.defineProperty(exports, "getSlideMasterHandler", { enumerable: true, get: function () { return misc_3.getSlideMasterHandler; } });
Object.defineProperty(exports, "setMasterBackgroundDefinition", { enumerable: true, get: function () { return misc_3.setMasterBackgroundDefinition; } });
Object.defineProperty(exports, "setMasterBackgroundHandler", { enumerable: true, get: function () { return misc_3.setMasterBackgroundHandler; } });
Object.defineProperty(exports, "addMasterElementDefinition", { enumerable: true, get: function () { return misc_3.addMasterElementDefinition; } });
Object.defineProperty(exports, "addMasterElementHandler", { enumerable: true, get: function () { return misc_3.addMasterElementHandler; } });
Object.defineProperty(exports, "set3DRotationDefinition", { enumerable: true, get: function () { return misc_3.set3DRotationDefinition; } });
Object.defineProperty(exports, "set3DRotationHandler", { enumerable: true, get: function () { return misc_3.set3DRotationHandler; } });
Object.defineProperty(exports, "set3DDepthDefinition", { enumerable: true, get: function () { return misc_3.set3DDepthDefinition; } });
Object.defineProperty(exports, "set3DDepthHandler", { enumerable: true, get: function () { return misc_3.set3DDepthHandler; } });
Object.defineProperty(exports, "set3DMaterialDefinition", { enumerable: true, get: function () { return misc_3.set3DMaterialDefinition; } });
Object.defineProperty(exports, "set3DMaterialHandler", { enumerable: true, get: function () { return misc_3.set3DMaterialHandler; } });
Object.defineProperty(exports, "addPptHyperlinkDefinition", { enumerable: true, get: function () { return misc_3.addPptHyperlinkDefinition; } });
Object.defineProperty(exports, "addPptHyperlinkHandler", { enumerable: true, get: function () { return misc_3.addPptHyperlinkHandler; } });
Object.defineProperty(exports, "removePptHyperlinkDefinition", { enumerable: true, get: function () { return misc_3.removePptHyperlinkDefinition; } });
Object.defineProperty(exports, "removePptHyperlinkHandler", { enumerable: true, get: function () { return misc_3.removePptHyperlinkHandler; } });
Object.defineProperty(exports, "findPptTextDefinition", { enumerable: true, get: function () { return misc_3.findPptTextDefinition; } });
Object.defineProperty(exports, "findPptTextHandler", { enumerable: true, get: function () { return misc_3.findPptTextHandler; } });
Object.defineProperty(exports, "replacePptTextDefinition", { enumerable: true, get: function () { return misc_3.replacePptTextDefinition; } });
Object.defineProperty(exports, "replacePptTextHandler", { enumerable: true, get: function () { return misc_3.replacePptTextHandler; } });
Object.defineProperty(exports, "startSlideShowDefinition", { enumerable: true, get: function () { return misc_3.startSlideShowDefinition; } });
Object.defineProperty(exports, "startSlideShowHandler", { enumerable: true, get: function () { return misc_3.startSlideShowHandler; } });
var table_3 = require("./table");
Object.defineProperty(exports, "insertPptTableDefinition", { enumerable: true, get: function () { return table_3.insertPptTableDefinition; } });
Object.defineProperty(exports, "insertPptTableHandler", { enumerable: true, get: function () { return table_3.insertPptTableHandler; } });
Object.defineProperty(exports, "setPptTableCellDefinition", { enumerable: true, get: function () { return table_3.setPptTableCellDefinition; } });
Object.defineProperty(exports, "setPptTableCellHandler", { enumerable: true, get: function () { return table_3.setPptTableCellHandler; } });
Object.defineProperty(exports, "getPptTableCellDefinition", { enumerable: true, get: function () { return table_3.getPptTableCellDefinition; } });
Object.defineProperty(exports, "getPptTableCellHandler", { enumerable: true, get: function () { return table_3.getPptTableCellHandler; } });
Object.defineProperty(exports, "setPptTableStyleDefinition", { enumerable: true, get: function () { return table_3.setPptTableStyleDefinition; } });
Object.defineProperty(exports, "setPptTableStyleHandler", { enumerable: true, get: function () { return table_3.setPptTableStyleHandler; } });
Object.defineProperty(exports, "setPptTableCellStyleDefinition", { enumerable: true, get: function () { return table_3.setPptTableCellStyleDefinition; } });
Object.defineProperty(exports, "setPptTableCellStyleHandler", { enumerable: true, get: function () { return table_3.setPptTableCellStyleHandler; } });
Object.defineProperty(exports, "setPptTableRowStyleDefinition", { enumerable: true, get: function () { return table_3.setPptTableRowStyleDefinition; } });
Object.defineProperty(exports, "setPptTableRowStyleHandler", { enumerable: true, get: function () { return table_3.setPptTableRowStyleHandler; } });
var beautify_advanced_3 = require("./beautify-advanced");
Object.defineProperty(exports, "setBackgroundGradientDefinition", { enumerable: true, get: function () { return beautify_advanced_3.setBackgroundGradientDefinition; } });
Object.defineProperty(exports, "setBackgroundGradientHandler", { enumerable: true, get: function () { return beautify_advanced_3.setBackgroundGradientHandler; } });
exports.default = exports.pptTools;
//# sourceMappingURL=index.js.map
/**
 * Input: PPT 图片操作参数（插入、删除、样式设置、幻灯片导出为图片）
 * Output: 图片操作结果
 * Pos: PPT 图片工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT图片Tools - 图片管理模块
 * 处理图片的插入、删除、样式设置以及幻灯片导出为位图操作
 *
 * 包含：
 * - wps_ppt_insert_ppt_image: 插入图片到幻灯片
 * - wps_ppt_delete_ppt_image: 删除幻灯片中的图片
 * - wps_ppt_set_image_style: 设置图片样式
 * - wps_ppt_export_slide_as_image: 将指定幻灯片导出为位图图片（PNG/JPG/GIF/BMP）
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
export declare const insertPptImageDefinition: ToolDefinition;
export declare const insertPptImageHandler: ToolHandler;
export declare const deletePptImageDefinition: ToolDefinition;
export declare const deletePptImageHandler: ToolHandler;
export declare const setImageStyleDefinition: ToolDefinition;
export declare const setImageStyleHandler: ToolHandler;
export declare const exportSlideAsImageDefinition: ToolDefinition;
export declare const exportSlideAsImageHandler: ToolHandler;
export declare const replacePptImageDefinition: ToolDefinition;
export declare const replacePptImageHandler: ToolHandler;
/**
 * 导出所有图片相关的Tools
 */
export declare const imageTools: RegisteredTool[];
export default imageTools;
//# sourceMappingURL=image.d.ts.map
/**
 * Input: PPT 幻灯片操作参数
 * Output: 幻灯片操作结果
 * Pos: PPT 幻灯片工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT幻灯片Tools - 幻灯片管理模块
 * 处理幻灯片的添加、美化、字体统一等操作
 *
 * 包含：
 * - wps_ppt_add_slide: 添加新幻灯片
 * - wps_ppt_beautify: 美化幻灯片（核心功能）
 * - wps_ppt_unify_font: 统一字体
 * - wps_ppt_set_font_color: 设置文字颜色
 * - wps_ppt_align_objects: 对齐幻灯片中的对象
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 添加新幻灯片
 * 可以指定布局、位置、标题和内容
 */
export declare const addSlideDefinition: ToolDefinition;
export declare const addSlideHandler: ToolHandler;
/**
 * 美化幻灯片
 * 一键优化排版、配色、字体和间距
 */
export declare const beautifyDefinition: ToolDefinition;
export declare const beautifyHandler: ToolHandler;
/**
 * 统一字体
 * 整个演示文稿的字体统一是PPT美观的基础
 */
export declare const unifyFontDefinition: ToolDefinition;
export declare const unifyFontHandler: ToolHandler;
/**
 * 设置文字颜色
 * 修改幻灯片中指定形状的文字颜色
 */
export declare const setFontColorDefinition: ToolDefinition;
export declare const setFontColorHandler: ToolHandler;
/**
 * 对齐幻灯片中的对象
 * 支持多种对齐方式
 */
export declare const alignObjectsDefinition: ToolDefinition;
export declare const alignObjectsHandler: ToolHandler;
/**
 * 导出所有幻灯片相关的Tools
 */
export declare const slideTools: RegisteredTool[];
export default slideTools;
//# sourceMappingURL=slide.d.ts.map
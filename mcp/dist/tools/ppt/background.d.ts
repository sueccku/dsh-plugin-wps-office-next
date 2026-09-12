/**
 * Input: PPT 背景、页脚、形状操作参数
 * Output: 背景/页脚/形状操作结果
 * Pos: PPT 背景与页面信息工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT背景Tools - 背景、页面信息、高级形状模块
 * 处理幻灯片背景设置、页码/页脚/日期时间、形状复制与层级调整
 *
 * 包含：
 * - wps_ppt_set_slide_background: 设置幻灯片背景（通用）
 * - wps_ppt_set_background_color: 设置幻灯片背景颜色
 * - wps_ppt_set_background_image: 设置幻灯片背景图片
 * - wps_ppt_set_slide_number: 设置幻灯片页码显示
 * - wps_ppt_set_ppt_footer: 设置PPT页脚
 * - wps_ppt_set_ppt_date_time: 设置PPT日期时间
 * - wps_ppt_duplicate_shape: 复制形状
 * - wps_ppt_set_shape_z_order: 设置形状层级顺序
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 设置幻灯片背景（通用）
 * 支持纯色、渐变、图片等多种背景类型
 */
export declare const setSlideBackgroundDefinition: ToolDefinition;
export declare const setSlideBackgroundHandler: ToolHandler;
/**
 * 设置幻灯片背景颜色
 * 快速设置纯色背景的便捷方法
 */
export declare const setBackgroundColorDefinition: ToolDefinition;
export declare const setBackgroundColorHandler: ToolHandler;
/**
 * 设置幻灯片背景图片
 * 使用图片文件作为幻灯片背景
 */
export declare const setBackgroundImageDefinition: ToolDefinition;
export declare const setBackgroundImageHandler: ToolHandler;
/**
 * 设置幻灯片页码显示
 * 控制页码的显示/隐藏和起始编号
 */
export declare const setSlideNumberDefinition: ToolDefinition;
export declare const setSlideNumberHandler: ToolHandler;
/**
 * 设置PPT页脚
 * 控制页脚文本的显示/隐藏
 */
export declare const setPptFooterDefinition: ToolDefinition;
export declare const setPptFooterHandler: ToolHandler;
/**
 * 设置PPT日期时间
 * 控制日期时间占位符的显示和格式
 */
export declare const setPptDateTimeDefinition: ToolDefinition;
export declare const setPptDateTimeHandler: ToolHandler;
/**
 * 复制形状
 * 在同一幻灯片内复制指定形状
 */
export declare const duplicateShapeDefinition: ToolDefinition;
export declare const duplicateShapeHandler: ToolHandler;
/**
 * 设置形状层级顺序
 * 调整形状在幻灯片中的前后层级
 */
export declare const setShapeZOrderDefinition: ToolDefinition;
export declare const setShapeZOrderHandler: ToolHandler;
/**
 * 导出所有背景、页面信息、高级形状相关的Tools
 */
export declare const backgroundTools: RegisteredTool[];
export default backgroundTools;
//# sourceMappingURL=background.d.ts.map
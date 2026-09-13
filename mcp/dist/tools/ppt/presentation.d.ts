/**
 * Input: 演示文稿管理工具参数
 * Output: 演示文稿操作结果
 * Pos: PPT 演示文稿管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 包含：
 * - wps_ppt_create_presentation: 新建空白演示文稿
 * - wps_ppt_open_presentation: 打开指定路径的演示文稿
 * - wps_ppt_close_presentation: 关闭演示文稿
 * - wps_ppt_get_open_presentations: 获取所有已打开的演示文稿列表
 * - wps_ppt_switch_presentation: 切换到指定演示文稿
 * - wps_ppt_copy_slide: 复制幻灯片
 * - wps_ppt_insert_slide_image: 在幻灯片中插入图片
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 新建空白演示文稿
 */
export declare const createPresentationDefinition: ToolDefinition;
export declare const createPresentationHandler: ToolHandler;
/**
 * 打开指定路径的演示文稿
 */
export declare const openPresentationDefinition: ToolDefinition;
export declare const openPresentationHandler: ToolHandler;
/**
 * 关闭演示文稿
 */
export declare const closePresentationDefinition: ToolDefinition;
export declare const closePresentationHandler: ToolHandler;
/**
 * 获取所有已打开的演示文稿列表
 */
export declare const getOpenPresentationsDefinition: ToolDefinition;
export declare const getOpenPresentationsHandler: ToolHandler;
/**
 * 切换到指定演示文稿
 */
export declare const switchPresentationDefinition: ToolDefinition;
export declare const switchPresentationHandler: ToolHandler;
export declare const copySlideDefinition: ToolDefinition;
export declare const copySlideHandler: ToolHandler;
/**
 * 从其它演示文稿导入整页幻灯片（跨PPT整合，保留来源格式）
 */
export declare const insertSlidesFromFileDefinition: ToolDefinition;
export declare const insertSlidesFromFileHandler: ToolHandler;
/**
 * 锁定后续所有 PPT 操作的目标演示文稿（避免多文稿打开时活动文稿漂移改错文件）
 */
export declare const setActiveTargetDefinition: ToolDefinition;
export declare const setActiveTargetHandler: ToolHandler;
/**
 * 应用演示文稿主题（模板文件）
 */
export declare const setSlideThemeDefinition: ToolDefinition;
export declare const setSlideThemeHandler: ToolHandler;
/**
 * 导出所有演示文稿管理相关的Tools
 */
export declare const presentationTools: RegisteredTool[];
export default presentationTools;
//# sourceMappingURL=presentation.d.ts.map
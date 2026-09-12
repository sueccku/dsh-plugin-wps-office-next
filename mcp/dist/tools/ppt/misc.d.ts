/**
 * Input: PPT 杂项操作参数（母版、3D、超链接、搜索、放映）
 * Output: 杂项操作结果
 * Pos: PPT 杂项工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT杂项Tools - 母版/3D/超链接/搜索/放映模块
 *
 * 包含：
 * - wps_ppt_get_slide_master: 获取母版信息
 * - wps_ppt_set_master_background: 设置母版背景
 * - wps_ppt_add_master_element: 添加母版元素
 * - wps_ppt_set_3d_rotation: 设置3D旋转
 * - wps_ppt_set_3d_depth: 设置3D深度
 * - wps_ppt_set_3d_material: 设置3D材质
 * - wps_ppt_create_3d_text: 创建3D文字
 * - wps_ppt_add_ppt_hyperlink: 添加超链接
 * - wps_ppt_remove_ppt_hyperlink: 移除超链接
 * - wps_ppt_find_ppt_text: 搜索文本
 * - wps_ppt_replace_ppt_text: 替换文本
 * - wps_ppt_start_slide_show: 开始放映
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 获取母版信息
 * 获取当前演示文稿的母版布局信息
 */
export declare const getSlideMasterDefinition: ToolDefinition;
export declare const getSlideMasterHandler: ToolHandler;
/**
 * 设置母版背景
 * 修改母版的背景样式
 */
export declare const setMasterBackgroundDefinition: ToolDefinition;
export declare const setMasterBackgroundHandler: ToolHandler;
/**
 * 添加母版元素
 * 向母版中添加新的元素（文本框、形状、图片等）
 */
export declare const addMasterElementDefinition: ToolDefinition;
export declare const addMasterElementHandler: ToolHandler;
/**
 * 设置3D旋转
 * 为形状设置3D旋转效果
 */
export declare const set3DRotationDefinition: ToolDefinition;
export declare const set3DRotationHandler: ToolHandler;
/**
 * 设置3D深度
 * 为形状设置3D挤出深度
 */
export declare const set3DDepthDefinition: ToolDefinition;
export declare const set3DDepthHandler: ToolHandler;
/**
 * 设置3D材质
 * 为形状设置3D材质效果
 */
export declare const set3DMaterialDefinition: ToolDefinition;
export declare const set3DMaterialHandler: ToolHandler;
/**
 * 创建3D文字
 * 在幻灯片中创建带有3D效果的文字
 */
export declare const create3DTextDefinition: ToolDefinition;
export declare const create3DTextHandler: ToolHandler;
/**
 * 添加超链接
 * 为形状添加超链接
 */
export declare const addPptHyperlinkDefinition: ToolDefinition;
export declare const addPptHyperlinkHandler: ToolHandler;
/**
 * 移除超链接
 * 删除形状上的超链接
 */
export declare const removePptHyperlinkDefinition: ToolDefinition;
export declare const removePptHyperlinkHandler: ToolHandler;
/**
 * 搜索文本
 * 在演示文稿中搜索指定文本
 */
export declare const findPptTextDefinition: ToolDefinition;
export declare const findPptTextHandler: ToolHandler;
/**
 * 替换文本
 * 在演示文稿中查找并替换文本
 */
export declare const replacePptTextDefinition: ToolDefinition;
export declare const replacePptTextHandler: ToolHandler;
/**
 * 开始放映
 * 从指定页开始幻灯片放映
 */
export declare const startSlideShowDefinition: ToolDefinition;
export declare const startSlideShowHandler: ToolHandler;
/**
 * 导出所有杂项Tools
 */
export declare const miscTools: RegisteredTool[];
export default miscTools;
//# sourceMappingURL=misc.d.ts.map
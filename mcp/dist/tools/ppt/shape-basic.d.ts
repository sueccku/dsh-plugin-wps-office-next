/**
 * Input: PPT 形状操作参数（删除、获取、位置、阴影、渐变、边框、透明度、对齐、分布、组合）
 * Output: 形状操作结果
 * Pos: PPT 形状基础工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT形状基础Tools - 形状管理模块
 * 处理形状的删除、获取、定位、样式设置、对齐分布和组合操作
 *
 * 包含：
 * - wps_ppt_delete_shape: 删除形状
 * - wps_ppt_get_shapes: 获取幻灯片中的形状列表
 * - wps_ppt_set_shape_position: 设置形状位置和大小
 * - wps_ppt_set_shape_shadow: 设置形状阴影
 * - wps_ppt_set_shape_gradient: 设置形状渐变填充
 * - wps_ppt_set_shape_border: 设置形状边框
 * - wps_ppt_set_shape_transparency: 设置形状透明度
 * - wps_ppt_align_shapes: 对齐多个形状
 * - wps_ppt_distribute_shapes: 等距分布多个形状
 * - wps_ppt_group_shapes: 组合多个形状
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
export declare const deleteShapeDefinition: ToolDefinition;
export declare const deleteShapeHandler: ToolHandler;
export declare const getShapesDefinition: ToolDefinition;
export declare const getShapesHandler: ToolHandler;
export declare const setShapePositionDefinition: ToolDefinition;
export declare const setShapePositionHandler: ToolHandler;
export declare const setShapeShadowDefinition: ToolDefinition;
export declare const setShapeShadowHandler: ToolHandler;
export declare const setShapeGradientDefinition: ToolDefinition;
export declare const setShapeGradientHandler: ToolHandler;
export declare const setShapeBorderDefinition: ToolDefinition;
export declare const setShapeBorderHandler: ToolHandler;
export declare const setShapeTransparencyDefinition: ToolDefinition;
export declare const setShapeTransparencyHandler: ToolHandler;
export declare const alignShapesDefinition: ToolDefinition;
export declare const alignShapesHandler: ToolHandler;
export declare const distributeShapesDefinition: ToolDefinition;
export declare const distributeShapesHandler: ToolHandler;
export declare const groupShapesDefinition: ToolDefinition;
export declare const groupShapesHandler: ToolHandler;
/**
 * 导出所有形状基础相关的Tools
 */
export declare const shapeBasicTools: RegisteredTool[];
export default shapeBasicTools;
//# sourceMappingURL=shape-basic.d.ts.map
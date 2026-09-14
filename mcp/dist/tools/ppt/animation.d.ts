/**
 * Input: PPT 动画与切换效果操作参数
 * Output: 动画/切换效果操作结果
 * Pos: PPT 动画与切换工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT动画Tools - 动画与幻灯片切换管理模块
 * 处理动画的添加、删除、排序、预设以及幻灯片切换效果
 *
 * 包含：
 * - wps_ppt_add_animation: 添加动画效果
 * - wps_ppt_remove_animation: 移除动画效果
 * - wps_ppt_get_animations: 获取动画列表
 * - wps_ppt_set_animation_order: 设置动画顺序
 * - wps_ppt_add_animation_preset: 添加预设入场动画
 * - wps_ppt_add_emphasis_animation: 添加强调动画
 * - wps_ppt_set_slide_transition: 设置幻灯片切换效果
 * - wps_ppt_remove_slide_transition: 移除幻灯片切换效果
 * - wps_ppt_apply_transition_to_all: 应用切换效果到所有幻灯片
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 移除动画效果
 * 从指定幻灯片中移除指定的动画
 */
export declare const removeAnimationDefinition: ToolDefinition;
export declare const removeAnimationHandler: ToolHandler;
/**
 * 获取动画列表
 * 查看指定幻灯片上所有动画效果
 */
export declare const getAnimationsDefinition: ToolDefinition;
export declare const getAnimationsHandler: ToolHandler;
/**
 * 设置动画顺序
 * 调整动画在时间线上的播放顺序
 */
export declare const setAnimationOrderDefinition: ToolDefinition;
export declare const setAnimationOrderHandler: ToolHandler;
/**
 * 设置幻灯片切换效果
 * 为幻灯片设置页面切换动画
 */
export declare const setSlideTransitionDefinition: ToolDefinition;
export declare const setSlideTransitionHandler: ToolHandler;
/**
 * 移除幻灯片切换效果
 * 清除指定幻灯片的切换动画
 */
export declare const removeSlideTransitionDefinition: ToolDefinition;
export declare const removeSlideTransitionHandler: ToolHandler;
/**
 * 应用切换效果到所有幻灯片
 * 一次性为所有幻灯片设置统一的切换效果
 */
export declare const applyTransitionToAllDefinition: ToolDefinition;
export declare const applyTransitionToAllHandler: ToolHandler;
/**
 * 导出所有动画与切换相关的Tools
 */
/** 动画：入场/退场、强调、整页预设（P4 由三个工具合并而来） */
export declare const setAnimationDefinition: ToolDefinition;
export declare const setAnimationHandler: ToolHandler;
export declare const animationTools: RegisteredTool[];
export default animationTools;
//# sourceMappingURL=animation.d.ts.map
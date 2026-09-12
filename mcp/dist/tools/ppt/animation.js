"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.animationTools = exports.applyTransitionToAllHandler = exports.applyTransitionToAllDefinition = exports.removeSlideTransitionHandler = exports.removeSlideTransitionDefinition = exports.setSlideTransitionHandler = exports.setSlideTransitionDefinition = exports.addEmphasisAnimationHandler = exports.addEmphasisAnimationDefinition = exports.addAnimationPresetHandler = exports.addAnimationPresetDefinition = exports.setAnimationOrderHandler = exports.setAnimationOrderDefinition = exports.getAnimationsHandler = exports.getAnimationsDefinition = exports.removeAnimationHandler = exports.removeAnimationDefinition = exports.addAnimationHandler = exports.addAnimationDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
// ==================== 1. 添加动画效果 ====================
/**
 * 添加动画效果
 * 为指定幻灯片上的形状添加动画
 */
exports.addAnimationDefinition = {
    name: 'wps_ppt_add_animation',
    description: `为幻灯片中的形状添加动画效果。

使用场景：
- "给第1页第2个形状加个淡入动画"
- "为标题添加飞入效果"
- "添加动画让元素依次出现"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            shapeIndex: {
                type: 'number',
                description: '形状索引（从1开始）',
            },
            effect: {
                type: 'string',
                description: '动画效果名称，如 "fadeIn"、"flyIn"、"wipe" 等',
            },
            trigger: {
                type: 'string',
                description: '触发方式',
                enum: ['onClick', 'withPrevious', 'afterPrevious'],
            },
        },
        required: ['slideIndex', 'shapeIndex', 'effect'],
    },
};
const addAnimationHandler = async (args) => {
    const { slideIndex, shapeIndex, effect, trigger } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addAnimation', {
            slideIndex,
            shapeIndex,
            effect,
            trigger: trigger || 'onClick',
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const triggerName = {
                onClick: '单击时',
                withPrevious: '与上一动画同时',
                afterPrevious: '上一动画之后',
            };
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `动画添加成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n效果: ${effect}\n触发: ${triggerName[trigger || 'onClick'] || trigger}\n动画序号: ${response.data.animationIndex}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `添加动画失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `添加动画出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.addAnimationHandler = addAnimationHandler;
// ==================== 2. 移除动画效果 ====================
/**
 * 移除动画效果
 * 从指定幻灯片中移除指定的动画
 */
exports.removeAnimationDefinition = {
    name: 'wps_ppt_remove_animation',
    description: `移除幻灯片中指定的动画效果。

使用场景：
- "删除第1页的第2个动画"
- "移除这个动画效果"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            animationIndex: {
                type: 'number',
                description: '动画索引（从1开始）',
            },
        },
        required: ['slideIndex', 'animationIndex'],
    },
};
const removeAnimationHandler = async (args) => {
    const { slideIndex, animationIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('removeAnimation', { slideIndex, animationIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `动画移除成功！\n幻灯片: 第 ${slideIndex} 页\n已移除第 ${animationIndex} 个动画`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `移除动画失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `移除动画出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.removeAnimationHandler = removeAnimationHandler;
// ==================== 3. 获取动画列表 ====================
/**
 * 获取动画列表
 * 查看指定幻灯片上所有动画效果
 */
exports.getAnimationsDefinition = {
    name: 'wps_ppt_get_animations',
    description: `获取幻灯片上所有动画效果的列表。

使用场景：
- "查看第1页有哪些动画"
- "列出这页的所有动画"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
        },
        required: ['slideIndex'],
    },
};
const getAnimationsHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getAnimations', { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const animations = response.data.animations;
            if (!animations || animations.length === 0) {
                return {
                    id: (0, uuid_1.v4)(),
                    success: true,
                    content: [
                        {
                            type: 'text',
                            text: `第 ${slideIndex} 页没有动画效果。`,
                        },
                    ],
                };
            }
            const triggerName = {
                onClick: '单击时',
                withPrevious: '与上一动画同时',
                afterPrevious: '上一动画之后',
            };
            let output = `第 ${slideIndex} 页的动画效果（共 ${animations.length} 个）：\n`;
            animations.forEach((anim) => {
                output += `  ${anim.index}. [${anim.shapeName}] ${anim.effect} - ${triggerName[anim.trigger] || anim.trigger}\n`;
            });
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: output }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取动画列表失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取动画列表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getAnimationsHandler = getAnimationsHandler;
// ==================== 4. 设置动画顺序 ====================
/**
 * 设置动画顺序
 * 调整动画在时间线上的播放顺序
 */
exports.setAnimationOrderDefinition = {
    name: 'wps_ppt_set_animation_order',
    description: `调整动画在时间线上的播放顺序。

使用场景：
- "把第3个动画移到第1个播放"
- "调整动画顺序"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            animationIndex: {
                type: 'number',
                description: '当前动画索引（从1开始）',
            },
            newOrder: {
                type: 'number',
                description: '新的播放顺序位置（从1开始）',
            },
        },
        required: ['slideIndex', 'animationIndex', 'newOrder'],
    },
};
const setAnimationOrderHandler = async (args) => {
    const { slideIndex, animationIndex, newOrder } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setAnimationOrder', { slideIndex, animationIndex, newOrder }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `动画顺序调整成功！\n幻灯片: 第 ${slideIndex} 页\n动画 ${animationIndex} → 新位置 ${newOrder}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `调整动画顺序失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `调整动画顺序出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setAnimationOrderHandler = setAnimationOrderHandler;
// ==================== 5. 添加预设入场动画 ====================
/**
 * 添加预设入场动画
 * 提供常用的入场动画预设，简化动画添加流程
 */
exports.addAnimationPresetDefinition = {
    name: 'wps_ppt_add_animation_preset',
    description: `为形状添加预设入场动画效果。

支持的预设：
- fadeIn: 淡入
- flyIn: 飞入
- wipe: 擦除
- zoom: 缩放
- bounce: 弹跳
- spin: 旋转

使用场景：
- "给标题加个淡入效果"
- "让这个形状飞入"
- "添加弹跳入场动画"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            shapeIndex: {
                type: 'number',
                description: '形状索引（从1开始）',
            },
            preset: {
                type: 'string',
                description: '预设动画类型',
                enum: ['fadeIn', 'flyIn', 'wipe', 'zoom', 'bounce', 'spin'],
            },
        },
        required: ['slideIndex', 'shapeIndex', 'preset'],
    },
};
const addAnimationPresetHandler = async (args) => {
    const { slideIndex, shapeIndex, preset } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addAnimationPreset', { slideIndex, shapeIndex, preset }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const presetName = {
                fadeIn: '淡入',
                flyIn: '飞入',
                wipe: '擦除',
                zoom: '缩放',
                bounce: '弹跳',
                spin: '旋转',
            };
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `预设动画添加成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n预设: ${presetName[preset] || preset}\n动画序号: ${response.data.animationIndex}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `添加预设动画失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `添加预设动画出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.addAnimationPresetHandler = addAnimationPresetHandler;
// ==================== 6. 添加强调动画 ====================
/**
 * 添加强调动画
 * 为已有形状添加强调型动画效果，用于突出显示
 */
exports.addEmphasisAnimationDefinition = {
    name: 'wps_ppt_add_emphasis_animation',
    description: `为形状添加强调动画效果，用于在演示时突出显示元素。

支持的效果：
- pulse: 脉冲
- spin: 陀螺旋
- grow: 放大/缩小
- teeter: 跷跷板
- colorPulse: 颜色脉冲

使用场景：
- "让这个元素闪烁突出"
- "添加脉冲强调效果"
- "让图片旋转强调"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            shapeIndex: {
                type: 'number',
                description: '形状索引（从1开始）',
            },
            effect: {
                type: 'string',
                description: '强调动画效果',
                enum: ['pulse', 'spin', 'grow', 'teeter', 'colorPulse'],
            },
        },
        required: ['slideIndex', 'shapeIndex', 'effect'],
    },
};
const addEmphasisAnimationHandler = async (args) => {
    const { slideIndex, shapeIndex, effect } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addEmphasisAnimation', { slideIndex, shapeIndex, effect }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const effectName = {
                pulse: '脉冲',
                spin: '陀螺旋',
                grow: '放大/缩小',
                teeter: '跷跷板',
                colorPulse: '颜色脉冲',
            };
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `强调动画添加成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n效果: ${effectName[effect] || effect}\n动画序号: ${response.data.animationIndex}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `添加强调动画失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `添加强调动画出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.addEmphasisAnimationHandler = addEmphasisAnimationHandler;
// ==================== 7. 设置幻灯片切换效果 ====================
/**
 * 设置幻灯片切换效果
 * 为幻灯片设置页面切换动画
 */
exports.setSlideTransitionDefinition = {
    name: 'wps_ppt_set_slide_transition',
    description: `设置幻灯片的页面切换效果。

常用切换效果：
- fade: 淡出
- push: 推入
- wipe: 擦除
- split: 分割
- reveal: 显露
- cover: 覆盖
- curtains: 帷幕
- blinds: 百叶窗

使用场景：
- "给第1页加个淡出切换"
- "设置页面切换为推入效果"
- "添加切换动画"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            effect: {
                type: 'string',
                description: '切换效果名称',
            },
            duration: {
                type: 'number',
                description: '切换持续时间（秒），默认1秒',
            },
            sound: {
                type: 'string',
                description: '切换时播放的声音文件路径（可选）',
            },
        },
        required: ['slideIndex', 'effect'],
    },
};
const setSlideTransitionHandler = async (args) => {
    const { slideIndex, effect, duration, sound } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSlideTransition', {
            slideIndex,
            effect,
            duration: duration || 1,
            sound,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            let output = `切换效果设置成功！\n幻灯片: 第 ${slideIndex} 页\n效果: ${effect}\n持续时间: ${duration || 1} 秒`;
            if (sound) {
                output += `\n声音: ${sound}`;
            }
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: output }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置切换效果失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置切换效果出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setSlideTransitionHandler = setSlideTransitionHandler;
// ==================== 8. 移除幻灯片切换效果 ====================
/**
 * 移除幻灯片切换效果
 * 清除指定幻灯片的切换动画
 */
exports.removeSlideTransitionDefinition = {
    name: 'wps_ppt_remove_slide_transition',
    description: `移除幻灯片的切换效果。

使用场景：
- "取消第1页的切换效果"
- "移除页面切换动画"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
        },
        required: ['slideIndex'],
    },
};
const removeSlideTransitionHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('removeSlideTransition', { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `切换效果已移除！\n幻灯片: 第 ${slideIndex} 页`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `移除切换效果失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `移除切换效果出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.removeSlideTransitionHandler = removeSlideTransitionHandler;
// ==================== 9. 应用切换效果到所有幻灯片 ====================
/**
 * 应用切换效果到所有幻灯片
 * 一次性为所有幻灯片设置统一的切换效果
 */
exports.applyTransitionToAllDefinition = {
    name: 'wps_ppt_apply_transition_to_all',
    description: `为所有幻灯片应用统一的切换效果。

使用场景：
- "给所有页面加上淡出切换"
- "统一设置切换效果"
- "所有幻灯片用推入切换"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            effect: {
                type: 'string',
                description: '切换效果名称，如 "fade"、"push"、"wipe" 等',
            },
            duration: {
                type: 'number',
                description: '切换持续时间（秒），默认1秒',
            },
        },
        required: ['effect'],
    },
};
const applyTransitionToAllHandler = async (args) => {
    const { effect, duration } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('applyTransitionToAll', {
            effect,
            duration: duration || 1,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `切换效果已应用到所有幻灯片！\n效果: ${effect}\n持续时间: ${duration || 1} 秒\n应用页数: ${response.data.slideCount} 页`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `应用切换效果失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `应用切换效果出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.applyTransitionToAllHandler = applyTransitionToAllHandler;
/**
 * 导出所有动画与切换相关的Tools
 */
exports.animationTools = [
    { definition: exports.addAnimationDefinition, handler: exports.addAnimationHandler },
    { definition: exports.removeAnimationDefinition, handler: exports.removeAnimationHandler },
    { definition: exports.getAnimationsDefinition, handler: exports.getAnimationsHandler },
    { definition: exports.setAnimationOrderDefinition, handler: exports.setAnimationOrderHandler },
    { definition: exports.addAnimationPresetDefinition, handler: exports.addAnimationPresetHandler },
    { definition: exports.addEmphasisAnimationDefinition, handler: exports.addEmphasisAnimationHandler },
    { definition: exports.setSlideTransitionDefinition, handler: exports.setSlideTransitionHandler },
    { definition: exports.removeSlideTransitionDefinition, handler: exports.removeSlideTransitionHandler },
    { definition: exports.applyTransitionToAllDefinition, handler: exports.applyTransitionToAllHandler },
];
exports.default = exports.animationTools;
//# sourceMappingURL=animation.js.map
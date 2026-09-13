"use strict";
/**
 * Input: PPT 高级美化操作参数
 * Output: 高级美化操作结果
 * Pos: PPT 高级美化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT高级美化Tools - 高级美化模块
 * 处理配色方案、自动美化、KPI卡片、装饰元素等高级美化操作
 *
 * 包含：
 * - wps_ppt_set_background_gradient: 设置渐变背景
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.beautifyAdvancedTools = exports.setBackgroundGradientHandler = exports.setBackgroundGradientDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
exports.setBackgroundGradientDefinition = {
    name: 'wps_ppt_set_background_gradient',
    description: `为幻灯片设置渐变色背景。

渐变配置对象（gradient）属性：
- type: 渐变类型（linear/radial），默认linear
- angle: 渐变角度（0-360度），仅linear有效，默认180
- colors: 渐变颜色数组（如 ["#1a1a2e", "#16213e", "#0f3460"]）
- stops: 颜色停靠点数组（如 [0, 0.5, 1]），与colors对应

使用场景：
- "设置蓝色渐变背景"
- "给第2页加个从深到浅的渐变"
- "设置PPT背景为渐变色"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            gradient: {
                type: 'object',
                description: '渐变配置对象，包含type（linear/radial）、angle（角度）、colors（颜色数组）、stops（停靠点数组）',
            },
        },
        required: ['slideIndex', 'gradient'],
    },
};
const setBackgroundGradientHandler = async (args) => {
    const { slideIndex, gradient } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setBackgroundGradient', { slideIndex, gradient }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const gradientType = gradient.type === 'radial' ? '径向渐变' : '线性渐变';
            const colors = Array.isArray(gradient.colors) ? gradient.colors.join(' → ') : '';
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `渐变背景设置成功！\n幻灯片: 第 ${slideIndex} 页\n渐变类型: ${gradientType}${gradient.angle ? `\n渐变角度: ${gradient.angle}°` : ''}${colors ? `\n颜色: ${colors}` : ''}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置渐变背景失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置渐变背景出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setBackgroundGradientHandler = setBackgroundGradientHandler;
/**
 * 导出所有高级美化相关的Tools
 */
exports.beautifyAdvancedTools = [
    { definition: exports.setBackgroundGradientDefinition, handler: exports.setBackgroundGradientHandler },
];
exports.default = exports.beautifyAdvancedTools;
//# sourceMappingURL=beautify-advanced.js.map
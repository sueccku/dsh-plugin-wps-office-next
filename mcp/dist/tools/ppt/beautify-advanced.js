"use strict";
/**
 * Input: PPT 高级美化操作参数
 * Output: 高级美化操作结果
 * Pos: PPT 高级美化工具实现。一旦我被修改，请更新我的头部注释，并在 docs/FIXES.md 记账。
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
    description: `为幻灯片设置两色渐变背景（WPS 桥实现只支持两色线性渐变）。

颜色用 color1/color2（也可放在 gradient:{color1,color2} 里）。
注意：type/angle/colors/stops 等高级渐变参数当前 WPS 桥不支持，传了会被当成未知参数。

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
            color1: { type: 'string', description: '起始颜色，如 #1a1a2e' },
            color2: { type: 'string', description: '结束颜色，如 #0f3460' },
            gradient: {
                type: 'object',
                description: '兼容写法：{color1:"#xxx",color2:"#xxx"}；桥会把它展平成 color1/color2',
            },
        },
        required: ['slideIndex'],
    },
};
const setBackgroundGradientHandler = async (args) => {
    const { slideIndex, gradient, color1, color2 } = args;
    try {
        const params = { slideIndex };
        if (color1 !== undefined)
            params.color1 = color1;
        if (color2 !== undefined)
            params.color2 = color2;
        if (gradient !== undefined)
            params.gradient = gradient;
        const response = await wps_client_1.wpsClient.executeMethod('setBackgroundGradient', params, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const g = gradient;
            const c1 = color1 ?? g?.color1;
            const c2 = color2 ?? g?.color2;
            const colorText = c1 || c2 ? `\n颜色: ${c1 ?? '?'} → ${c2 ?? '?'}` : '';
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `渐变背景设置成功！\n幻灯片: 第 ${slideIndex} 页（两色线性渐变）${colorText}`,
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
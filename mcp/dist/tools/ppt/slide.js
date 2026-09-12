"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.slideTools = exports.alignObjectsHandler = exports.alignObjectsDefinition = exports.setFontColorHandler = exports.setFontColorDefinition = exports.unifyFontHandler = exports.unifyFontDefinition = exports.beautifyHandler = exports.beautifyDefinition = exports.addSlideHandler = exports.addSlideDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 添加新幻灯片
 * 可以指定布局、位置、标题和内容
 */
exports.addSlideDefinition = {
    name: 'wps_ppt_add_slide',
    description: `添加新幻灯片到演示文稿。

支持的布局类型：
- title: 标题页
- title_content: 标题+内容（最常用）
- blank: 空白页
- two_column: 两栏内容
- comparison: 对比布局

使用场景：
- "新建一页PPT"
- "添加一个标题页"
- "在第3页后面插入一页"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            layout: {
                type: 'string',
                description: '幻灯片布局类型',
                enum: ['title', 'title_content', 'blank', 'two_column', 'comparison'],
            },
            position: {
                type: 'number',
                description: '插入位置（页码），不填则在末尾添加',
            },
            title: {
                type: 'string',
                description: '幻灯片标题',
            },
            content: {
                type: 'string',
                description: '幻灯片内容（针对有内容区域的布局）',
            },
        },
        required: [],
    },
};
const addSlideHandler = async (args) => {
    const { layout, position, title, content } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addSlide', {
            layout: layout || 'title_content',
            position,
            title,
            content,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const layoutName = {
                title: '标题页',
                title_content: '标题+内容',
                blank: '空白页',
                two_column: '两栏内容',
                comparison: '对比布局',
            }[response.data.layout] || response.data.layout;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片添加成功！\n位置: 第 ${response.data.slideIndex} 页\n布局: ${layoutName}${title ? `\n标题: ${title}` : ''}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `添加幻灯片失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `添加幻灯片出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.addSlideHandler = addSlideHandler;
/**
 * 美化幻灯片
 * 一键优化排版、配色、字体和间距
 */
exports.beautifyDefinition = {
    name: 'wps_ppt_beautify',
    description: `一键美化幻灯片，优化排版、配色、字体和间距。

支持的配色方案：
- business: 商务风（深蓝+灰色）
- tech: 科技风（蓝色+绿色）
- creative: 创意风（珊瑚红+金色）
- minimal: 简约风（黑白灰）

美化包含的操作：
- 统一字体
- 应用配色方案
- 对齐元素
- 优化间距

使用场景：
- "美化这页PPT"
- "用商务风格优化一下"
- "把PPT弄好看点"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slide_index: {
                type: 'number',
                description: '要美化的幻灯片页码，不填则美化当前页',
            },
            color_scheme: {
                type: 'string',
                description: '配色方案',
                enum: ['business', 'tech', 'creative', 'minimal'],
            },
            font: {
                type: 'string',
                description: '统一使用的字体，如 "微软雅黑"、"思源黑体"',
            },
            beautify_all: {
                type: 'boolean',
                description: '是否美化所有幻灯片，默认false只美化指定页',
            },
        },
        required: [],
    },
};
const beautifyHandler = async (args) => {
    const { slide_index, color_scheme, font, beautify_all } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('beautifySlide', {
            slideIndex: beautify_all ? 'all' : slide_index,
            style: {
                colorScheme: color_scheme || 'business',
                font: font || '微软雅黑',
            },
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const result = response.data;
            const schemeName = {
                business: '商务风',
                tech: '科技风',
                creative: '创意风',
                minimal: '简约风',
            }[color_scheme || 'business'] || color_scheme;
            let output = `幻灯片美化完成！\n`;
            output += `范围: ${beautify_all ? '所有幻灯片' : `第 ${result.slideIndex} 页`}\n`;
            output += `配色方案: ${schemeName}\n`;
            output += `字体: ${font || '微软雅黑'}\n\n`;
            output += `优化详情：\n`;
            result.operations.forEach((op) => {
                const opName = {
                    unify_font: '统一字体',
                    apply_color_scheme: '应用配色',
                    align: '对齐元素',
                    optimize_spacing: '优化间距',
                }[op.operation] || op.operation;
                output += `- ${opName}: 处理了 ${op.count} 个元素\n`;
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
                content: [{ type: 'text', text: `美化幻灯片失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `美化幻灯片出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.beautifyHandler = beautifyHandler;
/**
 * 统一字体
 * 整个演示文稿的字体统一是PPT美观的基础
 */
exports.unifyFontDefinition = {
    name: 'wps_ppt_unify_font',
    description: `统一演示文稿中所有幻灯片的字体。

使用场景：
- "把所有页面的字体都改成微软雅黑"
- "统一字体"
- "换个字体"

常用字体推荐：
- 微软雅黑：现代简洁，适合商务
- 思源黑体：开源免费，适合各种场合
- 黑体：传统正式
- 宋体：适合正式文档`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            font_name: {
                type: 'string',
                description: '要统一使用的字体名称，如 "微软雅黑"、"思源黑体"',
            },
            slide_index: {
                type: 'number',
                description: '只处理指定页，不填则处理所有页',
            },
            include_title: {
                type: 'boolean',
                description: '是否包含标题，默认true',
            },
            include_body: {
                type: 'boolean',
                description: '是否包含正文，默认true',
            },
        },
        required: ['font_name'],
    },
};
const unifyFontHandler = async (args) => {
    const { font_name, slide_index, include_title, include_body } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('unifyFont', {
            fontName: font_name,
            slideIndex: slide_index,
            includeTitle: include_title !== false,
            includeBody: include_body !== false,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const result = response.data;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `字体统一完成！\n字体: ${result.font}\n处理幻灯片: ${result.slideCount} 页\n修改文本框: ${result.count} 个`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `统一字体失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `统一字体出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.unifyFontHandler = unifyFontHandler;
/**
 * 设置文字颜色
 * 修改幻灯片中指定形状的文字颜色
 */
exports.setFontColorDefinition = {
    name: 'wps_ppt_set_font_color',
    description: `设置幻灯片中指定形状的文字颜色。

使用场景：
- "把标题改成红色"
- "设置第2页第1个文本框的文字颜色为蓝色"
- "修改文字颜色"`,
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
            color: {
                type: 'string',
                description: '颜色值，支持十六进制如 "#FF0000" 或颜色名如 "red"',
            },
        },
        required: ['slideIndex', 'shapeIndex', 'color'],
    },
};
const setFontColorHandler = async (args) => {
    const { slideIndex, shapeIndex, color } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setFontColor', // NOTE: macOS未实现，仅Windows支持
        { slideIndex, shapeIndex, color }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `文字颜色设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n颜色: ${color}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置文字颜色失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置文字颜色出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setFontColorHandler = setFontColorHandler;
/**
 * 对齐幻灯片中的对象
 * 支持多种对齐方式
 */
exports.alignObjectsDefinition = {
    name: 'wps_ppt_align_objects',
    description: `对齐幻灯片中的对象。

支持的对齐方式：
- left: 左对齐
- center: 水平居中
- right: 右对齐
- top: 顶部对齐
- middle: 垂直居中
- bottom: 底部对齐
- distribute_h: 水平等距分布
- distribute_v: 垂直等距分布

使用场景：
- "把这些元素居中对齐"
- "让所有对象左对齐"
- "等距分布这些形状"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            alignment: {
                type: 'string',
                description: '对齐方式',
                enum: ['left', 'center', 'right', 'top', 'middle', 'bottom', 'distribute_h', 'distribute_v'],
            },
        },
        required: ['slideIndex', 'alignment'],
    },
};
const alignObjectsHandler = async (args) => {
    const { slideIndex, alignment } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('alignShapes', { slideIndex, alignment }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const alignName = {
                left: '左对齐',
                center: '水平居中',
                right: '右对齐',
                top: '顶部对齐',
                middle: '垂直居中',
                bottom: '底部对齐',
                distribute_h: '水平等距分布',
                distribute_v: '垂直等距分布',
            };
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `对象对齐完成！\n幻灯片: 第 ${slideIndex} 页\n对齐方式: ${alignName[alignment] || alignment}${response.data?.count ? `\n处理对象: ${response.data.count} 个` : ''}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `对齐对象失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `对齐对象出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.alignObjectsHandler = alignObjectsHandler;
/**
 * 导出所有幻灯片相关的Tools
 */
exports.slideTools = [
    { definition: exports.addSlideDefinition, handler: exports.addSlideHandler },
    { definition: exports.beautifyDefinition, handler: exports.beautifyHandler },
    { definition: exports.unifyFontDefinition, handler: exports.unifyFontHandler },
    { definition: exports.setFontColorDefinition, handler: exports.setFontColorHandler },
    { definition: exports.alignObjectsDefinition, handler: exports.alignObjectsHandler },
];
exports.default = exports.slideTools;
//# sourceMappingURL=slide.js.map
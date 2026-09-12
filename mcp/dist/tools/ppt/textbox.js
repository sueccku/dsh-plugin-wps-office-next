"use strict";
/**
 * Input: 文本框与标题工具参数
 * Output: 文本框操作结果
 * Pos: PPT 文本框与标题工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.textboxTools = exports.setSlideContentHandler = exports.setSlideContentDefinition = exports.setSlideSubtitleHandler = exports.setSlideSubtitleDefinition = exports.getSlideTitleHandler = exports.getSlideTitleDefinition = exports.setTextboxStyleHandler = exports.setTextboxStyleDefinition = exports.setTextboxTextHandler = exports.setTextboxTextDefinition = exports.getTextboxesHandler = exports.getTextboxesDefinition = exports.deleteTextboxHandler = exports.deleteTextboxDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 删除文本框
 */
exports.deleteTextboxDefinition = {
    name: 'wps_ppt_delete_textbox',
    description: '删除幻灯片上指定的文本框。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
            textboxIndex: { type: 'number', description: '文本框索引' },
        },
        required: ['slideIndex', 'textboxIndex'],
    },
};
const deleteTextboxHandler = async (args) => {
    const { slideIndex, textboxIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteTextBox', { slideIndex, textboxIndex }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `删除文本框失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `第${slideIndex}页的文本框${textboxIndex}已删除` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `删除文本框出错: ${errMsg}` }], error: errMsg };
    }
};
exports.deleteTextboxHandler = deleteTextboxHandler;
/**
 * 获取文本框列表
 */
exports.getTextboxesDefinition = {
    name: 'wps_ppt_get_textboxes',
    description: '获取幻灯片上所有文本框的列表。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
        },
        required: ['slideIndex'],
    },
};
const getTextboxesHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getTextBoxes', { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取文本框列表失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `第${slideIndex}页文本框:\n${JSON.stringify(response.data?.textboxes || [], null, 2)}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取文本框列表出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getTextboxesHandler = getTextboxesHandler;
/**
 * 设置文本框文本
 */
exports.setTextboxTextDefinition = {
    name: 'wps_ppt_set_textbox_text',
    description: '设置指定文本框的文本内容。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
            textboxIndex: { type: 'number', description: '文本框索引' },
            text: { type: 'string', description: '文本内容' },
        },
        required: ['slideIndex', 'textboxIndex', 'text'],
    },
};
const setTextboxTextHandler = async (args) => {
    const { slideIndex, textboxIndex, text } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setTextBoxText', { slideIndex, textboxIndex, text }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置文本框内容失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `文本框内容已更新` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置文本框内容出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setTextboxTextHandler = setTextboxTextHandler;
/**
 * 设置文本框样式
 */
exports.setTextboxStyleDefinition = {
    name: 'wps_ppt_set_textbox_style',
    description: '设置文本框样式（字体大小、颜色、粗体等）。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
            textboxIndex: { type: 'number', description: '文本框索引' },
            style: {
                type: 'object',
                description: '样式对象，包含 fontSize/fontColor/bold/italic/align 等',
            },
        },
        required: ['slideIndex', 'textboxIndex', 'style'],
    },
};
const setTextboxStyleHandler = async (args) => {
    const { slideIndex, textboxIndex, style } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setTextBoxStyle', { slideIndex, textboxIndex, style }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置文本框样式失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `文本框样式已更新` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置文本框样式出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setTextboxStyleHandler = setTextboxStyleHandler;
/**
 * 获取幻灯片标题
 */
exports.getSlideTitleDefinition = {
    name: 'wps_ppt_get_slide_title',
    description: '获取指定幻灯片的标题。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
        },
        required: ['slideIndex'],
    },
};
const getSlideTitleHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSlideTitle', { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取标题失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `第${slideIndex}页标题: ${response.data?.title || '无标题'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取标题出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getSlideTitleHandler = getSlideTitleHandler;
/**
 * 设置幻灯片副标题
 */
exports.setSlideSubtitleDefinition = {
    name: 'wps_ppt_set_slide_subtitle',
    description: '设置幻灯片的副标题。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
            subtitle: { type: 'string', description: '副标题内容' },
        },
        required: ['slideIndex', 'subtitle'],
    },
};
const setSlideSubtitleHandler = async (args) => {
    const { slideIndex, subtitle } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSlideSubtitle', { slideIndex, subtitle }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置副标题失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `第${slideIndex}页副标题已设置` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置副标题出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setSlideSubtitleHandler = setSlideSubtitleHandler;
/**
 * 设置幻灯片正文内容
 */
exports.setSlideContentDefinition = {
    name: 'wps_ppt_set_slide_content',
    description: '设置幻灯片的正文内容区域。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
            content: { type: 'string', description: '正文内容' },
        },
        required: ['slideIndex', 'content'],
    },
};
const setSlideContentHandler = async (args) => {
    const { slideIndex, content } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSlideContent', { slideIndex, content }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置正文内容失败: ${response.error}` }], error: response.error };
        }
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `第${slideIndex}页正文内容已设置` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置正文内容出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setSlideContentHandler = setSlideContentHandler;
/**
 * 导出所有文本框与标题相关的Tools
 */
exports.textboxTools = [
    { definition: exports.deleteTextboxDefinition, handler: exports.deleteTextboxHandler },
    { definition: exports.getTextboxesDefinition, handler: exports.getTextboxesHandler },
    { definition: exports.setTextboxTextDefinition, handler: exports.setTextboxTextHandler },
    { definition: exports.setTextboxStyleDefinition, handler: exports.setTextboxStyleHandler },
    { definition: exports.getSlideTitleDefinition, handler: exports.getSlideTitleHandler },
    { definition: exports.setSlideSubtitleDefinition, handler: exports.setSlideSubtitleHandler },
    { definition: exports.setSlideContentDefinition, handler: exports.setSlideContentHandler },
];
exports.default = exports.textboxTools;
//# sourceMappingURL=textbox.js.map
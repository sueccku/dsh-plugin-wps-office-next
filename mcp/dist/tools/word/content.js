"use strict";
/**
 * Input: Word 内容操作参数
 * Output: 文档内容变更结果
 * Pos: Word 内容工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word内容操作Tools - 文档内容编辑模块
 * 处理文档内容的插入、查找替换、表格、段落格式、页眉页脚等操作
 *
 * 包含：
 * - wps_word_insert_text: 插入文本到文档
 * - wps_word_find_replace: 查找替换功能
 * - wps_word_insert_table: 插入表格
 * - wps_word_set_paragraph: 设置段落格式
 * - wps_word_insert_header: 插入页眉
 * - wps_word_insert_footer: 插入页脚
 * - wps_word_get_active_document: 获取当前文档信息
 * - wps_word_insert_page_break: 插入分页符
 * - wps_word_insert_comment: 插入批注
 * - wps_word_set_font_style: 设置选中文字的字体样式属性
 * - wps_word_set_text_color: 设置文字颜色
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.contentTools = exports.replaceBookmarkContentHandler = exports.replaceBookmarkContentDefinition = exports.smartFillFieldHandler = exports.smartFillFieldDefinition = exports.findInDocumentHandler = exports.findInDocumentDefinition = exports.getParagraphsHandler = exports.getParagraphsDefinition = exports.setTextColorHandler = exports.setTextColorDefinition = exports.insertCommentHandler = exports.insertCommentDefinition = exports.setFontStyleHandler = exports.setFontStyleDefinition = exports.insertPageBreakHandler = exports.insertPageBreakDefinition = exports.insertImageHandler = exports.insertImageDefinition = exports.getActiveDocumentHandler = exports.getActiveDocumentDefinition = exports.setParagraphHandler = exports.setParagraphDefinition = exports.insertTableHandler = exports.insertTableDefinition = exports.findReplaceHandler = exports.findReplaceDefinition = exports.insertTextHandler = exports.insertTextDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 插入文本到文档
 * 可以在光标位置、文档开头或结尾插入文本
 */
exports.insertTextDefinition = {
    name: 'wps_word_insert_text',
    description: `在Word文档中插入文本。

使用场景：
- "在文档开头加个标题"
- "在光标位置插入这段话"
- "在文档末尾添加备注"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            text: {
                type: 'string',
                description: '要插入的文本内容',
            },
            position: {
                type: 'string',
                description: '插入位置，可选值: "cursor"(光标位置), "start"(文档开头), "end"(文档结尾)。默认cursor',
                enum: ['cursor', 'start', 'end'],
            },
            style: {
                type: 'string',
                description: '插入后应用的样式，如 "标题 1"、"正文"',
            },
            new_paragraph: {
                type: 'boolean',
                description: '插入后是否新起一段，默认false',
            },
        },
        required: ['text'],
    },
};
const insertTextHandler = async (args) => {
    const { text, position, style, new_paragraph } = args;
    if (!text || text.trim() === '') {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '要插入的文本不能为空！' }],
            error: '文本内容为空',
        };
    }
    try {
        // 处理换行符，如果需要新段落
        const finalText = new_paragraph ? text + '\n' : text;
        const response = await wps_client_1.wpsClient.executeMethod('insertText', {
            text: finalText,
            position: position || 'cursor',
            style,
        }, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            const positionText = position === 'start' ? '文档开头' :
                position === 'end' ? '文档结尾' : '光标位置';
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `文本插入成功！\n位置: ${positionText}\n字符数: ${response.data.textLength}${style ? `\n应用样式: ${style}` : ''}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `插入文本失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `插入文本出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertTextHandler = insertTextHandler;
/**
 * 查找替换功能
 * 支持全文批量查找替换文本
 */
exports.findReplaceDefinition = {
    name: 'wps_word_find_replace',
    description: `在Word文档中查找并替换文本。

使用场景：
- "把所有的'公司'替换成'集团'"
- "把文档里的错别字改过来"
- "批量替换某个词"

支持选项：
- 区分大小写
- 全字匹配
- 全部替换或仅替换一处`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            find_text: {
                type: 'string',
                description: '要查找的文本',
            },
            replace_text: {
                type: 'string',
                description: '替换为的文本，如果只是查找不替换，可以不填',
            },
            replace_all: {
                type: 'boolean',
                description: '是否全部替换，默认true',
            },
            match_case: {
                type: 'boolean',
                description: '是否区分大小写，默认false',
            },
            match_whole_word: {
                type: 'boolean',
                description: '是否全字匹配，默认false',
            },
        },
        required: ['find_text'],
    },
};
const findReplaceHandler = async (args) => {
    const { find_text, replace_text, replace_all, match_case, match_whole_word } = args;
    if (!find_text || find_text.trim() === '') {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '查找文本不能为空！' }],
            error: '查找文本为空',
        };
    }
    try {
        // 如果没有替换文本，就只是查找
        const isReplaceMode = replace_text !== undefined && replace_text !== null;
        const response = await wps_client_1.wpsClient.executeMethod('findReplace', {
            findText: find_text,
            // replaceText is only meaningful in replace mode: sending an empty string while merely
            // searching used to be applied as a replacement and erased every match.
            replaceText: isReplaceMode ? replace_text : undefined,
            replaceAll: replace_all !== false, // 默认true
            matchCase: match_case || false,
            matchWholeWord: match_whole_word || false,
            replaceMode: isReplaceMode,
        }, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            const result = response.data;
            const found = typeof result.count === 'number' ? result.count : result.found;
            if (isReplaceMode) {
                if (found === 0) {
                    return {
                        id: (0, uuid_1.v4)(),
                        success: true,
                        content: [
                            {
                                type: 'text',
                                text: `未找到 "${find_text}"，没有进行替换`,
                            },
                        ],
                    };
                }
                const replaced = typeof found === 'number' ? `共 ${found} 处` : '已执行替换';
                return {
                    id: (0, uuid_1.v4)(),
                    success: true,
                    content: [
                        {
                            type: 'text',
                            text: `替换完成！\n查找: "${find_text}"\n替换为: "${replace_text}"\n${replaced}`,
                        },
                    ],
                };
            }
            // A search must report the real match count: the bridge counts matches without touching
            // the document, so an absent number here means the bridge is older, not that nothing matched.
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: typeof found === 'number'
                            ? `查找完成！\n"${find_text}" 在文档中出现了 ${found} 次`
                            : `查找完成！\n已查找 "${find_text}"（桥未回报匹配次数）`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `查找替换失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `查找替换出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.findReplaceHandler = findReplaceHandler;
/**
 * 插入表格到文档光标位置
 */
exports.insertTableDefinition = {
    name: 'wps_word_insert_table',
    description: '在Word文档光标位置插入表格',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            rows: { type: 'number', description: '表格行数' },
            cols: { type: 'number', description: '表格列数' },
        },
        required: ['rows', 'cols'],
    },
};
const insertTableHandler = async (args) => {
    const { rows, cols } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertTable', { rows, cols }, wps_1.WpsAppType.WRITER);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `已插入 ${rows}×${cols} 表格` }],
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `插入表格失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入表格出错: ${errMsg}` }], error: errMsg };
    }
};
exports.insertTableHandler = insertTableHandler;
/**
 * 设置当前段落格式
 */
exports.setParagraphDefinition = {
    name: 'wps_word_set_paragraph',
    description: '设置当前段落格式（对齐方式、行间距等）',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            alignment: { type: 'string', description: '对齐方式: left/center/right/justify' },
            lineSpacing: { type: 'number', description: '行间距倍数，如1.5、2' },
        },
    },
};
const setParagraphHandler = async (args) => {
    const params = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setParagraph', params, wps_1.WpsAppType.WRITER);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: '段落格式已设置' }],
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置段落失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置段落出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setParagraphHandler = setParagraphHandler;
/**
 * 获取当前活动文档信息
 */
exports.getActiveDocumentDefinition = {
    name: 'wps_word_get_active_document',
    description: '获取当前WPS Writer活动文档的基本信息',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {},
    },
};
const getActiveDocumentHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getActiveDocument', {}, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            const d = response.data;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `当前文档: ${d.name}\n路径: ${d.path}\n页数: ${d.pageCount}\n字数: ${d.wordCount}` }],
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取文档信息失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `获取文档信息出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getActiveDocumentHandler = getActiveDocumentHandler;
/**
 * 在文档中插入图片
 */
exports.insertImageDefinition = {
    name: 'wps_word_insert_image',
    description: `在Word文档中插入图片。

使用场景：
- "在文档中插入一张图片"
- "把这个截图放到文档里"
- "在光标位置插入logo"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            imagePath: {
                type: 'string',
                description: '图片文件路径',
            },
            width: {
                type: 'number',
                description: '图片宽度（磅），可选',
            },
            height: {
                type: 'number',
                description: '图片高度（磅），可选',
            },
        },
        required: ['imagePath'],
    },
};
const insertImageHandler = async (args) => {
    const { imagePath, width, height } = args;
    if (!imagePath || imagePath.trim() === '') {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '图片路径不能为空！' }],
            error: '图片路径为空',
        };
    }
    try {
        // The bridge reads "path"; the imagePath/filePath aliases were never read.
        const response = await wps_client_1.wpsClient.executeMethod('insertImage', { path: imagePath, width, height }, wps_1.WpsAppType.WRITER);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `图片插入成功！\n路径: ${imagePath}${width ? `\n宽度: ${width}磅` : ''}${height ? `\n高度: ${height}磅` : ''}`,
                    },
                ],
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `插入图片失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `插入图片出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertImageHandler = insertImageHandler;
/**
 * 导出所有内容操作相关的Tools
 */
exports.insertPageBreakDefinition = {
    name: 'wps_word_insert_page_break',
    description: '在文档光标位置插入分页符',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const insertPageBreakHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertPageBreak', {}, wps_1.WpsAppType.WRITER);
        return {
            id: (0, uuid_1.v4)(),
            success: response.success,
            content: [{ type: 'text', text: response.success ? '分页符已插入' : response.data?.message || '插入失败' }],
        };
    }
    catch (e) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入分页符出错: ${e.message}` }], error: e.message };
    }
};
exports.insertPageBreakHandler = insertPageBreakHandler;
exports.setFontStyleDefinition = {
    name: 'wps_word_set_font_style',
    description: '设置选中文字的字体样式属性',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            fontName: { type: 'string', description: '字体名称' },
            fontSize: { type: 'number', description: '字号' },
            bold: { type: 'boolean', description: '是否加粗' },
            italic: { type: 'boolean', description: '是否斜体' },
        },
        required: [],
    },
};
const setFontStyleHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setFont', args, wps_1.WpsAppType.WRITER);
        return {
            id: (0, uuid_1.v4)(),
            success: response.success,
            content: [{ type: 'text', text: response.success ? '字体已设置' : response.data?.message || '设置失败' }],
        };
    }
    catch (e) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置字体出错: ${e.message}` }], error: e.message };
    }
};
exports.setFontStyleHandler = setFontStyleHandler;
/**
 * 插入批注到文档选中内容
 */
exports.insertCommentDefinition = {
    name: 'wps_word_insert_comment',
    description: '在Word文档选中内容处插入批注',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            text: { type: 'string', description: '批注内容' },
        },
        required: ['text'],
    },
};
const insertCommentHandler = async (args) => {
    const { text } = args;
    if (!text || text.trim() === '') {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '批注内容不能为空！' }], error: '批注内容为空' };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addComment', { text }, wps_1.WpsAppType.WRITER);
        return {
            id: (0, uuid_1.v4)(),
            success: response.success,
            content: [{ type: 'text', text: response.success ? '批注已插入' : response.data?.message || '插入失败' }],
        };
    }
    catch (e) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入批注出错: ${e.message}` }], error: e.message };
    }
};
exports.insertCommentHandler = insertCommentHandler;
/**
 * 设置选中文字的颜色
 */
exports.setTextColorDefinition = {
    name: 'wps_word_set_text_color',
    description: '设置Word文档中选中文字的颜色',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            color: { type: 'string', description: '颜色值，如 "#FF0000"、"red"' },
        },
        required: ['color'],
    },
};
const setTextColorHandler = async (args) => {
    const { color } = args;
    if (!color || color.trim() === '') {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '颜色值不能为空！' }], error: '颜色值为空' };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setTextColor', // NOTE: macOS未实现，仅Windows支持
        { color }, wps_1.WpsAppType.WRITER);
        return {
            id: (0, uuid_1.v4)(),
            success: response.success,
            content: [{ type: 'text', text: response.success ? `文字颜色已设置为 ${color}` : response.data?.message || '设置失败' }],
        };
    }
    catch (e) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置文字颜色出错: ${e.message}` }], error: e.message };
    }
};
exports.setTextColorHandler = setTextColorHandler;
/**
 * 获取文档段落结构
 * 返回段落的文本、样式、位置信息，用于了解文档结构和识别填写位置
 */
exports.getParagraphsDefinition = {
    name: 'wps_word_get_paragraphs',
    description: `获取Word文档的段落结构信息，返回每段的文本、样式和字符位置。

使用场景：
- "了解文档的结构"
- "查看文档有哪些段落"
- "帮我看看模板里有哪些需要填写的位置"
- 在填写模板前，先读取文档结构以识别填写位置

返回信息包括：段落索引、文本内容、样式名称、字符起止位置。
支持分页获取（startParagraph/endParagraph），默认返回前50段。`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            start_paragraph: { type: 'number', description: '起始段落索引（从1开始），默认1' },
            end_paragraph: { type: 'number', description: '结束段落索引，默认为起始+49' },
        },
        required: [],
    },
};
const getParagraphsHandler = async (args) => {
    const { start_paragraph, end_paragraph } = args;
    try {
        const execParams = {};
        if (start_paragraph != null)
            execParams.startParagraph = start_paragraph;
        if (end_paragraph != null)
            execParams.endParagraph = end_paragraph;
        const response = await wps_client_1.wpsClient.executeMethod('getDocumentParagraphs', execParams, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            const { paragraphs, totalCount, returnedCount } = response.data;
            const lines = paragraphs.map((p) => `[${p.index}] (${p.style}) ${p.text}`);
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `文档段落结构（共${totalCount}段，返回${returnedCount}段）：\n${lines.join('\n')}` }],
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取段落结构失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取段落结构出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getParagraphsHandler = getParagraphsHandler;
/**
 * 查找文本返回位置（不替换）
 * 返回匹配文本的位置信息和上下文，供后续精确操作使用
 */
exports.findInDocumentDefinition = {
    name: 'wps_word_find_in_document',
    description: `在Word文档中查找文本并返回位置信息，不执行替换操作。

使用场景：
- "找一下'项目名称'在文档的哪个位置"
- "看看文档里有哪些地方需要填写"
- 在使用smart_fill_field之前，先用此工具定位关键字

返回信息包括：匹配文本、字符起止位置、所在段落索引、上下文（前后50字符）。
与find_replace不同，此工具仅查找不替换，返回位置信息供后续操作使用。`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            find_text: { type: 'string', description: '要查找的文本' },
            match_case: { type: 'boolean', description: '是否区分大小写，默认false' },
            match_whole_word: { type: 'boolean', description: '是否全字匹配，默认false' },
            max_results: { type: 'number', description: '最大返回结果数，默认20' },
        },
        required: ['find_text'],
    },
};
const findInDocumentHandler = async (args) => {
    const { find_text, match_case, match_whole_word, max_results } = args;
    if (!find_text || find_text.trim() === '') {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '查找文本不能为空！' }],
            error: '查找文本为空',
        };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('findInDocument', {
            findText: find_text,
            matchCase: match_case || false,
            matchWholeWord: match_whole_word || false,
            maxResults: max_results || 20,
        }, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            const { results, count, findText } = response.data;
            if (count === 0) {
                return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `未找到 "${findText}"` }] };
            }
            const lines = results.map((r) => `  段落${r.paragraphIndex} [${r.start}-${r.end}]: ...${r.context}...`);
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `找到 "${findText}" 共${count}处：\n${lines.join('\n')}` }],
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `查找失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `查找出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.findInDocumentHandler = findInDocumentHandler;
/**
 * 智能填写模板字段
 * 自动判断关键字附近的填写模式（下划线/冒号后/标签后/占位符），在正确位置插入内容并保持格式
 */
exports.smartFillFieldDefinition = {
    name: 'wps_word_smart_fill_field',
    description: `智能填写Word模板中的字段。自动识别关键字附近的填写模式，在正确位置插入内容并保持原有格式。

使用场景：
- "把项目名称填写为'XX信息化项目'"
- "填写建设单位为'XX公司'"
- "模板里的项目编号填上'2026-001'"
- 任何需要在模板文档中"填写"而非"替换"的场景

支持的填写模式（fill_mode，默认auto自动判断）：
- auto: 自动判断（推荐）
- underline: 关键字后有下划线___，替换下划线为填写内容
- afterColon: 关键字后有冒号（：或:），在冒号后插入
- afterLabel: 关键字是标签，直接在关键字后插入
- placeholder: 关键字被{}或【】包裹，替换整个占位符

重要：模板填写场景应优先使用此工具，而非find_replace。find_replace会删除关键字本身并可能破坏格式。`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            keyword: { type: 'string', description: '要填写的关键字（如"项目名称"、"建设单位"）' },
            value: { type: 'string', description: '要填写的值（如"XX信息化项目"）' },
            fill_mode: {
                type: 'string',
                description: '填写模式: auto(自动判断), underline(下划线), afterColon(冒号后), afterLabel(标签后), placeholder(占位符)。默认auto',
                enum: ['auto', 'underline', 'afterColon', 'afterLabel', 'placeholder'],
            },
        },
        required: ['keyword', 'value'],
    },
};
const smartFillFieldHandler = async (args) => {
    const { keyword, value, fill_mode } = args;
    if (!keyword || keyword.trim() === '') {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '关键字不能为空！' }], error: '关键字为空' };
    }
    if (value == null) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '填写值不能为空（空字符串将清除该字段内容）' }], error: '填写值为空' };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('smartFillField', { keyword, value, fillMode: fill_mode || 'auto' }, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            const { fillMode, result } = response.data;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `填写完成！\n关键字: "${keyword}"\n填写值: "${value}"\n填写模式: ${fillMode}\n结果: ${result}` }],
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `填写失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `填写出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.smartFillFieldHandler = smartFillFieldHandler;
/**
 * 替换书签内容
 * 通过书签名定位并替换内容，保持书签位置格式
 */
exports.replaceBookmarkContentDefinition = {
    name: 'wps_word_replace_bookmark_content',
    description: `替换Word文档中书签的内容。通过书签名定位，替换书签范围内的文本，保持原有格式。

使用场景：
- "把书签'project_name'的内容改为'XX项目'"
- 模板文档使用书签标记填写位置时使用

注意：替换后书签会自动重建，不会丢失。`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            name: { type: 'string', description: '书签名称' },
            text: { type: 'string', description: '要替换为的文本内容' },
        },
        required: ['name', 'text'],
    },
};
const replaceBookmarkContentHandler = async (args) => {
    const { name, text } = args;
    if (!name || name.trim() === '') {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '书签名称不能为空！' }], error: '书签名称为空' };
    }
    if (text == null) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '替换文本不能为空（空字符串将清空书签）' }], error: '替换文本为空' };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('replaceBookmarkContent', { name, text }, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `书签内容替换完成！\n书签: "${name}"\n新内容: "${text}"` }],
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `替换书签内容失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `替换书签内容出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.replaceBookmarkContentHandler = replaceBookmarkContentHandler;
exports.contentTools = [
    { definition: exports.insertTextDefinition, handler: exports.insertTextHandler },
    { definition: exports.findReplaceDefinition, handler: exports.findReplaceHandler },
    { definition: exports.insertTableDefinition, handler: exports.insertTableHandler },
    { definition: exports.setParagraphDefinition, handler: exports.setParagraphHandler },
    { definition: exports.getActiveDocumentDefinition, handler: exports.getActiveDocumentHandler },
    { definition: exports.insertImageDefinition, handler: exports.insertImageHandler },
    { definition: exports.insertPageBreakDefinition, handler: exports.insertPageBreakHandler },
    { definition: exports.setFontStyleDefinition, handler: exports.setFontStyleHandler },
    { definition: exports.insertCommentDefinition, handler: exports.insertCommentHandler },
    { definition: exports.setTextColorDefinition, handler: exports.setTextColorHandler },
    { definition: exports.getParagraphsDefinition, handler: exports.getParagraphsHandler },
    { definition: exports.findInDocumentDefinition, handler: exports.findInDocumentHandler },
    { definition: exports.smartFillFieldDefinition, handler: exports.smartFillFieldHandler },
    { definition: exports.replaceBookmarkContentDefinition, handler: exports.replaceBookmarkContentHandler },
];
exports.default = exports.contentTools;
//# sourceMappingURL=content.js.map
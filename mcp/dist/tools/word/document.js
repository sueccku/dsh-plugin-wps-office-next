"use strict";
/**
 * Input: 文档管理工具参数
 * Output: 文档操作结果
 * Pos: Word 文档管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word文档管理Tools
 * 处理文档的打开、切换、获取列表等管理操作
 *
 * 包含：
 * - wps_word_get_open_documents: 获取所有已打开的文档列表
 * - wps_word_switch_document: 切换到指定文档
 * - wps_word_open_document: 打开指定路径的文档
 * - wps_word_create_document: 新建空白文档
 * - wps_word_close_document: 关闭文档（可选保存）
 * - wps_word_get_document_text: 获取文档文本内容
 * - wps_word_insert_header: 设置页眉内容
 * - wps_word_insert_footer: 设置页脚内容
 * - wps_word_generate_doc_toc: 自动生成文档目录
 * - wps_word_insert_section_break: 插入分节符
 * - wps_word_set_line_spacing: 设置行距
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.documentTools = exports.closeDocumentHandler = exports.closeDocumentDefinition = exports.createDocumentHandler = exports.createDocumentDefinition = exports.setLineSpacingHandler = exports.setLineSpacingDefinition = exports.insertSectionBreakHandler = exports.insertSectionBreakDefinition = exports.generateDocTocHandler = exports.generateDocTocDefinition = exports.insertFooterHandler = exports.insertFooterDefinition = exports.insertHeaderHandler = exports.insertHeaderDefinition = exports.getDocumentTextHandler = exports.getDocumentTextDefinition = exports.openDocumentHandler = exports.openDocumentDefinition = exports.switchDocumentHandler = exports.switchDocumentDefinition = exports.getOpenDocumentsHandler = exports.getOpenDocumentsDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 获取所有已打开的文档列表
 */
exports.getOpenDocumentsDefinition = {
    name: 'wps_word_get_open_documents',
    description: `获取当前WPS Writer中所有已打开的文档列表。

使用场景：
- "看看现在打开了哪些文档"
- "列出所有打开的Word文件"
- "查看当前文档列表"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {},
    },
};
const getOpenDocumentsHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getOpenDocuments', {}, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            const docs = response.data.documents;
            if (!docs || docs.length === 0) {
                return {
                    id: (0, uuid_1.v4)(),
                    success: true,
                    content: [{ type: 'text', text: '当前没有打开任何文档' }],
                };
            }
            const docList = docs
                .map((doc, i) => `${i + 1}. ${doc.name}${doc.active ? ' (当前活动)' : ''}\n   路径: ${doc.path}`)
                .join('\n');
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `已打开的文档列表 (共${docs.length}个):\n${docList}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取文档列表失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取文档列表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getOpenDocumentsHandler = getOpenDocumentsHandler;
/**
 * 切换到指定文档
 */
exports.switchDocumentDefinition = {
    name: 'wps_word_switch_document',
    description: `切换到指定名称的文档。

使用场景：
- "切换到报告.docx"
- "打开另一个文档窗口"
- "切换到那个合同文档"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            name: {
                type: 'string',
                description: '要切换到的文档名称',
            },
        },
        required: ['name'],
    },
};
const switchDocumentHandler = async (args) => {
    const { name } = args;
    if (!name || name.trim() === '') {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '文档名称不能为空！' }],
            error: '文档名称为空',
        };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('switchDocument', { name }, wps_1.WpsAppType.WRITER);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `已切换到文档: ${name}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `切换文档失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `切换文档出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.switchDocumentHandler = switchDocumentHandler;
/**
 * 打开指定路径的文档
 */
exports.openDocumentDefinition = {
    name: 'wps_word_open_document',
    description: `打开指定路径的Word文档。

使用场景：
- "打开桌面上的报告.docx"
- "帮我打开这个文件路径的文档"
- "加载指定位置的Word文件"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            filePath: {
                type: 'string',
                description: '要打开的文档文件路径',
            },
        },
        required: ['filePath'],
    },
};
const openDocumentHandler = async (args) => {
    const { filePath } = args;
    if (!filePath || filePath.trim() === '') {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '文件路径不能为空！' }],
            error: '文件路径为空',
        };
    }
    try {
        const params = {
            path: filePath,
        };
        const response = await wps_client_1.wpsClient.executeMethod('openDocument', params, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `文档打开成功！\n文件: ${response.data.documentName || filePath}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `打开文档失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `打开文档出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.openDocumentHandler = openDocumentHandler;
/**
 * 获取文档文本内容
 */
exports.getDocumentTextDefinition = {
    name: 'wps_word_get_document_text',
    description: `获取当前Word文档的文本内容。

使用场景：
- "读取文档内容"
- "获取文档的前100个字符"
- "查看文档从第50到第200个字符的内容"

可指定起始和结束位置来获取部分文本。`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            start: {
                type: 'number',
                description: '起始位置（字符索引），默认从头开始',
            },
            end: {
                type: 'number',
                description: '结束位置（字符索引），默认到文档末尾',
            },
        },
    },
};
const getDocumentTextHandler = async (args) => {
    const { start, end } = args;
    try {
        const params = {};
        if (start !== undefined)
            params.start = start;
        if (end !== undefined)
            params.end = end;
        const response = await wps_client_1.wpsClient.executeMethod('getDocumentText', params, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            const rangeInfo = start !== undefined || end !== undefined
                ? `\n范围: ${start ?? 0} - ${end ?? '末尾'}`
                : '';
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `文档文本内容 (${response.data.length}字符)${rangeInfo}:\n\n${response.data.text}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取文档文本失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取文档文本出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getDocumentTextHandler = getDocumentTextHandler;
/**
 * 设置页眉内容
 */
exports.insertHeaderDefinition = {
    name: 'wps_word_insert_header',
    description: `设置页眉内容。

使用场景：
- "给文档加个页眉"
- "设置页眉为公司名称"
- "修改第2节的页眉"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            text: {
                type: 'string',
                description: '页眉文本内容',
            },
            section: {
                type: 'number',
                description: '节编号（从1开始），默认1',
                default: 1,
            },
        },
        required: ['text'],
    },
};
const insertHeaderHandler = async (args) => {
    const { text, section = 1 } = args;
    if (!text) {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '页眉文本不能为空！' }],
            error: '页眉文本为空',
        };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertHeader', { text, section }, wps_1.WpsAppType.WRITER);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `页眉已设置: "${text}" (第${section}节)` }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置页眉失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置页眉出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertHeaderHandler = insertHeaderHandler;
/**
 * 设置页脚内容
 */
exports.insertFooterDefinition = {
    name: 'wps_word_insert_footer',
    description: `设置页脚内容。

使用场景：
- "给文档加个页脚"
- "设置页脚为页码"
- "修改第1节的页脚"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            text: {
                type: 'string',
                description: '页脚文本内容',
            },
            section: {
                type: 'number',
                description: '节编号（从1开始），默认1',
                default: 1,
            },
        },
        required: ['text'],
    },
};
const insertFooterHandler = async (args) => {
    const { text, section = 1 } = args;
    if (!text) {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '页脚文本不能为空！' }],
            error: '页脚文本为空',
        };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertFooter', { text, section }, wps_1.WpsAppType.WRITER);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `页脚已设置: "${text}" (第${section}节)` }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置页脚失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置页脚出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertFooterHandler = insertFooterHandler;
/**
 * 自动生成文档目录
 */
exports.generateDocTocDefinition = {
    name: 'wps_word_generate_doc_toc',
    description: `自动生成文档目录。根据文档中的标题样式自动生成目录。

前提条件：文档中的标题必须使用"标题1"、"标题2"等样式。

使用场景：
- "帮我生成目录"
- "在文档开头插入目录"
- "自动生成文档目录"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            levels: {
                type: 'number',
                description: '目录包含的标题级别数，如3表示包含标题1-3，默认3',
                default: 3,
            },
        },
    },
};
const generateDocTocHandler = async (args) => {
    const { levels = 3 } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('generateTOC', { levels }, wps_1.WpsAppType.WRITER);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `目录已生成（包含标题1-${levels}级）` }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `生成目录失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `生成目录出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.generateDocTocHandler = generateDocTocHandler;
/**
 * 插入分节符
 */
exports.insertSectionBreakDefinition = {
    name: 'wps_word_insert_section_break',
    description: `插入分节符，用于将文档分为不同的节，以便对各节应用不同的页面设置。

使用场景：
- "插入一个分节符"
- "从下一页开始新的一节"
- "在这里分节"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            breakType: {
                type: 'string',
                description: '分节符类型：nextPage(下一页)、continuous(连续)、evenPage(偶数页)、oddPage(奇数页)，默认nextPage',
                default: 'nextPage',
            },
        },
    },
};
const insertSectionBreakHandler = async (args) => {
    const { breakType = 'nextPage' } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertSectionBreak', // NOTE: macOS未实现，仅Windows支持
        { breakType }, wps_1.WpsAppType.WRITER);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `分节符已插入（类型: ${breakType}）` }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `插入分节符失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `插入分节符出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertSectionBreakHandler = insertSectionBreakHandler;
/**
 * 设置行距
 */
exports.setLineSpacingDefinition = {
    name: 'wps_word_set_line_spacing',
    description: `设置段落行距。

使用场景：
- "把行距设为1.5倍"
- "设置第3段的行距为2倍"
- "调整行距"`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            lineSpacing: {
                type: 'number',
                description: '行距值（如1.0、1.5、2.0等）',
            },
            paragraphIndex: {
                type: 'number',
                description: '段落索引（从0开始），不指定则应用于当前段落或全文',
            },
        },
        required: ['lineSpacing'],
    },
};
const setLineSpacingHandler = async (args) => {
    const { lineSpacing, paragraphIndex } = args;
    if (lineSpacing === undefined || lineSpacing <= 0) {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '行距值必须为正数！' }],
            error: '行距值无效',
        };
    }
    try {
        const params = { lineSpacing };
        if (paragraphIndex !== undefined)
            params.paragraphIndex = paragraphIndex;
        const response = await wps_client_1.wpsClient.executeMethod('setLineSpacing', // NOTE: macOS未实现，仅Windows支持
        params, wps_1.WpsAppType.WRITER);
        if (response.success) {
            const target = paragraphIndex !== undefined ? `第${paragraphIndex}段` : '当前段落';
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `行距已设置为 ${lineSpacing} 倍（${target}）` }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置行距失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置行距出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setLineSpacingHandler = setLineSpacingHandler;
/**
 * 新建空白文档
 */
exports.createDocumentDefinition = {
    name: 'wps_word_create_document',
    description: `新建一个空白 Word 文档（不是打开已有文件）。

使用场景：
- "新建一个 Word 文档"
- "把结论写进一个新文档里"
- 需要从零起草，而不是编辑现有文档时

新建后用 wps_word_insert_text 写内容，用 wps_common_save_as 保存到磁盘。`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {},
    },
};
const createDocumentHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('createDocument', {}, wps_1.WpsAppType.WRITER);
        if (response.success) {
            const name = response.data?.name;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `已新建空白文档${name ? ': ' + name : ''}` }],
            };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `新建文档失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `新建文档出错: ${errMsg}` }], error: errMsg };
    }
};
exports.createDocumentHandler = createDocumentHandler;
/**
 * 关闭文档
 */
exports.closeDocumentDefinition = {
    name: 'wps_word_close_document',
    description: `关闭 Word 文档，可选是否保存。

使用场景：
- "关掉这个文档，别留着"
- 一批任务收尾时清理打开的文档

从未保存到磁盘的文档不会被强制保存（不会弹出保存对话框），此时结果里会带 warning 说明。`,
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            name: { type: 'string', description: '要关闭的文档名称；不填则关闭当前活动文档' },
            save: { type: 'boolean', description: '是否保存后关闭，默认 true' },
        },
    },
};
const closeDocumentHandler = async (args) => {
    const { name, save } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('closeDocument', { name, save }, wps_1.WpsAppType.WRITER);
        if (response.success && response.data) {
            const d = response.data;
            const lines = [`文档已关闭: ${d.closed}（${d.saved ? '已保存' : '未保存'}）`];
            if (d.warning)
                lines.push('注意: ' + d.warning);
            return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
        }
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `关闭文档失败: ${response.error}` }],
            error: response.error,
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `关闭文档出错: ${errMsg}` }], error: errMsg };
    }
};
exports.closeDocumentHandler = closeDocumentHandler;
/**
 * 导出所有文档管理相关的Tools
 */
exports.documentTools = [
    { definition: exports.getOpenDocumentsDefinition, handler: exports.getOpenDocumentsHandler },
    { definition: exports.switchDocumentDefinition, handler: exports.switchDocumentHandler },
    { definition: exports.openDocumentDefinition, handler: exports.openDocumentHandler },
    { definition: exports.createDocumentDefinition, handler: exports.createDocumentHandler },
    { definition: exports.closeDocumentDefinition, handler: exports.closeDocumentHandler },
    { definition: exports.getDocumentTextDefinition, handler: exports.getDocumentTextHandler },
    { definition: exports.insertHeaderDefinition, handler: exports.insertHeaderHandler },
    { definition: exports.insertFooterDefinition, handler: exports.insertFooterHandler },
    { definition: exports.generateDocTocDefinition, handler: exports.generateDocTocHandler },
    { definition: exports.insertSectionBreakDefinition, handler: exports.insertSectionBreakHandler },
    { definition: exports.setLineSpacingDefinition, handler: exports.setLineSpacingHandler },
];
exports.default = exports.documentTools;
//# sourceMappingURL=document.js.map
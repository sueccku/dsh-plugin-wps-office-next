"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.wordLongTailTools = exports.mailMergeHandler = exports.mailMergeDefinition = exports.insertCrossReferenceHandler = exports.insertCrossReferenceDefinition = exports.insertIndexHandler = exports.insertIndexDefinition = exports.getNotesHandler = exports.getNotesDefinition = exports.addEndnoteHandler = exports.addEndnoteDefinition = exports.addFootnoteHandler = exports.addFootnoteDefinition = exports.addContentControlHandler = exports.addContentControlDefinition = exports.getContentControlsHandler = exports.getContentControlsDefinition = void 0;
/**
 * Input: Word 长尾工具的调用参数
 * Output: 内容控件 / 脚注尾注 / 索引 / 交叉引用 / 邮件合并 的执行结果
 * Pos: Word 长尾（P3-4，D2 的最后一块）。支持面已用裸 COM 量过（test/.artifacts/e2e/word-probe*.ps1）。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P3 状态。
 */
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
const CONTROL_NAMES = { 0: '富文本', 1: '纯文本', 2: '复选框', 3: '组合框', 4: '下拉列表', 5: '日期选择器', 7: '图片' };
function longtailFail(prefix, error) {
    return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: prefix + ': ' + error }], error: error };
}
/** 内容控件（读） */
exports.getContentControlsDefinition = {
    name: 'wps_word_get_content_controls',
    description: '列出文档里的内容控件（可填写的结构化区域）：序号、类型、标题、标签、当前文本。使用场景：看清模板里有哪些可填字段，配合 wps_word_add_content_control 使用。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {} },
};
const getContentControlsHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getContentControls', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return longtailFail('读取内容控件失败', response.error);
        const controls = response.data?.controls || [];
        if (!controls.length)
            return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '这个文档里没有内容控件。' }] };
        const lines = controls.map((c) => {
            const kind = CONTROL_NAMES[Number(c.type)] || ('类型 ' + String(c.type));
            const label = c.title || c.tag || '';
            return '  ' + String(c.index) + '. ' + kind + (label ? '「' + label + '」' : '') + '：' + String(c.text || '');
        });
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '内容控件（' + controls.length + ' 个）:\n' + lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return longtailFail('读取内容控件出错', errMsg);
    }
};
exports.getContentControlsHandler = getContentControlsHandler;
/** 内容控件（加） */
exports.addContentControlDefinition = {
    name: 'wps_word_add_content_control',
    description: '在光标处插入一个内容控件（可填写的结构化区域），可指定类型、标题、标签与初始文本。使用场景：做一份可复用的表单模板。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            type: { type: 'string', enum: ['richText', 'plainText', 'checkBox', 'comboBox', 'dropDownList', 'datePicker', 'picture'], description: '控件类型，默认 richText（富文本）' },
            title: { type: 'string', description: '标题（界面上显示的说明）' },
            tag: { type: 'string', description: '标签（给程序看的键名）' },
            text: { type: 'string', description: '初始文本' },
        },
    },
};
const addContentControlHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addContentControl', { type: args.type, title: args.title, tag: args.tag, text: args.text }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return longtailFail('插入内容控件失败', response.error);
        const d = response.data || {};
        const kind = CONTROL_NAMES[Number(d.type)] || String(d.type);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已插入' + kind + '内容控件' + (d.title ? '（标题 ' + d.title + '）' : '') + '，当前内容「' + String(d.text || '') + '」；文档里共 ' + String(d.total ?? 0) + ' 个' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return longtailFail('插入内容控件出错', errMsg);
    }
};
exports.addContentControlHandler = addContentControlHandler;
/** 脚注 / 尾注（加） */
exports.addFootnoteDefinition = {
    name: 'wps_word_add_footnote',
    description: '在光标处插入脚注（页面底部的注释）。使用场景：给一个术语加解释。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: { text: { type: 'string', description: '脚注内容' } }, required: ['text'] },
};
const addFootnoteHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addFootnote', { text: args.text }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return longtailFail('插入脚注失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已插入脚注「' + String(response.data?.text || '') + '」；文档里共 ' + String(response.data?.total ?? 0) + ' 条' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return longtailFail('插入脚注出错', errMsg);
    }
};
exports.addFootnoteHandler = addFootnoteHandler;
exports.addEndnoteDefinition = {
    name: 'wps_word_add_endnote',
    description: '在光标处插入尾注（文档末尾的注释）。使用场景：给引用加出处，统一列在文末。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: { text: { type: 'string', description: '尾注内容' } }, required: ['text'] },
};
const addEndnoteHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addEndnote', { text: args.text }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return longtailFail('插入尾注失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已插入尾注「' + String(response.data?.text || '') + '」；文档里共 ' + String(response.data?.total ?? 0) + ' 条' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return longtailFail('插入尾注出错', errMsg);
    }
};
exports.addEndnoteHandler = addEndnoteHandler;
/** 脚注 / 尾注（读） */
exports.getNotesDefinition = {
    name: 'wps_word_get_notes',
    description: '列出文档里的脚注与尾注及其内容。使用场景：核对注释有没有漏、内容对不对。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {} },
};
const getNotesHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getNotes', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return longtailFail('读取注释失败', response.error);
        const d = response.data || {};
        const footnotes = d.footnotes || [];
        const endnotes = d.endnotes || [];
        const lines = ['脚注 ' + String(d.footnoteCount ?? footnotes.length) + ' 条，尾注 ' + String(d.endnoteCount ?? endnotes.length) + ' 条'];
        for (const f of footnotes)
            lines.push('  脚注 ' + String(f.index) + ': ' + String(f.text || ''));
        for (const e of endnotes)
            lines.push('  尾注 ' + String(e.index) + ': ' + String(e.text || ''));
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return longtailFail('读取注释出错', errMsg);
    }
};
exports.getNotesHandler = getNotesHandler;
/** 索引 */
exports.insertIndexDefinition = {
    name: 'wps_word_insert_index',
    description: '在文档末尾插入索引（按索引项自动生成）。使用场景：长文档要一张术语/条目索引。注意：只有先用 Word 的「标记索引项」标记过内容，索引里才会有条目；没标记过会显示「未找到索引项」。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {} },
};
const insertIndexHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertIndex', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return longtailFail('插入索引失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已插入索引（文档里共 ' + String(response.data?.indexes ?? 0) + ' 个）。' + String(response.data?.message || '') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return longtailFail('插入索引出错', errMsg);
    }
};
exports.insertIndexHandler = insertIndexHandler;
/** 交叉引用 */
exports.insertCrossReferenceDefinition = {
    name: 'wps_word_insert_cross_reference',
    description: '在光标处插入交叉引用（引用标题/书签/脚注等）。参数直接对应 Word 的 InsertCrossReference(ReferenceType, ReferenceKind, ReferenceItem)：referenceType 1=标题 2=书签 3=脚注 4=尾注，referenceKind -1=正文文本，referenceItem 是要引用的编号或名字。这两个枚举在 WPS 上没有被验证过的友好映射，所以照实透传，不做猜测性翻译。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            referenceType: { type: 'number', description: '引用类型（Word 的 WdReferenceType）：1=标题 2=书签 3=脚注 4=尾注；默认 1' },
            referenceKind: { type: 'number', description: '引用内容（Word 的 WdReferenceKind）：-1=正文文本；默认 -1' },
            referenceItem: { type: 'string', description: '要引用的项目：编号（例如 1）或名字（例如书签名）' },
        },
        required: ['referenceItem'],
    },
};
const insertCrossReferenceHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertCrossReference', { referenceType: args.referenceType, referenceKind: args.referenceKind, referenceItem: args.referenceItem }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return longtailFail('插入交叉引用失败', response.error);
        const d = response.data || {};
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已插入交叉引用（类型 ' + String(d.referenceType) + '，内容 ' + String(d.referenceKind) + '，引用 ' + String(d.referenceItem) + '）' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return longtailFail('插入交叉引用出错', errMsg);
    }
};
exports.insertCrossReferenceHandler = insertCrossReferenceHandler;
/** 邮件合并 */
exports.mailMergeDefinition = {
    name: 'wps_word_mail_merge',
    description: '邮件合并：以当前文档为母版，接一个数据文件（CSV），把指定字段插到光标处，再按每一行数据生成一个新文档（母版不动）。使用场景：一份通知模板 + 一张名单，每人一份。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            dataFile: { type: 'string', description: '数据文件路径（CSV，第一行是字段名）' },
            fields: { type: 'array', description: '要插入的字段名列表，按顺序插到光标处，例如 Name、City；不填则只打开数据源不插字段' },
        },
        required: ['dataFile'],
    },
};
const mailMergeHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('mailMerge', { dataFile: args.dataFile, fields: args.fields }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return longtailFail('邮件合并失败', response.error);
        const d = response.data || {};
        const lines = [d.message || '已合并'];
        lines.push('  新文档: ' + String(d.mergedDocument || '?') + '（母版 ' + String(d.sourceDocument || '?') + ' 未改动）');
        const fields = d.fields || [];
        if (fields.length)
            lines.push('  字段: ' + fields.join(', '));
        if (d.preview)
            lines.push('  开头: ' + String(d.preview).replace(/[\r\n]+/g, ' / '));
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return longtailFail('邮件合并出错', errMsg);
    }
};
exports.mailMergeHandler = mailMergeHandler;
exports.wordLongTailTools = [
    { definition: exports.getContentControlsDefinition, handler: exports.getContentControlsHandler },
    { definition: exports.addContentControlDefinition, handler: exports.addContentControlHandler },
    { definition: exports.addFootnoteDefinition, handler: exports.addFootnoteHandler },
    { definition: exports.addEndnoteDefinition, handler: exports.addEndnoteHandler },
    { definition: exports.getNotesDefinition, handler: exports.getNotesHandler },
    { definition: exports.insertIndexDefinition, handler: exports.insertIndexHandler },
    { definition: exports.insertCrossReferenceDefinition, handler: exports.insertCrossReferenceHandler },
    { definition: exports.mailMergeDefinition, handler: exports.mailMergeHandler },
];
exports.default = exports.wordLongTailTools;
//# sourceMappingURL=longtail.js.map
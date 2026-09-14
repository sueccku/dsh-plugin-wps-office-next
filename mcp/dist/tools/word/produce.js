"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.wordProduceTools = exports.deleteCommentHandler = exports.deleteCommentDefinition = exports.rejectRevisionsHandler = exports.rejectRevisionsDefinition = exports.acceptRevisionsHandler = exports.acceptRevisionsDefinition = exports.getRevisionsHandler = exports.getRevisionsDefinition = exports.setColumnsHandler = exports.setColumnsDefinition = exports.insertPageNumbersHandler = exports.insertPageNumbersDefinition = void 0;
/**
 * Input: Word 文档生产族工具的调用参数
 * Output: 页码 / 分栏 / 修订 / 批注删除 的执行结果
 * Pos: Word 文档生产族（P3-3）。支持面已用裸 COM 量过（见 test/.artifacts/e2e/word-probe*.ps1）。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P3 状态。
 */
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
const sectionParam = { type: 'number', description: '第几节（从 1 开始）；不填则用第 1 节' };
/** 修订类型（Word 的 WdRevisionType），只翻译常见的几种，其余如实给数字 */
const REVISION_NAMES = {
    1: '插入', 2: '删除', 3: '属性', 4: '段落编号', 5: '域显示', 7: '冲突', 8: '样式',
    9: '替换', 10: '段落属性', 11: '表格属性', 12: '节属性', 16: '单元格插入', 17: '单元格删除', 18: '单元格合并',
};
function produceFail(prefix, error) {
    return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: prefix + ': ' + error }], error: error };
}
/** 页码 */
exports.insertPageNumbersDefinition = {
    name: 'wps_word_insert_page_numbers',
    description: '给某一节的页眉或页脚插入页码。使用场景：文档要打印，页脚右下角要有页码。页码是域，页数变化会自动更新。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            section: sectionParam,
            position: { type: 'string', enum: ['footer', 'header'], description: '放页脚还是页眉，默认 footer' },
            alignment: { type: 'string', enum: ['left', 'center', 'right'], description: '对齐方式，默认 right' },
            showFirstPage: { type: 'boolean', description: '是否在首页显示，默认 true' },
        },
    },
};
const insertPageNumbersHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertPageNumbers', { section: args.section, position: args.position, alignment: args.alignment, showFirstPage: args.showFirstPage }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return produceFail('插入页码失败', response.error);
        const d = response.data || {};
        const where = d.position === 'header' ? '页眉' : '页脚';
        const align = d.alignment === 'center' ? '居中' : d.alignment === 'left' ? '左对齐' : '右对齐';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '第 ' + String(d.section) + ' 节' + where + '已插入页码（' + align + '，该处共 ' + String(d.pageNumbers ?? 0) + ' 个页码域）' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return produceFail('插入页码出错', errMsg);
    }
};
exports.insertPageNumbersHandler = insertPageNumbersHandler;
/** 分栏 */
exports.setColumnsDefinition = {
    name: 'wps_word_set_columns',
    description: '设置分栏：栏数、栏间距、是否加分隔线。使用场景：把长文排成两栏。count 用 1 就是取消分栏。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            section: sectionParam,
            count: { type: 'number', description: '栏数（1-12），默认 1（取消分栏）' },
            spacing: { type: 'number', description: '栏间距（磅）' },
            lineBetween: { type: 'boolean', description: '是否在栏间加分隔线' },
        },
    },
};
const setColumnsHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setColumns', { section: args.section, count: args.count, spacing: args.spacing, lineBetween: args.lineBetween }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return produceFail('设置分栏失败', response.error);
        const d = response.data || {};
        const applied = d.applied || [];
        // 单栏时 WPS 的 TextColumns.Spacing 是哨兵值 9999999，印出来像 bug，所以只在分栏时报告间距。
        const spacing = Number(d.count) > 1 ? '，栏间距 ' + String(d.spacing) + ' 磅' : '';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '第 ' + String(d.section) + ' 节现在是 ' + String(d.count) + ' 栏（已应用 ' + applied.join(', ') + spacing + '）' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return produceFail('设置分栏出错', errMsg);
    }
};
exports.setColumnsHandler = setColumnsHandler;
/** 修订列表 */
exports.getRevisionsDefinition = {
    name: 'wps_word_get_revisions',
    description: '列出文档里的修订（插入/删除/替换等），并报告「修订跟踪」当前是否打开。使用场景：接手别人的稿子，先看改了什么。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {} },
};
const getRevisionsHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getRevisions', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return produceFail('读取修订失败', response.error);
        const d = response.data || {};
        const revisions = d.revisions || [];
        const head = '修订跟踪: ' + (d.trackChanges ? '已打开' : '已关闭');
        if (!revisions.length)
            return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: head + '；文档里没有修订。' }] };
        const lines = revisions.map((r) => '  ' + String(r.index) + '. ' + (REVISION_NAMES[Number(r.type)] || ('类型 ' + String(r.type))) + '：' + String(r.text || '') + (r.author ? '（' + r.author + '）' : ''));
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: head + '；修订（' + revisions.length + ' 处）:\n' + lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return produceFail('读取修订出错', errMsg);
    }
};
exports.getRevisionsHandler = getRevisionsHandler;
/** 接受修订 */
exports.acceptRevisionsDefinition = {
    name: 'wps_word_accept_revisions',
    description: '接受修订：给 index 只接受那一处，不填则接受全部。使用场景：审阅通过，把改动定稿。接受后文字变成正文，修订记录消失。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: { index: { type: 'number', description: '只接受第几处修订（从 1 开始）；不填则全部接受' } } },
};
const acceptRevisionsHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('acceptRevisions', { index: args.index }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return produceFail('接受修订失败', response.error);
        const d = response.data || {};
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已接受 ' + String(d.accepted ?? 0) + ' 处修订；还剩 ' + String(d.remaining ?? 0) + ' 处' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return produceFail('接受修订出错', errMsg);
    }
};
exports.acceptRevisionsHandler = acceptRevisionsHandler;
/** 拒绝修订 */
exports.rejectRevisionsDefinition = {
    name: 'wps_word_reject_revisions',
    description: '拒绝修订：给 index 只拒绝那一处，不填则拒绝全部（回到改之前的原文）。使用场景：这版改动不要，恢复原样。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: { index: { type: 'number', description: '只拒绝第几处修订（从 1 开始）；不填则全部拒绝' } } },
};
const rejectRevisionsHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('rejectRevisions', { index: args.index }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return produceFail('拒绝修订失败', response.error);
        const d = response.data || {};
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已拒绝 ' + String(d.rejected ?? 0) + ' 处修订；还剩 ' + String(d.remaining ?? 0) + ' 处' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return produceFail('拒绝修订出错', errMsg);
    }
};
exports.rejectRevisionsHandler = rejectRevisionsHandler;
/** 删除批注 */
exports.deleteCommentDefinition = {
    name: 'wps_word_delete_comment',
    description: '删除批注：给 index 只删那一条（序号见 wps_word_get_comments），不填则全部删除。使用场景：意见处理完了，清掉批注再交付。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: { index: { type: 'number', description: '只删第几条批注（从 1 开始）；不填则全部删除' } } },
};
const deleteCommentHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteComment', { index: args.index }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return produceFail('删除批注失败', response.error);
        const d = response.data || {};
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已删除 ' + String(d.deleted ?? 0) + ' 条批注；还剩 ' + String(d.remaining ?? 0) + ' 条' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return produceFail('删除批注出错', errMsg);
    }
};
exports.deleteCommentHandler = deleteCommentHandler;
exports.wordProduceTools = [
    { definition: exports.insertPageNumbersDefinition, handler: exports.insertPageNumbersHandler },
    { definition: exports.setColumnsDefinition, handler: exports.setColumnsHandler },
    { definition: exports.getRevisionsDefinition, handler: exports.getRevisionsHandler },
    { definition: exports.acceptRevisionsDefinition, handler: exports.acceptRevisionsHandler },
    { definition: exports.rejectRevisionsDefinition, handler: exports.rejectRevisionsHandler },
    { definition: exports.deleteCommentDefinition, handler: exports.deleteCommentHandler },
];
exports.default = exports.wordProduceTools;
//# sourceMappingURL=produce.js.map
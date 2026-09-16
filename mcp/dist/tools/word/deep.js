"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.wordDeepTools = exports.convertTableToTextHandler = exports.convertTableToTextDefinition = exports.setTableFormatHandler = exports.setTableFormatDefinition = exports.splitTableCellHandler = exports.splitTableCellDefinition = exports.mergeTableCellsHandler = exports.mergeTableCellsDefinition = exports.deleteTableLineHandler = exports.deleteTableLineDefinition = exports.addTableLinesHandler = exports.addTableLinesDefinition = exports.setTableCellHandler = exports.setTableCellDefinition = exports.getTableDataHandler = exports.getTableDataDefinition = exports.getTablesHandler = exports.getTablesDefinition = exports.insertHyperlinkHandler = exports.insertHyperlinkDefinition = exports.getDocumentStatsHandler = exports.getDocumentStatsDefinition = exports.getCommentsHandler = exports.getCommentsDefinition = exports.getBookmarksHandler = exports.getBookmarksDefinition = void 0;
/**
 * Input: Word 深水区工具的调用参数
 * Output: 书签/批注/文档统计/超链接 与 表格读写编辑的执行结果
 * Pos: Word 深水区工具（P3）。P3-1 把桥里已有却没有出口的能力挂出来；P3-2 是表格族的新 COM 代码。
 *      实测缺口：水印（表头 Shapes 全不支持）与文档属性（集合是坏壳）在 WPS 里做不到，见 docs/FIXES.md。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P3 状态。
 */
const uuid_1 = require("uuid");
const impact_1 = require("../impact");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
const tableParam = { type: 'number', description: '第几张表（从 1 开始）；不填则用第 1 张' };
function deepFail(prefix, error) {
    return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: prefix + ': ' + error }], error: error };
}
/** 书签 */
exports.getBookmarksDefinition = {
    name: 'wps_word_get_bookmarks',
    description: '列出文档里的全部书签（名字与位置）。使用场景：填模板之前先看清有哪些占位书签。配合 wps_word_replace_bookmark_content 使用。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {} },
};
const getBookmarksHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getBookmarks', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('读取书签失败', response.error);
        const bookmarks = response.data?.bookmarks || [];
        if (!bookmarks.length)
            return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '这个文档里没有书签。' }] };
        const lines = bookmarks.map((b) => '  ' + (b.name || '?') + '（位置 ' + String(b.start) + '-' + String(b.end) + '）');
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '书签（' + bookmarks.length + ' 个）:\n' + lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('读取书签出错', errMsg);
    }
};
exports.getBookmarksHandler = getBookmarksHandler;
/** 批注（读） */
exports.getCommentsDefinition = {
    name: 'wps_word_get_comments',
    description: '列出文档里的全部批注：序号、正文、作者、时间。使用场景：汇总一批审阅意见。加批注用 wps_word_insert_comment。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {} },
};
const getCommentsHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getComments', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('读取批注失败', response.error);
        const comments = response.data?.comments || [];
        if (!comments.length)
            return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '这个文档里没有批注。' }] };
        const lines = comments.map((c) => {
            const text = String(c.text || '').replace(/[\r\n\x07]+/g, ' ').trim();
            const who = c.author ? '（' + c.author + (c.date ? ' ' + c.date : '') + '）' : '';
            return '  ' + String(c.index) + '. ' + text + who;
        });
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '批注（' + comments.length + ' 条）:\n' + lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('读取批注出错', errMsg);
    }
};
exports.getCommentsHandler = getCommentsHandler;
/** 文档统计 */
exports.getDocumentStatsDefinition = {
    name: 'wps_word_get_document_stats',
    description: '文档统计：页数、字数、字符数、段落数、行数。使用场景：「这份文档多少字」「排出来几页」。页数按当前排版计算。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {} },
};
const getDocumentStatsHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getDocumentStats', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('读取文档统计失败', response.error);
        const d = response.data || {};
        const lines = ['文档 ' + (d.name || '?') + ':'];
        lines.push('  页数 ' + String(d.pages) + '；字数 ' + String(d.words) + '；字符数 ' + String(d.characters));
        lines.push('  段落 ' + String(d.paragraphs) + '；行数 ' + String(d.lines));
        if (d.path)
            lines.push('  路径: ' + d.path);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('读取文档统计出错', errMsg);
    }
};
exports.getDocumentStatsHandler = getDocumentStatsHandler;
/** 超链接 */
exports.insertHyperlinkDefinition = {
    name: 'wps_word_insert_hyperlink',
    description: '在光标处插入超链接。如果当前选中了文字，就把它变成链接（用选中的文字当显示文本）；否则先插入 text 再把它变成链接。使用场景：给「详见官网」加上链接。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            url: { type: 'string', description: '链接地址，如 https://example.com' },
            text: { type: 'string', description: '显示文本；不填则直接用地址。选中了文字时忽略' },
        },
        required: ['url'],
    },
};
const insertHyperlinkHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertHyperlink', { url: args.url, text: args.text }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('插入超链接失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '已插入超链接：' + (response.data?.text || '') + ' → ' + (response.data?.url || String(args.url)) }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('插入超链接出错', errMsg);
    }
};
exports.insertHyperlinkHandler = insertHyperlinkHandler;
/** 表格：列表 */
exports.getTablesDefinition = {
    name: 'wps_word_get_tables',
    description: '列出文档里的全部表格：序号、行列数、样式、文本预览。使用场景：先看清文档里有几张表、哪张是要改的，再用其他表格工具按序号操作。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {} },
};
const getTablesHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getDocumentTables', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('读取表格失败', response.error);
        const tables = response.data?.tables || [];
        if (!tables.length)
            return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '这个文档里没有表格。' }] };
        const lines = tables.map((t) => '  ' + String(t.index) + '. ' + String(t.rows) + ' 行 x ' + String(t.columns) + ' 列' + (t.style ? '，样式 ' + t.style : '') + (t.text ? '\n     ' + t.text.replace(/[\r\n\x07]+/g, ' / ') : ''));
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '表格（' + tables.length + ' 张）:\n' + lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('读取表格出错', errMsg);
    }
};
exports.getTablesHandler = getTablesHandler;
/** 表格：读数据 */
exports.getTableDataDefinition = {
    name: 'wps_word_get_table_data',
    description: '按行列读出某张表的全部单元格文本（合并单元格的非起点格子会是空字符串）。使用场景：把 Word 表格里的数据取出来核对。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: { table: tableParam } },
};
const getTableDataHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getTableData', { table: args.table }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('读取表格数据失败', response.error);
        const d = response.data || {};
        const rows = d.data || [];
        const lines = rows.map((row, i) => '  第' + (i + 1) + '行: ' + row.join(' | '));
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '表 ' + String(d.table) + '（' + String(d.rows) + ' 行 x ' + String(d.columns) + ' 列）:\n' + lines.join('\n') }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('读取表格数据出错', errMsg);
    }
};
exports.getTableDataHandler = getTableDataHandler;
/** 表格：写单元格 */
exports.setTableCellDefinition = {
    name: 'wps_word_set_table_cell',
    description: '写入表格的某个单元格（行列都从 1 开始）。使用场景：填表格、改一处数据。合并单元格只有起点能写。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            table: tableParam,
            row: { type: 'number', description: '第几行（从 1 开始）' },
            column: { type: 'number', description: '第几列（从 1 开始）' },
            text: { type: 'string', description: '要写入的文本' },
        },
        required: ['row', 'column', 'text'],
    },
};
const setTableCellHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setTableCell', { table: args.table, row: args.row, column: args.column, text: args.text }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('写入单元格失败', response.error);
        const d = response.data || {};
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '表 ' + String(d.table) + ' 第 ' + String(d.row) + ' 行第 ' + String(d.column) + ' 列已写入: ' + String(d.text) }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('写入单元格出错', errMsg);
    }
};
exports.setTableCellHandler = setTableCellHandler;
/** 表格：增行/增列 */
exports.addTableLinesDefinition = {
    name: 'wps_word_add_table_lines',
    description: '给表格增加行或列。kind 选 row/column，count 默认 1；给 position 就插在那一行/列之前，否则追加到末尾。使用场景：表格要再加几行。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            table: tableParam,
            kind: { type: 'string', enum: ['row', 'column'], description: '加行还是加列，默认 row' },
            count: { type: 'number', description: '加几行/几列，默认 1' },
            position: { type: 'number', description: '插在第几行/列之前；不填则追加到末尾' },
        },
        required: ['kind'],
    },
};
const addTableLinesHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addTableLines', { table: args.table, kind: args.kind, count: args.count, position: args.position }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('增加行列失败', response.error);
        const d = response.data || {};
        const what = d.kind === 'column' ? '列' : '行';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '表 ' + String(d.table) + ' 已增加 ' + String(d.added) + ' ' + what + '，现在 ' + String(d.rows) + ' 行 x ' + String(d.columns) + ' 列' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('增加行列出错', errMsg);
    }
};
exports.addTableLinesHandler = addTableLinesHandler;
/** 表格：删行/删列 */
exports.deleteTableLineDefinition = {
    name: 'wps_word_delete_table_line',
    description: '删除表格的第几行或第几列（从 1 开始）。使用场景：删掉一行多余记录。删整行用 kind=row。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            table: tableParam,
            kind: { type: 'string', enum: ['row', 'column'], description: '删行还是删列，默认 row' },
            lineIndex: { type: 'number', description: '第几行/列（从 1 开始）' },
        },
        required: ['kind', 'lineIndex'],
    },
};
const deleteTableLineHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteTableLine', { table: args.table, kind: args.kind, lineIndex: args.lineIndex }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('删除行列失败', response.error);
        const d = response.data || {};
        const what = d.kind === 'column' ? '列' : '行';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '表 ' + String(d.table) + ' 已删除 1 ' + what + '，现在 ' + String(d.rows) + ' 行 x ' + String(d.columns) + ' 列' + (0, impact_1.impactText)(d.impact) }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('删除行列出错', errMsg);
    }
};
exports.deleteTableLineHandler = deleteTableLineHandler;
/** 表格：合并单元格 */
exports.mergeTableCellsDefinition = {
    name: 'wps_word_merge_table_cells',
    description: '把一块矩形区域合并成一个单元格：从（startRow, startColumn）到（endRow, endColumn）。使用场景：表头跨列、跨行。行列都从 1 开始。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            table: tableParam,
            startRow: { type: 'number', description: '起始行（从 1 开始）' },
            startColumn: { type: 'number', description: '起始列（从 1 开始）' },
            endRow: { type: 'number', description: '结束行（从 1 开始）' },
            endColumn: { type: 'number', description: '结束列（从 1 开始）' },
        },
        required: ['startRow', 'startColumn', 'endRow', 'endColumn'],
    },
};
const mergeTableCellsHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('mergeTableCells', { table: args.table, startRow: args.startRow, startColumn: args.startColumn, endRow: args.endRow, endColumn: args.endColumn }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('合并单元格失败', response.error);
        const d = response.data || {};
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: (d.message || '已合并单元格') + '；表 ' + String(d.table) + ' 现在 ' + String(d.rows) + ' 行 x ' + String(d.columns) + ' 列' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('合并单元格出错', errMsg);
    }
};
exports.mergeTableCellsHandler = mergeTableCellsHandler;
/** 表格：拆分单元格 */
exports.splitTableCellDefinition = {
    name: 'wps_word_split_table_cell',
    description: '把一个单元格拆成 rows x columns 个小格（默认 1 x 2）。使用场景：合并错了要还原、或者一格里本来就该分两列。行列都从 1 开始。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            table: tableParam,
            row: { type: 'number', description: '第几行（从 1 开始）' },
            column: { type: 'number', description: '第几列（从 1 开始）' },
            rows: { type: 'number', description: '拆成几行，默认 1' },
            columns: { type: 'number', description: '拆成几列，默认 2' },
        },
        required: ['row', 'column'],
    },
};
const splitTableCellHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('splitTableCell', { table: args.table, row: args.row, column: args.column, rows: args.rows, columns: args.columns }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('拆分单元格失败', response.error);
        const d = response.data || {};
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: (d.message || '已拆分单元格') + '；表 ' + String(d.table) + ' 现在 ' + String(d.rows) + ' 行 x ' + String(d.columns) + ' 列' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('拆分单元格出错', errMsg);
    }
};
exports.splitTableCellHandler = splitTableCellHandler;
/** 表格：样式 */
exports.setTableFormatDefinition = {
    name: 'wps_word_set_table_format',
    description: '设置表格外观：表格样式名、是否显示边框、按内容或按窗口自动调整宽度、给表头行加底纹色。使用场景：把裸表格弄成带框线、表头有底色、宽度合适的正式表格。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            table: tableParam,
            style: { type: 'string', description: '表格样式名，如「网格型」「普通表格」' },
            borders: { type: 'boolean', description: '是否显示全部边框' },
            autoFit: { type: 'string', enum: ['content', 'window'], description: '按内容自适应或按页面宽度自适应' },
            headerShading: { type: 'string', description: '表头行底纹色，十六进制如 #D9E2F3' },
        },
    },
};
const setTableFormatHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setTableFormat', { table: args.table, style: args.style, borders: args.borders, autoFit: args.autoFit, headerShading: args.headerShading }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('设置表格外观失败', response.error);
        const d = response.data || {};
        const applied = d.applied || [];
        const head = applied.length ? '表 ' + String(d.table) + ' 已更新（' + applied.join(', ') + '）' : '表 ' + String(d.table) + ' 没有变化';
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: head + '；' + String(d.rows) + ' 行 x ' + String(d.columns) + ' 列' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('设置表格外观出错', errMsg);
    }
};
exports.setTableFormatHandler = setTableFormatHandler;
/** 表格 → 文本 */
exports.convertTableToTextDefinition = {
    name: 'wps_word_convert_table_to_text',
    description: '把表格转成普通文本，列之间用一个分隔符（tab/comma/paragraph，或直接给一个字符）。使用场景：要按段落正文交付，不想留表格对象。转换后表格对象消失，文字保留。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            table: tableParam,
            separator: { type: 'string', description: '列分隔符：tab / comma / paragraph，或直接给一个字符；不填用 Word 默认' },
        },
    },
};
const convertTableToTextHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('convertTableToText', { table: args.table, separator: args.separator }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return deepFail('表格转文本失败', response.error);
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: '表格已转成文本；文档里还剩 ' + String(response.data?.tablesRemaining ?? 0) + ' 张表' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return deepFail('表格转文本出错', errMsg);
    }
};
exports.convertTableToTextHandler = convertTableToTextHandler;
exports.wordDeepTools = [
    { definition: exports.getBookmarksDefinition, handler: exports.getBookmarksHandler },
    { definition: exports.getCommentsDefinition, handler: exports.getCommentsHandler },
    { definition: exports.getDocumentStatsDefinition, handler: exports.getDocumentStatsHandler },
    { definition: exports.insertHyperlinkDefinition, handler: exports.insertHyperlinkHandler },
    { definition: exports.getTablesDefinition, handler: exports.getTablesHandler },
    { definition: exports.getTableDataDefinition, handler: exports.getTableDataHandler },
    { definition: exports.setTableCellDefinition, handler: exports.setTableCellHandler },
    { definition: exports.addTableLinesDefinition, handler: exports.addTableLinesHandler },
    { definition: exports.deleteTableLineDefinition, handler: exports.deleteTableLineHandler },
    { definition: exports.mergeTableCellsDefinition, handler: exports.mergeTableCellsHandler },
    { definition: exports.splitTableCellDefinition, handler: exports.splitTableCellHandler },
    { definition: exports.setTableFormatDefinition, handler: exports.setTableFormatHandler },
    { definition: exports.convertTableToTextDefinition, handler: exports.convertTableToTextHandler },
];
exports.default = exports.wordDeepTools;
//# sourceMappingURL=deep.js.map
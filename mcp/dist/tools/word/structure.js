"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.wordStructureTools = exports.createEnvelopeHandler = exports.createEnvelopeDefinition = exports.createLabelSheetHandler = exports.createLabelSheetDefinition = exports.resolveRevisionsHandler = exports.resolveRevisionsDefinition = exports.setContentControlHandler = exports.setContentControlDefinition = exports.updateStyleHandler = exports.updateStyleDefinition = exports.compareDocumentsHandler = exports.compareDocumentsDefinition = exports.insertTableOfFiguresHandler = exports.insertTableOfFiguresDefinition = exports.setTableFormulaHandler = exports.setTableFormulaDefinition = exports.insertCaptionHandler = exports.insertCaptionDefinition = exports.getProtectionHandler = exports.getProtectionDefinition = exports.protectDocumentHandler = exports.protectDocumentDefinition = exports.selectRangeHandler = exports.selectRangeDefinition = exports.updateFieldsHandler = exports.updateFieldsDefinition = exports.getFieldsHandler = exports.getFieldsDefinition = void 0;
/**
 * Input: Word 域 / 选区 / 保护的参数
 * Output: 这些能力的工具结果
 * Pos: P1（2026-10-19）新增的 Word 结构能力。工具默认隐藏，不占广告面。
 *      每个 handler 都只做 dispatch；COM 细节在 mcp/scripts/wps-com.ps1 的对应 action。
 */
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
exports.getFieldsDefinition = {
    name: 'wps_word_get_fields',
    description: '列出 Word 文档里的域（PAGE/NUMPAGES/SEQ/TOC/REF/INCLUDEPICTURE 等）：序号、类型、域代码、当前结果。用于题注/交叉引用/目录/页码的体检。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {}, required: [] },
};
const getFieldsHandler = async () => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getDocumentFields', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `读取域失败: ${response.error}` }], error: response.error };
        const d = response.data || { count: 0, fields: [] };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `域数量: ${d.count}\n${JSON.stringify(d.fields, null, 2)}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `读取域出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getFieldsHandler = getFieldsHandler;
exports.updateFieldsDefinition = {
    name: 'wps_word_update_fields',
    description: '更新 Word 的域（相当于 F9）：不传 index 更新全部，传 index 只更新那一个。题注/目录/交叉引用/页码改完后用它刷新。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: { index: { type: 'number', description: '只更新第几个域（从1开始）；不填更新全部' } },
        required: [],
    },
};
const updateFieldsHandler = async (args) => {
    const { index } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('updateDocumentFields', index === undefined ? {} : { index }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `更新域失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `域已更新: ${response.data?.updated ?? 0} 个` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `更新域出错: ${errMsg}` }], error: errMsg };
    }
};
exports.updateFieldsHandler = updateFieldsHandler;
exports.selectRangeDefinition = {
    name: 'wps_word_select_range',
    description: '把 Word 选区定位到 0 基字符范围 [start,end)。后续 insert_*（图片/内容控件/交叉引用等）会插到这里，不再只能依赖当前用户选区。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            start: { type: 'number', description: '起始字符位置（0 基）' },
            end: { type: 'number', description: '结束字符位置（不含）' },
        },
        required: ['start', 'end'],
    },
};
const selectRangeHandler = async (args) => {
    const { start, end } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('selectDocumentRange', { start, end }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `定位选区失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `选区已定位: ${response.data?.start} – ${response.data?.end}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `定位选区出错: ${errMsg}` }], error: errMsg };
    }
};
exports.selectRangeHandler = selectRangeHandler;
exports.protectDocumentDefinition = {
    name: 'wps_word_protect',
    description: '保护/取消保护 Word 文档：type=readOnly/comments/revisions/forms，可带 password。protect=false 取消保护。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            protect: { type: 'boolean', description: 'true 保护（默认），false 取消保护' },
            type: { type: 'string', enum: ['readOnly', 'comments', 'revisions', 'forms'], description: '保护类型，默认 readOnly' },
            password: { type: 'string', description: '保护密码（可选）' },
        },
        required: [],
    },
};
const protectDocumentHandler = async (args) => {
    const { protect, type, password } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('protectDocument', { protect: protect !== false, type, password }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `文档保护失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: response.data?.protect ? `文档已保护（${response.data?.type}）` : '文档已取消保护' }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `文档保护出错: ${errMsg}` }], error: errMsg };
    }
};
exports.protectDocumentHandler = protectDocumentHandler;
exports.getProtectionDefinition = {
    name: 'wps_word_get_protection',
    description: '读取 Word 文档当前保护状态（none/revisions/comments/forms/readOnly）。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: { type: 'object', properties: {}, required: [] },
};
const getProtectionHandler = async () => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getDocumentProtection', {}, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `读取保护状态失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `保护状态: ${response.data?.protection ?? 'unknown'} (${response.data?.protectionType})` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `读取保护状态出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getProtectionHandler = getProtectionHandler;
exports.insertCaptionDefinition = {
    name: 'wps_word_insert_caption',
    description: '在光标/选区处插入题注（图1、表1；可带章节号）。题注是 SEQ 域，插入后用 wps_word_update_fields 刷新编号。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            label: { type: 'string', description: '题注标签，如 图、表；默认 图' },
            title: { type: 'string', description: '题注文字，如 "系统架构"（可选）' },
            position: { type: 'string', enum: ['above', 'below'], description: '题注在对象上方/下方，默认 below' },
        },
        required: [],
    },
};
const insertCaptionHandler = async (args) => {
    const { label, title, position } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertCaption', { label, title, position }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入题注失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `题注已插入: ${response.data?.label} ${response.data?.title ?? ''}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入题注出错: ${errMsg}` }], error: errMsg };
    }
};
exports.insertCaptionHandler = insertCaptionHandler;
exports.setTableFormulaDefinition = {
    name: 'wps_word_set_table_formula',
    description: '给 Word 表格单元格设置公式（如 =SUM(ABOVE)、=SUM(LEFT)），写入 Cell.Formula；结果可用 update_fields 刷新。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            table: { type: 'number', description: '第几张表（从1开始），默认1' },
            row: { type: 'number', description: '行号（从1开始）' },
            column: { type: 'number', description: '列号（从1开始）' },
            formula: { type: 'string', description: '公式，如 =SUM(ABOVE)' },
        },
        required: ['row', 'column', 'formula'],
    },
};
const setTableFormulaHandler = async (args) => {
    const { table, row, column, formula } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setTableFormula', { table, row, column, formula }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置表格公式失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `表格公式已设置: ${response.data?.formula}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置表格公式出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setTableFormulaHandler = setTableFormulaHandler;
exports.insertTableOfFiguresDefinition = {
    name: 'wps_word_insert_table_of_figures',
    description: '在光标/选区处插入图表目录（TablesOfFigures，按题注标签汇总图/表并带页码）。插入后用 wps_word_update_fields 刷新。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: { label: { type: 'string', description: '题注标签，如 图、表；默认 图' } },
        required: [],
    },
};
const insertTableOfFiguresHandler = async (args) => {
    const { label } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertTableOfFigures', { label }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入图表目录失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `图表目录已插入（标签 ${response.data?.label}，共 ${response.data?.count} 个目录对象）` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `插入图表目录出错: ${errMsg}` }], error: errMsg };
    }
};
exports.insertTableOfFiguresHandler = insertTableOfFiguresHandler;
exports.compareDocumentsDefinition = {
    name: 'wps_word_compare_documents',
    description: '比较两份已打开的 Word 文档（原始稿 + 修订稿），生成带修订标记的比较结果文档。两份文档都必须已打开。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            original: { type: 'string', description: '原始稿文档名（已打开）' },
            revised: { type: 'string', description: '修订稿文档名（已打开）' },
        },
        required: ['original', 'revised'],
    },
};
const compareDocumentsHandler = async (args) => {
    const { original, revised } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('compareDocuments', { original, revised }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `文档比较失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `文档比较完成: ${response.data?.name}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `文档比较出错: ${errMsg}` }], error: errMsg };
    }
};
exports.compareDocumentsHandler = compareDocumentsHandler;
exports.updateStyleDefinition = {
    name: 'wps_word_update_style',
    description: '修改 Word 样式定义（字体名/字号/加粗/颜色、行距、首行缩进、段前段后），一次改完所有套用该样式的段落；用于"正文字体/行距统一"。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            styleName: { type: 'string', description: '样式名，如 正文、标题 1' },
            fontName: { type: 'string', description: '字体名' },
            fontSize: { type: 'number', description: '字号（磅）' },
            bold: { type: 'boolean', description: '是否加粗' },
            color: { type: 'number', description: '字体颜色（BGR 整数；可选）' },
            lineSpacing: { type: 'number', description: '行距（磅）' },
            firstLineIndent: { type: 'number', description: '首行缩进（磅）' },
            spaceBefore: { type: 'number', description: '段前（磅）' },
            spaceAfter: { type: 'number', description: '段后（磅）' },
        },
        required: ['styleName'],
    },
};
const updateStyleHandler = async (args) => {
    const { styleName, ...rest } = args;
    if (!styleName)
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: 'styleName 不能为空' }], error: '缺少 styleName' };
    try {
        const response = await wps_client_1.wpsClient.executeMethod('updateStyle', { styleName, ...rest }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `修改样式失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `样式已更新: ${response.data?.styleName}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `修改样式出错: ${errMsg}` }], error: errMsg };
    }
};
exports.updateStyleHandler = updateStyleHandler;
exports.setContentControlDefinition = {
    name: 'wps_word_set_content_control',
    description: '按序号/标题/标签填写内容控件：文本/下拉/日期用 value，复选框用 checked。用于模板表单批量填值。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            index: { type: 'number', description: '第几个内容控件（从1开始）' },
            title: { type: 'string', description: '按标题查找控件' },
            tag: { type: 'string', description: '按标签查找控件' },
            value: { type: 'string', description: '要填入的文本/下拉项/日期' },
            checked: { type: 'boolean', description: '复选框是否勾选' },
        },
        required: [],
    },
};
const setContentControlHandler = async (args) => {
    const { index, title, tag, value, checked } = args;
    if (index === undefined && !title && !tag) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: '需要 index、title 或 tag 之一来定位内容控件' }], error: '缺少控件定位' };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setContentControl', { index, title, tag, value, checked }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `填写内容控件失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `内容控件已更新${response.data?.title ? '（' + response.data.title + '）' : ''}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `填写内容控件出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setContentControlHandler = setContentControlHandler;
exports.resolveRevisionsDefinition = {
    name: 'wps_word_resolve_revisions',
    description: '按作者批量接受/拒绝修订：mode=accept/reject，author 不填则全部。用于"只处理张三的修订"这类审阅。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            mode: { type: 'string', enum: ['accept', 'reject'], description: '接受或拒绝' },
            author: { type: 'string', description: '只处理该作者的修订；不填处理全部' },
        },
        required: ['mode'],
    },
};
const resolveRevisionsHandler = async (args) => {
    const { mode, author } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('resolveRevisions', { mode: mode === 'reject' ? 'reject' : 'accept', author }, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `处理修订失败: ${response.error}` }], error: response.error };
        const d = response.data || { mode: 'accept', resolved: 0, remaining: 0 };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `已${d.mode === 'reject' ? '拒绝' : '接受'} ${d.resolved} 处修订；还剩 ${d.remaining} 处` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `处理修订出错: ${errMsg}` }], error: errMsg };
    }
};
exports.resolveRevisionsHandler = resolveRevisionsHandler;
exports.createLabelSheetDefinition = {
    name: 'wps_word_create_labels',
    description: '新建标签/席卡版式文档：按行列、标签宽高、页边距建网格（WPS 的 MailingLabel 在 12.1 上返回空文档，故用表格版式实现）。建好后可用 mail_merge 往当前文档填数据。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            rows: { type: 'number', description: '行数，默认5' },
            columns: { type: 'number', description: '列数，默认2' },
            labelWidth: { type: 'number', description: '单个标签宽度（磅），默认200' },
            labelHeight: { type: 'number', description: '单个标签高度（磅），默认100' },
            margin: { type: 'number', description: '页边距（磅），默认36' },
        },
        required: [],
    },
};
const createLabelSheetHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('createLabelSheet', args, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `创建标签版式失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `标签版式已创建: ${response.data?.rows} 行 × ${response.data?.columns} 列` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `创建标签版式出错: ${errMsg}` }], error: errMsg };
    }
};
exports.createLabelSheetHandler = createLabelSheetHandler;
exports.createEnvelopeDefinition = {
    name: 'wps_word_create_envelope',
    description: '新建信封文档：可设宽高（磅）、收件人/发件人文字；用于打印信封或做席卡底稿。',
    category: tools_1.ToolCategory.DOCUMENT,
    inputSchema: {
        type: 'object',
        properties: {
            width: { type: 'number', description: '信封宽度（磅），默认 A4 宽' },
            height: { type: 'number', description: '信封高度（磅），默认 A4 高' },
            recipient: { type: 'string', description: '收件人文字' },
            sender: { type: 'string', description: '发件人文字' },
        },
        required: [],
    },
};
const createEnvelopeHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('createEnvelope', args, wps_1.WpsAppType.WRITER);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `创建信封失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `信封已创建: ${response.data?.width} × ${response.data?.height} 磅` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `创建信封出错: ${errMsg}` }], error: errMsg };
    }
};
exports.createEnvelopeHandler = createEnvelopeHandler;
exports.wordStructureTools = [
    { definition: exports.getFieldsDefinition, handler: exports.getFieldsHandler },
    { definition: exports.updateFieldsDefinition, handler: exports.updateFieldsHandler },
    { definition: exports.selectRangeDefinition, handler: exports.selectRangeHandler },
    { definition: exports.protectDocumentDefinition, handler: exports.protectDocumentHandler },
    { definition: exports.getProtectionDefinition, handler: exports.getProtectionHandler },
    { definition: exports.insertCaptionDefinition, handler: exports.insertCaptionHandler },
    { definition: exports.setTableFormulaDefinition, handler: exports.setTableFormulaHandler },
    { definition: exports.insertTableOfFiguresDefinition, handler: exports.insertTableOfFiguresHandler },
    { definition: exports.compareDocumentsDefinition, handler: exports.compareDocumentsHandler },
    { definition: exports.updateStyleDefinition, handler: exports.updateStyleHandler },
    { definition: exports.setContentControlDefinition, handler: exports.setContentControlHandler },
    { definition: exports.resolveRevisionsDefinition, handler: exports.resolveRevisionsHandler },
    { definition: exports.createLabelSheetDefinition, handler: exports.createLabelSheetHandler },
    { definition: exports.createEnvelopeDefinition, handler: exports.createEnvelopeHandler },
];
//# sourceMappingURL=structure.js.map
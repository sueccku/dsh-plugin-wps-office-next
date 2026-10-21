"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.HELP_SCENARIOS = exports.HELP_SYNONYMS = exports.STANDARD_TOOLS = exports.FACADE_TOOLS = exports.DEFAULT_MODE = void 0;
exports.resolveMode = resolveMode;
exports.selectToolNames = selectToolNames;
exports.selectTools = selectTools;
exports.compactDescription = compactDescription;
exports.tokenizeHelpQuery = tokenizeHelpQuery;
exports.scoreToolHelpMatch = scoreToolHelpMatch;
exports.expandHelpQuery = expandHelpQuery;
exports.scoreToolHelpMatchExpanded = scoreToolHelpMatchExpanded;
exports.matchHelpScenarios = matchHelpScenarios;
/** Mode used when WPS_OFFICE_TOOLSET is unset or invalid. */
exports.DEFAULT_MODE = 'standard';
/** Tools that are always advertised: the dispatcher facade. */
exports.FACADE_TOOLS = [
    'wps_status',
    'wps_help',
    'wps_call',
    'wps_batch',
];
/**
 * Curated high-frequency tools for the standard mode. Everything else stays reachable
 * through wps_call; wps_help lists and describes the hidden tail on demand.
 */
exports.STANDARD_TOOLS = [
    // Excel
    'wps_excel_get_sheet_list',
    'wps_excel_read_range',
    'wps_excel_write_range',
    'wps_excel_set_cell_format',
    'wps_excel_set_number_format',
    'wps_excel_set_formula',
    'wps_excel_create_chart',
    'wps_excel_create_pivot_table',
    'wps_excel_find_replace',
    'wps_excel_get_open_workbooks',
    'wps_excel_open_workbook',
    'wps_excel_get_sheet_info',
    'wps_excel_auto_fit',
    'wps_excel_find_in_sheet',
    'wps_excel_get_named_ranges',
    'wps_excel_copy_format',
    'wps_excel_clear_formats',
    'wps_excel_calculate',
    'wps_excel_create_list_object',
    'wps_excel_get_list_objects',
    'wps_excel_add_list_row',
    'wps_excel_get_sheet_settings',
    'wps_excel_set_sheet_page_setup',
    'wps_excel_goal_seek',
    'wps_excel_clear_pivot_table',
    'wps_excel_set_chart_labels',
    // FIXES 90（D1-C 重平衡）：下面是真实调用数据里最常被 wps_help + wps_call 两步绕进来的工具
    // （742 次调用统计：close_workbook 36 次、create_sheet 16 次、export_chart_as_image 17 次、
    //  get_conditional_formats 15 次…）。把它们提到广告面，省掉"查一次 + 调一次"的往返与猜名失败。
    'wps_excel_close_workbook',
    'wps_excel_create_sheet',
    'wps_excel_switch_sheet',
    'wps_excel_switch_workbook',
    'wps_excel_export_chart_as_image',
    'wps_excel_export_range_as_image',
    'wps_excel_get_conditional_formats',
    'wps_excel_set_conditional_format',
    'wps_excel_set_sheet_header_footer',
    'wps_excel_set_sheet_print_titles',
    'wps_excel_get_cell_info',
    'wps_excel_sort_range',
    // P2-2（2026-10-21）：80 场景里反复出现的外链排查入口；换入后广告面到 100。
    'wps_excel_get_links',
    // D4（2026-10-11，场景 S6 走查后用户拍板）：把真实走查里**被需要、却只能用 wps_help + wps_call 两步绕进来**的
    // 高频隐藏工具提到广告面。取证口径：S1/S4/S6 六轮真机轨迹里经 wps_call / wps_batch 触达的隐藏工具，
    // 按「跨轮次出现的场景数 → 调用总次数」排序（复算脚本 test/.artifacts/scenario-s6/tools/rank-hidden-tools.mjs、
    // 预算测算 plan-ad-surface.mjs）。准入阈值：**跨 ≥2 轮出现，或单轮被调用 ≥5 次**。
    // 最硬的一条依据：wps_excel_freeze_panes 在 S6 attempt 1 被模型直接调用时拿到 `unknown tool`（一次硬失败 + 一次重试）。
    // 本轮仍留在隐藏面、等下一次用调用数据再评估的：paste_range(2)、set_print_area(2)、word_find_in_document(2)、
    // hide_column(2)、word_insert_table(1)、word_insert_header(1)、remove_duplicates(1)、update_chart(1)、delete_chart(1)。
    'wps_excel_evaluate_formula',
    'wps_excel_set_column_width',
    'wps_excel_copy_range',
    'wps_excel_set_sheet_appearance',
    'wps_excel_unmerge_cells',
    'wps_excel_delete_rows',
    'wps_excel_freeze_panes',
    'wps_excel_clear_range',
    // Word
    'wps_word_get_active_document',
    'wps_word_get_document_text',
    'wps_word_get_paragraphs',
    'wps_word_insert_text',
    'wps_word_find_replace',
    // D4：S4 重跑里同一个工具被调用 21 次（按范围替换正文），是 Word 侧绕路最多的一条。
    'wps_word_replace_range',
    'wps_word_apply_style',
    'wps_word_set_font',
    'wps_word_set_paragraph',
    'wps_word_generate_toc',
    // WO-001：页面设置的读与写一起进广告面。S4 真机走查里模型要读页边距/方向却只能直呼一个不存在的
    // 工具、最后绕开插件改 OOXML；而"把页边距改成 2cm"这类要求此前也要靠 wps_help + wps_call 两步
    // 才够得到 set。set/get 分居广告面内外是这次补完的最后一个不对称。
    'wps_word_get_page_setup',
    'wps_word_set_page_setup',
    'wps_word_smart_fill_field',
    'wps_word_open_document',
    'wps_word_create_document',
    'wps_word_get_document_stats',
    'wps_word_get_tables',
    'wps_word_get_table_data',
    'wps_word_get_comments',
    'wps_word_set_table_cell',
    'wps_word_insert_page_numbers',
    'wps_word_get_revisions',
    'wps_word_mail_merge',
    'wps_word_get_notes',
    // P2-2（2026-10-21）：80 场景里的域、保护、比较是 Word 长文交付的三条高频缺口。
    'wps_word_get_fields',
    'wps_word_protect',
    'wps_word_compare_documents',
    // FIXES 90：Word 侧同样缺"收尾"与"多文档"这两个高频动作（get_open_documents 19 次、close_document 15 次、
    // switch_document 2 次）。广告面以前只有 open/create，没有 close —— 生命周期不对称。
    'wps_word_get_open_documents',
    'wps_word_close_document',
    'wps_word_switch_document',
    // Presentation
    'wps_ppt_get_slide_count',
    'wps_ppt_get_slide_info',
    'wps_ppt_get_shapes',
    'wps_ppt_set_shape_fill',
    'wps_ppt_set_shape_text',
    'wps_ppt_add_slide',
    'wps_ppt_set_slide_title',
    'wps_ppt_set_slide_content',
    'wps_ppt_add_textbox',
    'wps_ppt_insert_table',
    'wps_ppt_insert_ppt_image',
    'wps_ppt_set_table_format',
    'wps_ppt_export_slide_as_image',
    'wps_ppt_open_presentation',
    // P2-2（2026-10-21）：备注导出是 PT10 的收尾动作；本地写文件，不依赖桥。
    'wps_ppt_export_notes',
    // Common and conversion
    'wps_common_save',
    'wps_common_save_as',
    'wps_convert_to_pdf',
    'wps_convert_format',
];
/** Resolve a raw environment value into a supported mode. */
function resolveMode(raw) {
    const value = (raw || '').trim().toLowerCase();
    if (value === 'minimal' || value === 'standard' || value === 'full')
        return value;
    return exports.DEFAULT_MODE;
}
/** Names advertised for a mode, in catalog order. */
function selectToolNames(mode, allNames) {
    if (mode === 'full')
        return [...allNames];
    const wanted = new Set(exports.FACADE_TOOLS);
    if (mode === 'standard')
        for (const name of exports.STANDARD_TOOLS)
            wanted.add(name);
    return allNames.filter((name) => wanted.has(name));
}
/** Tool definitions advertised for a mode. */
function selectTools(mode, tools) {
    const names = new Set(selectToolNames(mode, tools.map((t) => t.name)));
    return tools.filter((t) => names.has(t.name));
}
/**
 * Shorten a tool description for the advertised surface. Full text stays available
 * through wps_help, so the per-request copy only needs to route the model correctly.
 */
function compactDescription(text, max = 110) {
    const flat = (text || '').replace(/\s+/g, ' ').trim();
    if (flat.length <= max)
        return flat;
    const stop = flat.search(/[。.!?]/);
    if (stop >= 0 && stop + 1 <= max)
        return flat.slice(0, stop + 1);
    return flat.slice(0, max - 1) + '…';
}
/**
 * Split a wps_help free-text query into searchable tokens.
 * Whitespace and punctuation separate tokens; CJK runs stay whole because there is
 * nothing to split on, and the bigram path below covers those.
 */
function tokenizeHelpQuery(query) {
    return (query || '')
        .toLowerCase()
        .split(/[\s,，、;；:：/|+()（）\[\]【】"'“”]+/)
        .map((token) => token.trim())
        .filter((token) => token.length > 0);
}
/**
 * Relevance score for a wps_help query against one tool; 0 means "not a match".
 *
 * A model asks in whatever words fit its task ("close workbook 关闭", "新建 文档 create
 * new"), so a whole-phrase substring test answered "matched: 0" for queries a human would
 * call obvious. Score per token instead, and give CJK queries a bigram path: a space-less
 * query like "关闭工作簿" is not a substring of "关闭指定的Excel工作簿" but shares three of
 * its four bigrams.
 */
function scoreToolHelpMatch(name, description, query, tokens) {
    const toolName = (name || '').toLowerCase();
    const toolDescription = (description || '').toLowerCase();
    const phrase = (query || '').trim().toLowerCase();
    if (!phrase)
        return 0;
    let score = 0;
    if (toolName.includes(phrase))
        score += 12;
    else if (toolDescription.includes(phrase))
        score += 10;
    for (const token of tokens) {
        if (token.length < 2)
            continue;
        if (toolName.includes(token))
            score += 4;
        else if (toolDescription.includes(token))
            score += 2;
    }
    const cjk = [...phrase].filter((char) => /[\u3400-\u9fff]/.test(char));
    if (cjk.length >= 2) {
        const bigrams = new Set();
        for (let i = 0; i + 1 < cjk.length; i++)
            bigrams.add(cjk[i] + cjk[i + 1]);
        let hits = 0;
        for (const bigram of bigrams) {
            if (toolName.includes(bigram) || toolDescription.includes(bigram))
                hits++;
        }
        const coverage = hits / bigrams.size;
        if (coverage >= 0.5)
            score += 6 * coverage;
    }
    return score;
}
/**
 * User words that do not appear in any tool name/description but are the normal way a Chinese
 * office user names the task. Expanded terms are scored in addition to the original query; this
 * keeps the original bigram behaviour while letting "题注/文献/母版" reach the right family.
 */
exports.HELP_SYNONYMS = {
    '题注': ['caption', '题注', '编号', 'seq', '交叉引用'],
    '交叉引用': ['cross reference', '交叉引用', '引用', 'ref', '题注'],
    '图表目录': ['table of figures', '图表目录', '目录', 'toc'],
    '域': ['field', '域', '更新域', 'fields'],
    '更新域': ['update fields', '更新域', '域', '刷新'],
    '奇偶页': ['odd even', '奇偶页', '页眉', 'header'],
    '页眉横线': ['header border', '页眉横线', '边框', '页眉'],
    '文档比较': ['compare', '比较', '文档比较', '修订', '对比'],
    '标签': ['label', '标签', '邮件合并', '信封'],
    '信封': ['envelope', '信封', '标签', '邮件合并'],
    '席卡': ['label', '桌贴', '席卡', '邮件合并'],
    '内容控件': ['content control', '内容控件', '表单', '控件'],
    '限制编辑': ['restrict editing', '限制编辑', '保护', '只读', '窗体'],
    '字体嵌入': ['embed font', '字体嵌入', '嵌入'],
    '图片压缩': ['compress picture', '图片压缩', '压缩', '图片'],
    '隐藏页': ['hidden slide', '隐藏页', '隐藏幻灯片'],
    '备注页': ['notes page', '备注页', '备注'],
    '母版': ['master', '母版', '版式', 'layout'],
    '主题色': ['theme color', '主题色', '主题', 'theme'],
};
/** Expand a free-text query with the synonym groups it contains (deduplicated, original first). */
function expandHelpQuery(query) {
    const phrase = (query || '').trim().toLowerCase();
    if (!phrase)
        return [];
    const out = [phrase];
    for (const [key, extras] of Object.entries(exports.HELP_SYNONYMS)) {
        if (!phrase.includes(key.toLowerCase()))
            continue;
        for (const extra of extras) {
            const term = extra.toLowerCase();
            if (!out.includes(term))
                out.push(term);
        }
    }
    return out;
}
/** Score a query with synonym expansion; best match across the original and expanded terms. */
function scoreToolHelpMatchExpanded(name, description, query, tokens) {
    let best = scoreToolHelpMatch(name, description, query, tokens);
    for (const term of expandHelpQuery(query)) {
        if (term === query.trim().toLowerCase())
            continue;
        const termTokens = tokenizeHelpQuery(term);
        const score = scoreToolHelpMatch(name, description, term, termTokens);
        if (score > best)
            best = score;
    }
    return best;
}
exports.HELP_SCENARIOS = [
    { keywords: ['批量合同', '一人一份', '邮件合并', '合同'], tools: ['wps_word_mail_merge'], recipe: 'skills/wps-word/recipes/fill-template.md' },
    { keywords: ['目录页码', '目录', '页码'], tools: ['wps_word_generate_toc', 'wps_word_insert_page_numbers'], recipe: 'skills/wps-word/SKILL.md#6-目录与页码' },
    { keywords: ['长文排版', '论文', '公文'], tools: ['wps_word_get_paragraphs', 'wps_word_set_font', 'wps_word_apply_style'], recipe: 'skills/wps-word/recipes/format-existing-report.md' },
    { keywords: ['多表合并', '多工作簿', '总表'], tools: ['wps_excel_read_range', 'wps_excel_write_range', 'wps_excel_open_workbook'], recipe: 'skills/wps-excel/recipes/summarize-report.md' },
    { keywords: ['透视表', '透视'], tools: ['wps_excel_create_pivot_table', 'wps_excel_get_pivot_tables', 'wps_excel_refresh_pivot_tables'], recipe: 'skills/wps-excel/SKILL.md' },
    { keywords: ['条件格式', '整行变色'], tools: ['wps_excel_set_conditional_format', 'wps_excel_get_conditional_formats'], recipe: 'skills/wps-excel/SKILL.md' },
    { keywords: ['打印', '打印预览', '页眉页脚'], tools: ['wps_excel_set_sheet_page_setup', 'wps_excel_set_sheet_print_titles', 'wps_excel_set_sheet_header_footer'], recipe: 'skills/wps-excel/SKILL.md' },
    { keywords: ['跨应用', 'excel到ppt', '数据到ppt', '汇报'], tools: ['wps_ppt_insert_ppt_chart', 'wps_ppt_set_ppt_chart_data'], recipe: 'skills/wps-office-next/recipes/cross-app-report.md' },
    { keywords: ['文档比较', '两版', '对比'], tools: ['wps_word_get_revisions'], recipe: 'skills/wps-word/SKILL.md', note: '当前无 compareDocuments 工具，先按"比较结果修订"处理或人工比较' },
    { keywords: ['限制编辑', '只读', '保护'], tools: [], recipe: 'skills/wps-word/SKILL.md', note: 'Word 侧尚无 protect/restrict 工具；可先用文件只读或交付 PDF' },
];
/** Return scenario hints matching a user query. */
function matchHelpScenarios(query) {
    const phrase = (query || '').trim().toLowerCase();
    if (!phrase)
        return [];
    return exports.HELP_SCENARIOS.filter((hint) => hint.keywords.some((word) => phrase.includes(word.toLowerCase())));
}
//# sourceMappingURL=toolset.js.map
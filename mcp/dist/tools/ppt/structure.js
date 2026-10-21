"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.pptStructureTools = exports.exportPdfHandler = exports.exportPdfDefinition = exports.compressImagesHandler = exports.compressImagesDefinition = exports.embedFontsHandler = exports.embedFontsDefinition = exports.duplicateMasterLayoutHandler = exports.duplicateMasterLayoutDefinition = exports.setPlaceholderLockHandler = exports.setPlaceholderLockDefinition = exports.addSectionHandler = exports.addSectionDefinition = exports.setThemeFontsHandler = exports.setThemeFontsDefinition = exports.setThemeColorsHandler = exports.setThemeColorsDefinition = exports.getFontsHandler = exports.getFontsDefinition = exports.getHiddenSlidesHandler = exports.getHiddenSlidesDefinition = exports.setSlideHiddenHandler = exports.setSlideHiddenDefinition = void 0;
/**
 * Input: PPT 隐藏页 / 字体回读参数
 * Output: P1（2026-10-19）新增的 PPT 结构能力结果
 * Pos: 工具默认隐藏；COM 细节在 wps-com.ps1 的 setSlideHidden/getHiddenSlides/getSlideFonts。
 */
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
exports.setSlideHiddenDefinition = {
    name: 'wps_ppt_set_slide_hidden',
    description: '设置/取消幻灯片隐藏（hidden=true 隐藏，false 显示）。隐藏页在放映时不出现；导出 PDF 是否包含取决于导出选项。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片页码（从1开始）' },
            hidden: { type: 'boolean', description: 'true 隐藏（默认），false 显示' },
        },
        required: ['slideIndex'],
    },
};
const setSlideHiddenHandler = async (args) => {
    const { slideIndex, hidden } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSlideHidden', { slideIndex, hidden: hidden !== false }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置隐藏页失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `第 ${response.data?.slideIndex} 页已${response.data?.hidden ? '隐藏' : '显示'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置隐藏页出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setSlideHiddenHandler = setSlideHiddenHandler;
exports.getHiddenSlidesDefinition = {
    name: 'wps_ppt_get_hidden_slides',
    description: '列出演示文稿里被隐藏的幻灯片页码与总数，用于导出/放映前的检查。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: { type: 'object', properties: {}, required: [] },
};
const getHiddenSlidesHandler = async () => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getHiddenSlides', {}, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `读取隐藏页失败: ${response.error}` }], error: response.error };
        const d = response.data || { count: 0, hiddenSlides: [], totalSlides: 0 };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `隐藏页 ${d.count}/${d.totalSlides}: ${d.hiddenSlides.join(', ') || '无'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `读取隐藏页出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getHiddenSlidesHandler = getHiddenSlidesHandler;
exports.getFontsDefinition = {
    name: 'wps_ppt_get_fonts',
    description: '回读幻灯片文字的字体名/字号/加粗/颜色（每页去重）。用于"只改字色不改字号"这类任务的前后核对；不传 slideIndex 读全篇。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: { slideIndex: { type: 'number', description: '只读这一页；不填读全部' } },
        required: [],
    },
};
const getFontsHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSlideFonts', slideIndex === undefined ? {} : { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `读取字体失败: ${response.error}` }], error: response.error };
        const d = response.data || { count: 0, fonts: [] };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `字体条目: ${d.count}\n${JSON.stringify(d.fonts, null, 2)}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `读取字体出错: ${errMsg}` }], error: errMsg };
    }
};
exports.getFontsHandler = getFontsHandler;
exports.setThemeColorsDefinition = {
    name: 'wps_ppt_set_theme_colors',
    description: '改演示文稿主题配色（dark1/light1/dark2/light2/accent1..6/hyperlink/followedHyperlink，十六进制）；只改变量，不逐个改形状。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            dark1: { type: 'string', description: '#RRGGBB' }, light1: { type: 'string', description: '#RRGGBB' },
            dark2: { type: 'string', description: '#RRGGBB' }, light2: { type: 'string', description: '#RRGGBB' },
            accent1: { type: 'string', description: '#RRGGBB' }, accent2: { type: 'string', description: '#RRGGBB' },
            accent3: { type: 'string', description: '#RRGGBB' }, accent4: { type: 'string', description: '#RRGGBB' },
            accent5: { type: 'string', description: '#RRGGBB' }, accent6: { type: 'string', description: '#RRGGBB' },
            hyperlink: { type: 'string', description: '#RRGGBB' }, followedHyperlink: { type: 'string', description: '#RRGGBB' },
        },
        required: [],
    },
};
const setThemeColorsHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setThemeColors', args, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置主题色失败: ${response.error}` }], error: response.error };
        const applied = response.data?.applied || {};
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `主题色已更新: ${Object.keys(applied).join(', ') || '（无变化）'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置主题色出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setThemeColorsHandler = setThemeColorsHandler;
exports.setThemeFontsDefinition = {
    name: 'wps_ppt_set_theme_fonts',
    description: '改演示文稿主题字体（标题字体 majorFont / 正文字体 minorFont）；只改主题字体，不逐个文本框改。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            majorFont: { type: 'string', description: '标题字体名' },
            minorFont: { type: 'string', description: '正文字体名' },
        },
        required: [],
    },
};
const setThemeFontsHandler = async (args) => {
    const { majorFont, minorFont } = args;
    if (!majorFont && !minorFont)
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: 'majorFont 与 minorFont 至少给一个' }], error: '缺少字体' };
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setThemeFonts', { majorFont, minorFont }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置主题字体失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `主题字体已更新: 标题=${response.data?.majorFont || '不变'}，正文=${response.data?.minorFont || '不变'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置主题字体出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setThemeFontsHandler = setThemeFontsHandler;
exports.addSectionDefinition = {
    name: 'wps_ppt_add_section',
    description: '在指定页前新增一节（PowerPoint 节），便于 30+ 页长 deck 分组。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            beforeSlide: { type: 'number', description: '在第几页前插入节（从1开始），默认1' },
            name: { type: 'string', description: '节名' },
        },
        required: [],
    },
};
const addSectionHandler = async (args) => {
    const { beforeSlide, name } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addSection', { beforeSlide, name }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `新增节失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `已新增节「${response.data?.name}」；共 ${response.data?.sections} 节` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `新增节出错: ${errMsg}` }], error: errMsg };
    }
};
exports.addSectionHandler = addSectionHandler;
exports.setPlaceholderLockDefinition = {
    name: 'wps_ppt_set_placeholder_lock',
    description: '锁定/解锁占位符或形状（Shape.Locked），防止误拖误改；按 shapeIndex 或 shapeName 定位。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '页码（从1开始），默认1' },
            shapeIndex: { type: 'number', description: '形状序号（从1开始）' },
            shapeName: { type: 'string', description: '形状名' },
            locked: { type: 'boolean', description: 'true 锁定，false 解锁' },
        },
        required: ['locked'],
    },
};
const setPlaceholderLockHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setPlaceholderLock', args, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置占位符锁定失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `${response.data?.shape} 已${response.data?.locked ? '锁定' : '解锁'}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `设置占位符锁定出错: ${errMsg}` }], error: errMsg };
    }
};
exports.setPlaceholderLockHandler = setPlaceholderLockHandler;
exports.duplicateMasterLayoutDefinition = {
    name: 'wps_ppt_duplicate_master_layout',
    description: '复制母版版式（CustomLayouts.Duplicate）；WPS 12.1 的 CustomLayouts 没有 Add，复制现有版式是新增版式的可行路径。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: { layoutIndex: { type: 'number', description: '第几个母版版式（从1开始），默认1' } },
        required: [],
    },
};
const duplicateMasterLayoutHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('duplicateMasterLayout', args, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `复制母版版式失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `版式已复制: ${response.data?.source} → ${response.data?.duplicated}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `复制母版版式出错: ${errMsg}` }], error: errMsg };
    }
};
exports.duplicateMasterLayoutHandler = duplicateMasterLayoutHandler;
exports.embedFontsDefinition = {
    name: 'wps_ppt_embed_fonts',
    description: '把当前演示文稿另存一份并打开字体嵌入标志（SaveCopyAs 第三个参数 EmbedTrueTypeFonts）；原演示文稿路径不变。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: { outputPath: { type: 'string', description: '输出 .pptx 路径' } },
        required: ['outputPath'],
    },
};
const embedFontsHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('embedFonts', args, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `嵌入字体失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `字体嵌入副本已保存: ${response.data?.outputPath}` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `嵌入字体出错: ${errMsg}` }], error: errMsg };
    }
};
exports.embedFontsHandler = embedFontsHandler;
exports.compressImagesDefinition = {
    name: 'wps_ppt_compress_images',
    description: '压缩幻灯片图片：把每张图按 quality 比例导出为 JPG 再替换（保持原位置/尺寸），会损失清晰度；不传 slideIndex 处理全篇。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            quality: { type: 'number', description: '保留比例（0–1），默认 0.6' },
            slideIndex: { type: 'number', description: '只处理这一页；不填全篇' },
        },
        required: [],
    },
};
const compressImagesHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('compressImages', args, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `压缩图片失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `已压缩 ${response.data?.images ?? 0} 张图片（保留比例 ${response.data?.quality}）` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `压缩图片出错: ${errMsg}` }], error: errMsg };
    }
};
exports.compressImagesHandler = compressImagesHandler;
exports.exportPdfDefinition = {
    name: 'wps_ppt_export_pdf',
    description: '把演示文稿导出为 PDF（Presentation.Export，可指定像素宽高控制缩放/清晰度）。导出的是幻灯片页；备注页导出请用 export_notes。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            outputPath: { type: 'string', description: 'PDF 输出路径' },
            width: { type: 'number', description: '导出宽度（像素），默认1280' },
            height: { type: 'number', description: '导出高度（像素），默认720' },
        },
        required: ['outputPath'],
    },
};
const exportPdfHandler = async (args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('exportPdf', args, wps_1.WpsAppType.PRESENTATION);
        if (!response.success)
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `导出 PDF 失败: ${response.error}` }], error: response.error };
        return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `PDF 已导出: ${response.data?.outputPath}（${response.data?.width}×${response.data?.height}）` }] };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `导出 PDF 出错: ${errMsg}` }], error: errMsg };
    }
};
exports.exportPdfHandler = exportPdfHandler;
exports.pptStructureTools = [
    { definition: exports.setSlideHiddenDefinition, handler: exports.setSlideHiddenHandler },
    { definition: exports.getHiddenSlidesDefinition, handler: exports.getHiddenSlidesHandler },
    { definition: exports.getFontsDefinition, handler: exports.getFontsHandler },
    { definition: exports.setThemeColorsDefinition, handler: exports.setThemeColorsHandler },
    { definition: exports.setThemeFontsDefinition, handler: exports.setThemeFontsHandler },
    { definition: exports.addSectionDefinition, handler: exports.addSectionHandler },
    { definition: exports.setPlaceholderLockDefinition, handler: exports.setPlaceholderLockHandler },
    { definition: exports.duplicateMasterLayoutDefinition, handler: exports.duplicateMasterLayoutHandler },
    { definition: exports.embedFontsDefinition, handler: exports.embedFontsHandler },
    { definition: exports.compressImagesDefinition, handler: exports.compressImagesHandler },
    { definition: exports.exportPdfDefinition, handler: exports.exportPdfHandler },
];
//# sourceMappingURL=structure.js.map
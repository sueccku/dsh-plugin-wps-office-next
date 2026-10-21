/**
 * Input: PPT 隐藏页 / 字体回读参数
 * Output: P1（2026-10-19）新增的 PPT 结构能力结果
 * Pos: 工具默认隐藏；COM 细节在 wps-com.ps1 的 setSlideHidden/getHiddenSlides/getSlideFonts。
 */
import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

export const setSlideHiddenDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_hidden',
  description: '设置/取消幻灯片隐藏（hidden=true 隐藏，false 显示）。隐藏页在放映时不出现；导出 PDF 是否包含取决于导出选项。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片页码（从1开始）' },
      hidden: { type: 'boolean', description: 'true 隐藏（默认），false 显示' },
    },
    required: ['slideIndex'],
  },
};

export const setSlideHiddenHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { slideIndex, hidden } = args as { slideIndex: number; hidden?: boolean };
  try {
    const response = await wpsClient.executeMethod<{ slideIndex: number; hidden: boolean }>(
      'setSlideHidden', { slideIndex, hidden: hidden !== false }, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置隐藏页失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `第 ${response.data?.slideIndex} 页已${response.data?.hidden ? '隐藏' : '显示'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置隐藏页出错: ${errMsg}` }], error: errMsg };
  }
};

export const getHiddenSlidesDefinition: ToolDefinition = {
  name: 'wps_ppt_get_hidden_slides',
  description: '列出演示文稿里被隐藏的幻灯片页码与总数，用于导出/放映前的检查。',
  category: ToolCategory.PRESENTATION,
  inputSchema: { type: 'object', properties: {}, required: [] },
};

export const getHiddenSlidesHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ count: number; hiddenSlides: number[]; totalSlides: number }>(
      'getHiddenSlides', {}, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取隐藏页失败: ${response.error}` }], error: response.error };
    const d = response.data || { count: 0, hiddenSlides: [], totalSlides: 0 };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `隐藏页 ${d.count}/${d.totalSlides}: ${d.hiddenSlides.join(', ') || '无'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取隐藏页出错: ${errMsg}` }], error: errMsg };
  }
};

export const getFontsDefinition: ToolDefinition = {
  name: 'wps_ppt_get_fonts',
  description: '回读幻灯片文字的字体名/字号/加粗/颜色（每页去重）。用于"只改字色不改字号"这类任务的前后核对；不传 slideIndex 读全篇。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: { slideIndex: { type: 'number', description: '只读这一页；不填读全部' } },
    required: [],
  },
};

export const getFontsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex?: number };
  try {
    const response = await wpsClient.executeMethod<{ count: number; fonts: Array<Record<string, unknown>> }>(
      'getSlideFonts', slideIndex === undefined ? {} : { slideIndex }, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取字体失败: ${response.error}` }], error: response.error };
    const d = response.data || { count: 0, fonts: [] };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `字体条目: ${d.count}\n${JSON.stringify(d.fonts, null, 2)}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取字体出错: ${errMsg}` }], error: errMsg };
  }
};

export const setThemeColorsDefinition: ToolDefinition = {
  name: 'wps_ppt_set_theme_colors',
  description: '改演示文稿主题配色（dark1/light1/dark2/light2/accent1..6/hyperlink/followedHyperlink，十六进制）；只改变量，不逐个改形状。',
  category: ToolCategory.PRESENTATION,
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

export const setThemeColorsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ applied: Record<string, string> }>(
      'setThemeColors', args, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置主题色失败: ${response.error}` }], error: response.error };
    const applied = response.data?.applied || {};
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `主题色已更新: ${Object.keys(applied).join(', ') || '（无变化）'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置主题色出错: ${errMsg}` }], error: errMsg };
  }
};

export const setThemeFontsDefinition: ToolDefinition = {
  name: 'wps_ppt_set_theme_fonts',
  description: '改演示文稿主题字体（标题字体 majorFont / 正文字体 minorFont）；只改主题字体，不逐个文本框改。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      majorFont: { type: 'string', description: '标题字体名' },
      minorFont: { type: 'string', description: '正文字体名' },
    },
    required: [],
  },
};

export const setThemeFontsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { majorFont, minorFont } = args as { majorFont?: string; minorFont?: string };
  if (!majorFont && !minorFont) return { id: uuidv4(), success: false, content: [{ type: 'text', text: 'majorFont 与 minorFont 至少给一个' }], error: '缺少字体' };
  try {
    const response = await wpsClient.executeMethod<{ majorFont?: string; minorFont?: string }>(
      'setThemeFonts', { majorFont, minorFont }, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置主题字体失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `主题字体已更新: 标题=${response.data?.majorFont || '不变'}，正文=${response.data?.minorFont || '不变'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置主题字体出错: ${errMsg}` }], error: errMsg };
  }
};

export const addSectionDefinition: ToolDefinition = {
  name: 'wps_ppt_add_section',
  description: '在指定页前新增一节（PowerPoint 节），便于 30+ 页长 deck 分组。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      beforeSlide: { type: 'number', description: '在第几页前插入节（从1开始），默认1' },
      name: { type: 'string', description: '节名' },
    },
    required: [],
  },
};

export const addSectionHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { beforeSlide, name } = args as { beforeSlide?: number; name?: string };
  try {
    const response = await wpsClient.executeMethod<{ name: string; sections: number }>(
      'addSection', { beforeSlide, name }, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `新增节失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `已新增节「${response.data?.name}」；共 ${response.data?.sections} 节` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `新增节出错: ${errMsg}` }], error: errMsg };
  }
};

export const setPlaceholderLockDefinition: ToolDefinition = {
  name: 'wps_ppt_set_placeholder_lock',
  description: '锁定/解锁占位符或形状（Shape.Locked），防止误拖误改；按 shapeIndex 或 shapeName 定位。',
  category: ToolCategory.PRESENTATION,
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

export const setPlaceholderLockHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ shape: string; locked: boolean }>(
      'setPlaceholderLock', args, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置占位符锁定失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `${response.data?.shape} 已${response.data?.locked ? '锁定' : '解锁'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置占位符锁定出错: ${errMsg}` }], error: errMsg };
  }
};

export const duplicateMasterLayoutDefinition: ToolDefinition = {
  name: 'wps_ppt_duplicate_master_layout',
  description: '复制母版版式（CustomLayouts.Duplicate）；WPS 12.1 的 CustomLayouts 没有 Add，复制现有版式是新增版式的可行路径。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: { layoutIndex: { type: 'number', description: '第几个母版版式（从1开始），默认1' } },
    required: [],
  },
};

export const duplicateMasterLayoutHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ source: string; duplicated: string; count: number }>(
      'duplicateMasterLayout', args, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `复制母版版式失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `版式已复制: ${response.data?.source} → ${response.data?.duplicated}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `复制母版版式出错: ${errMsg}` }], error: errMsg };
  }
};

export const embedFontsDefinition: ToolDefinition = {
  name: 'wps_ppt_embed_fonts',
  description: '把当前演示文稿另存一份并打开字体嵌入标志（SaveCopyAs 第三个参数 EmbedTrueTypeFonts）；原演示文稿路径不变。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: { outputPath: { type: 'string', description: '输出 .pptx 路径' } },
    required: ['outputPath'],
  },
};

export const embedFontsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ outputPath: string; exists: boolean }>(
      'embedFonts', args, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `嵌入字体失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `字体嵌入副本已保存: ${response.data?.outputPath}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `嵌入字体出错: ${errMsg}` }], error: errMsg };
  }
};

export const compressImagesDefinition: ToolDefinition = {
  name: 'wps_ppt_compress_images',
  description: '压缩幻灯片图片：把每张图按 quality 比例导出为 JPG 再替换（保持原位置/尺寸），会损失清晰度；不传 slideIndex 处理全篇。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      quality: { type: 'number', description: '保留比例（0–1），默认 0.6' },
      slideIndex: { type: 'number', description: '只处理这一页；不填全篇' },
    },
    required: [],
  },
};

export const compressImagesHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ images: number; quality: number }>(
      'compressImages', args, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `压缩图片失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `已压缩 ${response.data?.images ?? 0} 张图片（保留比例 ${response.data?.quality}）` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `压缩图片出错: ${errMsg}` }], error: errMsg };
  }
};

export const exportPdfDefinition: ToolDefinition = {
  name: 'wps_ppt_export_pdf',
  description: '把演示文稿导出为 PDF（Presentation.Export，可指定像素宽高控制缩放/清晰度）。导出的是幻灯片页；备注页导出请用 export_notes。',
  category: ToolCategory.PRESENTATION,
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

export const exportPdfHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ outputPath: string; width: number; height: number }>(
      'exportPdf', args, WpsAppType.PRESENTATION
    );
    if (!response.success) return { id: uuidv4(), success: false, content: [{ type: 'text', text: `导出 PDF 失败: ${response.error}` }], error: response.error };
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `PDF 已导出: ${response.data?.outputPath}（${response.data?.width}×${response.data?.height}）` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `导出 PDF 出错: ${errMsg}` }], error: errMsg };
  }
};

export const pptStructureTools: RegisteredTool[] = [
  { definition: setSlideHiddenDefinition, handler: setSlideHiddenHandler },
  { definition: getHiddenSlidesDefinition, handler: getHiddenSlidesHandler },
  { definition: getFontsDefinition, handler: getFontsHandler },
  { definition: setThemeColorsDefinition, handler: setThemeColorsHandler },
  { definition: setThemeFontsDefinition, handler: setThemeFontsHandler },
  { definition: addSectionDefinition, handler: addSectionHandler },
  { definition: setPlaceholderLockDefinition, handler: setPlaceholderLockHandler },
  { definition: duplicateMasterLayoutDefinition, handler: duplicateMasterLayoutHandler },
  { definition: embedFontsDefinition, handler: embedFontsHandler },
  { definition: compressImagesDefinition, handler: compressImagesHandler },
  { definition: exportPdfDefinition, handler: exportPdfHandler },
];

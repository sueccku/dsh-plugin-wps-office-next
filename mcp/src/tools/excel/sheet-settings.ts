/**
 * Input: 页面设置/打印/工作表外观/公式审计 工具的调用参数
 * Output: 这些设置的执行结果与回读快照
 * Pos: Excel 页面与打印工具（P2-3）。Excel 此前完全没有页面设置能力（桥里的 setPageSetup 是 Word 的）。
 *      打印与打印预览刻意不做成工具：前者是物理副作用，后者会开模态窗口卡住常驻宿主。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P2 状态。
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

interface SheetSettings {
  sheet?: string;
  orientation?: number;
  orientationName?: string;
  paperSize?: number;
  zoom?: number | boolean;
  fitToPagesWide?: number;
  fitToPagesTall?: number;
  topMargin?: number;
  bottomMargin?: number;
  leftMargin?: number;
  rightMargin?: number;
  headerMargin?: number;
  footerMargin?: number;
  centerHorizontally?: boolean;
  centerVertically?: boolean;
  printGridlines?: boolean;
  printHeadings?: boolean;
  printArea?: string;
  printTitleRows?: string;
  printTitleColumns?: string;
  leftHeader?: string;
  centerHeader?: string;
  rightHeader?: string;
  leftFooter?: string;
  centerFooter?: string;
  rightFooter?: string;
  visible?: number;
  tabColor?: number;
  tabColorIndex?: number;
  hPageBreaks?: number;
}

interface SettingsResponse extends SheetSettings {
  applied?: string[];
  settings?: SheetSettings;
  message?: string;
}

const PAPER_NAMES: Record<number, string> = { 1: 'Letter', 3: 'Tabloid', 5: 'Legal', 8: 'A3', 9: 'A4', 11: 'A5', 13: 'B5' };

/** 清掉标签色之后 Tab.Color 读回 0（黑），只有 ColorIndex = -4142 才说明「用的是默认色」 */
function describeTabColor(s: SheetSettings): string {
  if (s.tabColorIndex === -4142) return '(默认)';
  if (s.tabColor === undefined || s.tabColor === null) return '(默认)';
  return String(s.tabColor);
}

function yesNo(value: unknown): string {
  return value ? '是' : '否';
}

/** 把快照整理成人能读的几行；空字段不占行 */
function describeSettings(s: SheetSettings): string {
  const paper = s.paperSize !== undefined && s.paperSize !== null ? (PAPER_NAMES[Number(s.paperSize)] || String(s.paperSize)) : '(未知)';
  const lines = ['工作表 ' + (s.sheet || '?') + ' 的页面设置:'];
  lines.push('  方向: ' + (s.orientationName || '?') + '；纸张: ' + paper);
  if (s.zoom === false) lines.push('  缩放: 按页适配（宽 ' + String(s.fitToPagesWide ?? '?') + ' 页 / 高 ' + String(s.fitToPagesTall ?? '?') + ' 页）');
  else lines.push('  缩放: ' + String(s.zoom ?? '?') + '%');
  lines.push('  页边距（磅）: 上 ' + String(s.topMargin ?? '?') + ' 下 ' + String(s.bottomMargin ?? '?') + ' 左 ' + String(s.leftMargin ?? '?') + ' 右 ' + String(s.rightMargin ?? '?') + ' 页眉 ' + String(s.headerMargin ?? '?') + ' 页脚 ' + String(s.footerMargin ?? '?'));
  lines.push('  居中: 水平 ' + yesNo(s.centerHorizontally) + ' / 垂直 ' + yesNo(s.centerVertically) + '；打印网格线 ' + yesNo(s.printGridlines) + '；打印行列标题 ' + yesNo(s.printHeadings));
  lines.push('  打印区域: ' + (s.printArea || '(未设置)') + '；打印标题: 行 ' + (s.printTitleRows || '(无)') + '，列 ' + (s.printTitleColumns || '(无)'));
  const headers = [s.leftHeader, s.centerHeader, s.rightHeader].filter(Boolean);
  if (headers.length) lines.push('  页眉: 左 ' + (s.leftHeader || '空') + ' / 中 ' + (s.centerHeader || '空') + ' / 右 ' + (s.rightHeader || '空'));
  const footers = [s.leftFooter, s.centerFooter, s.rightFooter].filter(Boolean);
  if (footers.length) lines.push('  页脚: 左 ' + (s.leftFooter || '空') + ' / 中 ' + (s.centerFooter || '空') + ' / 右 ' + (s.rightFooter || '空'));
  const vis = s.visible === 2 ? '深度隐藏' : s.visible === 0 ? '隐藏' : '可见';
  lines.push('  工作表: ' + vis + '；标签色: ' + describeTabColor(s) + '；手动分页符: ' + String(s.hPageBreaks ?? 0));
  return lines.join('\n');
}

function settingsFail(prefix: string, error?: string): ToolCallResult {
  return { id: uuidv4(), success: false, content: [{ type: 'text', text: prefix + ': ' + error }], error: error };
}

/** 一次返回“已应用 + 回读快照”给调用方，避免工具只说成功不说结果 */
function settingsResult(prefix: string, data: SettingsResponse | undefined): ToolCallResult {
  const applied = data?.applied || [];
  const head = applied.length ? prefix + '（已应用 ' + applied.join(', ') + '）' : prefix;
  return { id: uuidv4(), success: true, content: [{ type: 'text', text: head + '\n' + describeSettings(data?.settings || {}) }] };
}

const sheetParam = { type: 'string' as const, description: '工作表名或序号；不填则用当前活动工作表' };
/** 读全部页面/打印/外观设置 */
export const getSheetSettingsDefinition: ToolDefinition = {
  name: 'wps_excel_get_sheet_settings',
  description: '读一张工作表的页面设置、打印设置、页眉页脚与外观：方向、纸张、页边距（磅）、缩放或按页适配、是否居中、打印区域与打印标题、页眉页脚、可见性、标签色、手动分页符数量。使用场景：打印前先看清现状；改完再回读确认。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: { type: 'object', properties: { sheet: sheetParam } },
};

export const getSheetSettingsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<SheetSettings>(
      'getSheetSettings',
      { sheet: args.sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return settingsFail('读取页面设置失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: describeSettings(response.data || {}) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return settingsFail('读取页面设置出错', errMsg);
  }
};

/** 页面设置 */
export const setSheetPageSetupDefinition: ToolDefinition = {
  name: 'wps_excel_set_sheet_page_setup',
  description: '设置工作表的页面：方向、纸张、页边距、缩放、是否居中、是否打印网格线与行列标题。页边距单位是磅（1 厘米 ≈ 28.35 磅），与 Word 侧一致。缩放比例与按页适配互斥，同时给以后者为准。使用场景：把表调成横向 A4、一页宽、水平居中再打印。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      orientation: { type: 'string', enum: ['portrait', 'landscape'], description: '纸张方向，默认纵向' },
      paperSize: { type: 'string', enum: ['A4', 'A3', 'A5', 'B5', 'letter', 'legal', 'tabloid'], description: '纸张大小，默认不变（当前多为 A4）' },
      topMargin: { type: 'number', description: '上边距（磅）' },
      bottomMargin: { type: 'number', description: '下边距（磅）' },
      leftMargin: { type: 'number', description: '左边距（磅）' },
      rightMargin: { type: 'number', description: '右边距（磅）' },
      headerMargin: { type: 'number', description: '页眉距顶边（磅）' },
      footerMargin: { type: 'number', description: '页脚距底边（磅）' },
      zoom: { type: 'number', description: '缩放百分比（如 90）；与 fitToPages* 互斥' },
      fitToPagesWide: { type: 'number', description: '按页适配：横向压到几页宽（1 表示一页宽）' },
      fitToPagesTall: { type: 'number', description: '按页适配：纵向压到几页高' },
      centerHorizontally: { type: 'boolean', description: '水平居中打印' },
      centerVertically: { type: 'boolean', description: '垂直居中打印' },
      printGridlines: { type: 'boolean', description: '打印网格线' },
      printHeadings: { type: 'boolean', description: '打印行号列标' },
      sheet: sheetParam,
    },
  },
};

export const setSheetPageSetupHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<SettingsResponse>(
      'setSheetPageSetup',
      {
        sheet: args.sheet, orientation: args.orientation, paperSize: args.paperSize,
        topMargin: args.topMargin, bottomMargin: args.bottomMargin, leftMargin: args.leftMargin,
        rightMargin: args.rightMargin, headerMargin: args.headerMargin, footerMargin: args.footerMargin,
        zoom: args.zoom, fitToPagesWide: args.fitToPagesWide, fitToPagesTall: args.fitToPagesTall,
        centerHorizontally: args.centerHorizontally, centerVertically: args.centerVertically,
        printGridlines: args.printGridlines, printHeadings: args.printHeadings,
      },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return settingsFail('设置页面失败', response.error);
    return settingsResult('页面设置已更新', response.data);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return settingsFail('设置页面出错', errMsg);
  }
};

/** 打印标题 */
export const setSheetPrintTitlesDefinition: ToolDefinition = {
  name: 'wps_excel_set_sheet_print_titles',
  description: '设置打印时每页重复的行/列（打印标题）：如行 $1:$1 让表头每页都出现，列 $A:$A 让第一列每页都出现。使用场景：多页表格打印出来每一页都有表头。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      printTitleRows: { type: 'string', description: '每页重复的行，如 $1:$1；传空字符串清除' },
      printTitleColumns: { type: 'string', description: '每页重复的列，如 $A:$A；传空字符串清除' },
      sheet: sheetParam,
    },
  },
};

export const setSheetPrintTitlesHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<SettingsResponse>(
      'setSheetPrintTitles',
      { sheet: args.sheet, printTitleRows: args.printTitleRows, printTitleColumns: args.printTitleColumns },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return settingsFail('设置打印标题失败', response.error);
    return settingsResult('打印标题已更新', response.data);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return settingsFail('设置打印标题出错', errMsg);
  }
};

/** 页眉页脚 */
export const setSheetHeaderFooterDefinition: ToolDefinition = {
  name: 'wps_excel_set_sheet_header_footer',
  description: '设置打印页眉页脚。文本里可以用 Excel 的域代码：&P 页码、&N 总页数、&D 日期、&T 时间、&F 文件名、&A 工作表名。使用场景：页脚写「第 &P 页 / 共 &N 页」。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      leftHeader: { type: 'string', description: '页眉左侧文本；传空字符串清除' },
      centerHeader: { type: 'string', description: '页眉中间文本' },
      rightHeader: { type: 'string', description: '页眉右侧文本' },
      leftFooter: { type: 'string', description: '页脚左侧文本' },
      centerFooter: { type: 'string', description: '页脚中间文本，如 第 &P 页 / 共 &N 页' },
      rightFooter: { type: 'string', description: '页脚右侧文本' },
      sheet: sheetParam,
    },
  },
};

export const setSheetHeaderFooterHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<SettingsResponse>(
      'setSheetHeaderFooter',
      {
        sheet: args.sheet, leftHeader: args.leftHeader, centerHeader: args.centerHeader, rightHeader: args.rightHeader,
        leftFooter: args.leftFooter, centerFooter: args.centerFooter, rightFooter: args.rightFooter,
      },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return settingsFail('设置页眉页脚失败', response.error);
    return settingsResult('页眉页脚已更新', response.data);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return settingsFail('设置页眉页脚出错', errMsg);
  }
};
/** 工作表外观：隐藏与标签色 */
export const setSheetAppearanceDefinition: ToolDefinition = {
  name: 'wps_excel_set_sheet_appearance',
  description: '设置工作表的可见性与标签色：可见 / 隐藏 / 深度隐藏（veryHidden，用户界面上无法取消隐藏），以及标签颜色（十六进制如 #FF9900）。使用场景：把中间计算表藏起来、给关键工作表标个颜色。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      visible: { type: 'string', enum: ['visible', 'hidden', 'veryHidden'], description: '可见性，默认不变；veryHidden 在界面上无法恢复，脚本可恢复' },
      tabColor: { type: 'string', description: '标签颜色，如 #FF9900；传空字符串恢复默认' },
      sheet: sheetParam,
    },
  },
};

export const setSheetAppearanceHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<SettingsResponse>(
      'setSheetAppearance',
      { sheet: args.sheet, visible: args.visible, tabColor: args.tabColor },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return settingsFail('设置工作表外观失败', response.error);
    return settingsResult('工作表外观已更新', response.data);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return settingsFail('设置工作表外观出错', errMsg);
  }
};

/** 分级显示 */
export const setOutlineLevelsDefinition: ToolDefinition = {
  name: 'wps_excel_set_outline_levels',
  description: '控制分级显示的展开层级与汇总位置：rowLevels/columnLevels 指定行/列显示到第几级（1 表示全部折叠），summaryRow/summaryColumn 指定汇总行在上还是下、汇总列在左还是右。使用场景：分组之后把明细收起来只留汇总。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      rowLevels: { type: 'number', description: '行显示到第几级（1 表示只显示第 1 级，深层的折叠）' },
      columnLevels: { type: 'number', description: '列显示到第几级' },
      summaryRow: { type: 'string', enum: ['above', 'below'], description: '汇总行在明细的上方还是下方' },
      summaryColumn: { type: 'string', enum: ['left', 'right'], description: '汇总列在明细的左侧还是右侧' },
      sheet: sheetParam,
    },
  },
};

export const setOutlineLevelsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ applied?: string[]; sheet?: string; summaryRow?: number; summaryColumn?: number }>(
      'setOutlineLevels',
      { sheet: args.sheet, rowLevels: args.rowLevels, columnLevels: args.columnLevels, summaryRow: args.summaryRow, summaryColumn: args.summaryColumn },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return settingsFail('设置分级显示失败', response.error);
    const d = response.data || {};
    const applied = d.applied || [];
    const head = applied.length ? '分级显示已更新（' + applied.join(', ') + '）' : '分级显示未变化';
    const where = (d.summaryRow === 0 ? '上' : '下') + ' / ' + (d.summaryColumn === -1 ? '左' : '右');
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: head + '；汇总行/列位置: ' + where }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return settingsFail('设置分级显示出错', errMsg);
  }
};

/** 清除分页符 */
export const resetPageBreaksDefinition: ToolDefinition = {
  name: 'wps_excel_reset_page_breaks',
  description: '清除工作表上的手动分页符，恢复按内容自动分页。使用场景：手工插过分页符之后想回到自动分页。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: { type: 'object', properties: { sheet: sheetParam } },
};

export const resetPageBreaksHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ sheet?: string; hPageBreaks?: number; message?: string }>(
      'resetPageBreaks',
      { sheet: args.sheet },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return settingsFail('清除分页符失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: (response.data?.message || '手动分页符已清除') + '；剩余手动分页符 ' + String(response.data?.hPageBreaks ?? 0) + ' 条' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return settingsFail('清除分页符出错', errMsg);
  }
};

/** 公式审计：引用追踪 */
export const getFormulaAuditDefinition: ToolDefinition = {
  name: 'wps_excel_get_formula_audit',
  description: '审计一个单元格的公式依赖：它引用了谁（precedents）、谁引用了它（dependents）、直接引用几处，并可选在界面上画出追踪箭头。使用场景："这个数是怎么算出来的"、"改这个格子会影响哪些单元格"。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      cell: { type: 'string', description: '要审计的单元格，如 C5' },
      showPrecedents: { type: 'boolean', description: '在界面上画出引用来源箭头' },
      showDependents: { type: 'boolean', description: '在界面上画出被引用箭头' },
      clearArrows: { type: 'boolean', description: '先清除工作表上的所有追踪箭头' },
      sheet: sheetParam,
    },
    required: ['cell'],
  },
};

export const getFormulaAuditHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      sheet?: string; cell?: string; formula?: string; hasFormula?: boolean; precedents?: number; dependents?: number;
      directPrecedents?: number; precedentAddress?: string; dependentAddress?: string;
    }>(
      'getFormulaAudit',
      { sheet: args.sheet, cell: args.cell, showPrecedents: args.showPrecedents, showDependents: args.showDependents, clearArrows: args.clearArrows },
      WpsAppType.SPREADSHEET
    );
    if (!response.success) return settingsFail('公式审计失败', response.error);
    const d = response.data || {};
    const cell = d.cell || String(args.cell);
    const label = d.hasFormula ? '公式' : '不是公式，是常量（当前值';
    const value = d.hasFormula ? (d.formula || '(空)') : ((d.formula || '空') + '）');
    const lines = [cell + ' 的' + label + ': ' + value];
    lines.push('  引用来源（precedents）: ' + String(d.precedents ?? 0) + ' 处' + (d.precedentAddress ? ' → ' + d.precedentAddress : ''));
    lines.push('  直接引用: ' + String(d.directPrecedents ?? 0) + ' 处');
    lines.push('  被引用（dependents）: ' + String(d.dependents ?? 0) + ' 处' + (d.dependentAddress ? ' → ' + d.dependentAddress : ''));
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return settingsFail('公式审计出错', errMsg);
  }
};

export const sheetSettingsTools: RegisteredTool[] = [
  { definition: getSheetSettingsDefinition, handler: getSheetSettingsHandler },
  { definition: setSheetPageSetupDefinition, handler: setSheetPageSetupHandler },
  { definition: setSheetPrintTitlesDefinition, handler: setSheetPrintTitlesHandler },
  { definition: setSheetHeaderFooterDefinition, handler: setSheetHeaderFooterHandler },
  { definition: setSheetAppearanceDefinition, handler: setSheetAppearanceHandler },
  { definition: setOutlineLevelsDefinition, handler: setOutlineLevelsHandler },
  { definition: resetPageBreaksDefinition, handler: resetPageBreaksHandler },
  { definition: getFormulaAuditDefinition, handler: getFormulaAuditHandler },
];

export default sheetSettingsTools;

/**
 * Input: Word 长尾工具的调用参数
 * Output: 内容控件 / 脚注尾注 / 索引 / 交叉引用 / 邮件合并 的执行结果
 * Pos: Word 长尾（P3-4，D2 的最后一块）。支持面已用裸 COM 量过（test/.artifacts/e2e/word-probe*.ps1）。
 *      一旦我被修改，请更新我的头部注释，并在 docs/FIXES.md 记账。
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

const CONTROL_NAMES: Record<number, string> = { 0: '富文本', 1: '纯文本', 2: '复选框', 3: '组合框', 4: '下拉列表', 5: '日期选择器', 7: '图片' };

function longtailFail(prefix: string, error?: string): ToolCallResult {
  return { id: uuidv4(), success: false, content: [{ type: 'text', text: prefix + ': ' + error }], error: error };
}

/** 内容控件（读） */
export const getContentControlsDefinition: ToolDefinition = {
  name: 'wps_word_get_content_controls',
  description: '列出文档里的内容控件（可填写的结构化区域）：序号、类型、标题、标签、当前文本。使用场景：看清模板里有哪些可填字段，配合 wps_word_add_content_control 使用。',
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: {} },
};

export const getContentControlsHandler: ToolHandler = async (_args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ controls?: Array<{ index?: number; type?: number; title?: string; tag?: string; text?: string }>; count?: number }>(
      'getContentControls',
      {},
      WpsAppType.WRITER
    );
    if (!response.success) return longtailFail('读取内容控件失败', response.error);
    const controls = response.data?.controls || [];
    if (!controls.length) return { id: uuidv4(), success: true, content: [{ type: 'text', text: '这个文档里没有内容控件。' }] };
    const lines = controls.map((c) => {
      const kind = CONTROL_NAMES[Number(c.type)] || ('类型 ' + String(c.type));
      const label = c.title || c.tag || '';
      return '  ' + String(c.index) + '. ' + kind + (label ? '「' + label + '」' : '') + '：' + String(c.text || '');
    });
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '内容控件（' + controls.length + ' 个）:\n' + lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return longtailFail('读取内容控件出错', errMsg);
  }
};

/** 内容控件（加） */
export const addContentControlDefinition: ToolDefinition = {
  name: 'wps_word_add_content_control',
  description: '在光标处插入一个内容控件（可填写的结构化区域），可指定类型、标题、标签与初始文本。使用场景：做一份可复用的表单模板。',
  category: ToolCategory.DOCUMENT,
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

export const addContentControlHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ type?: string; title?: string; text?: string; total?: number }>(
      'addContentControl',
      { type: args.type, title: args.title, tag: args.tag, text: args.text },
      WpsAppType.WRITER
    );
    if (!response.success) return longtailFail('插入内容控件失败', response.error);
    const d = response.data || {};
    const kind = CONTROL_NAMES[Number(d.type)] || String(d.type);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已插入' + kind + '内容控件' + (d.title ? '（标题 ' + d.title + '）' : '') + '，当前内容「' + String(d.text || '') + '」；文档里共 ' + String(d.total ?? 0) + ' 个' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return longtailFail('插入内容控件出错', errMsg);
  }
};

/** 脚注 / 尾注（加） */
export const addFootnoteDefinition: ToolDefinition = {
  name: 'wps_word_add_footnote',
  description: '在光标处插入脚注（页面底部的注释）。使用场景：给一个术语加解释。',
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: { text: { type: 'string', description: '脚注内容' } }, required: ['text'] },
};

export const addFootnoteHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ text?: string; total?: number }>(
      'addFootnote',
      { text: args.text },
      WpsAppType.WRITER
    );
    if (!response.success) return longtailFail('插入脚注失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已插入脚注「' + String(response.data?.text || '') + '」；文档里共 ' + String(response.data?.total ?? 0) + ' 条' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return longtailFail('插入脚注出错', errMsg);
  }
};

export const addEndnoteDefinition: ToolDefinition = {
  name: 'wps_word_add_endnote',
  description: '在光标处插入尾注（文档末尾的注释）。使用场景：给引用加出处，统一列在文末。',
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: { text: { type: 'string', description: '尾注内容' } }, required: ['text'] },
};

export const addEndnoteHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ text?: string; total?: number }>(
      'addEndnote',
      { text: args.text },
      WpsAppType.WRITER
    );
    if (!response.success) return longtailFail('插入尾注失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已插入尾注「' + String(response.data?.text || '') + '」；文档里共 ' + String(response.data?.total ?? 0) + ' 条' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return longtailFail('插入尾注出错', errMsg);
  }
};

/** 脚注 / 尾注（读） */
export const getNotesDefinition: ToolDefinition = {
  name: 'wps_word_get_notes',
  description: '列出文档里的脚注与尾注及其内容。使用场景：核对注释有没有漏、内容对不对。',
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: {} },
};

export const getNotesHandler: ToolHandler = async (_args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ footnotes?: Array<{ index?: number; text?: string }>; endnotes?: Array<{ index?: number; text?: string }>; footnoteCount?: number; endnoteCount?: number }>(
      'getNotes',
      {},
      WpsAppType.WRITER
    );
    if (!response.success) return longtailFail('读取注释失败', response.error);
    const d = response.data || {};
    const footnotes = d.footnotes || [];
    const endnotes = d.endnotes || [];
    const lines = ['脚注 ' + String(d.footnoteCount ?? footnotes.length) + ' 条，尾注 ' + String(d.endnoteCount ?? endnotes.length) + ' 条'];
    for (const f of footnotes) lines.push('  脚注 ' + String(f.index) + ': ' + String(f.text || ''));
    for (const e of endnotes) lines.push('  尾注 ' + String(e.index) + ': ' + String(e.text || ''));
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return longtailFail('读取注释出错', errMsg);
  }
};

/** 索引 */
export const insertIndexDefinition: ToolDefinition = {
  name: 'wps_word_insert_index',
  description: '在文档末尾插入索引（按索引项自动生成）。使用场景：长文档要一张术语/条目索引。注意：只有先用 Word 的「标记索引项」标记过内容，索引里才会有条目；没标记过会显示「未找到索引项」。',
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: {} },
};

export const insertIndexHandler: ToolHandler = async (_args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ indexes?: number; message?: string }>(
      'insertIndex',
      {},
      WpsAppType.WRITER
    );
    if (!response.success) return longtailFail('插入索引失败', response.error);
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已插入索引（文档里共 ' + String(response.data?.indexes ?? 0) + ' 个）。' + String(response.data?.message || '') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return longtailFail('插入索引出错', errMsg);
  }
};

/** 交叉引用 */
export const insertCrossReferenceDefinition: ToolDefinition = {
  name: 'wps_word_insert_cross_reference',
  description: '在光标处插入交叉引用（引用标题/书签/脚注等）。参数直接对应 Word 的 InsertCrossReference(ReferenceType, ReferenceKind, ReferenceItem)：referenceType 1=标题 2=书签 3=脚注 4=尾注，referenceKind -1=正文文本，referenceItem 是要引用的编号或名字。这两个枚举在 WPS 上没有被验证过的友好映射，所以照实透传，不做猜测性翻译。',
  category: ToolCategory.DOCUMENT,
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

export const insertCrossReferenceHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ referenceType?: number; referenceKind?: number; referenceItem?: string }>(
      'insertCrossReference',
      { referenceType: args.referenceType, referenceKind: args.referenceKind, referenceItem: args.referenceItem },
      WpsAppType.WRITER
    );
    if (!response.success) return longtailFail('插入交叉引用失败', response.error);
    const d = response.data || {};
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已插入交叉引用（类型 ' + String(d.referenceType) + '，内容 ' + String(d.referenceKind) + '，引用 ' + String(d.referenceItem) + '）' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return longtailFail('插入交叉引用出错', errMsg);
  }
};

/** 邮件合并 */
export const mailMergeDefinition: ToolDefinition = {
  name: 'wps_word_mail_merge',
  description: '邮件合并：以当前文档为母版，接一个数据文件（CSV），把表格里的值填到母版的占位符处；默认合并成一份新文档，给了 splitDir 就每个数据行单独存一份（母版不动）。使用场景：一份通知模板 + 一张名单，每人一份。占位符在模板里写成 ⟨字段名⟩（也认 {{字段名}} / «字段名»）。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      dataFile: { type: 'string', description: '数据文件路径（CSV，UTF-8，第一行是字段名）' },
      fields: { type: 'array', description: '要替换的字段名列表，按顺序填到对应占位符，例如 Name、Dept；不填则用 CSV 表头里的全部字段' },
      splitDir: { type: 'string', description: '给了这个目录就每个数据行单独生成一份 docx 存进去（文件名用 fileNameField，缺省用第一个字段）；不给则合并成一份新文档' },
      fileNameField: { type: 'string', description: 'splitDir 模式下决定文件名的字段名（取该列的值），或一段带 {列名} 占位符的模板，例如 通知书_{部门}_{姓名}_{工号}；缺省用第一个字段' },
      outputPath: { type: 'string', description: '单文档模式下的保存路径（.docx）；不给则新文档留在内存里，由调用方自己另存' },
    },
    required: ['dataFile'],
  },
};

export const mailMergeHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ mergedDocument?: string; sourceDocument?: string; fields?: string[]; fieldsInserted?: number; preview?: string; message?: string; mode?: string; records?: number; outputDir?: string; outputPath?: string; files?: string[]; written?: number }>(
      'mailMerge',
      {
        dataFile: args.dataFile,
        fields: args.fields,
        splitDir: args.splitDir,
        fileNameField: args.fileNameField,
        outputPath: args.outputPath,
      },
      WpsAppType.WRITER
    );
    if (!response.success) return longtailFail('邮件合并失败', response.error);
    const d = response.data || {};
    const lines = [d.message || '已合并'];
    if (d.mode === 'splitByRecord') {
      lines.push('  模式: 每个数据行一份文档（共 ' + String(d.records) + ' 行，已写出 ' + String(d.written) + ' 份）');
      lines.push('  目录: ' + String(d.outputDir || '?'));
      for (const f of (d.files || []).slice(0, 30)) lines.push('    ' + f);
      lines.push('  母版 ' + String(d.sourceDocument || '?') + ' 未改动');
    } else {
      const out = d.outputPath ? '（已存到 ' + String(d.outputPath) + '）' : '（未存盘）';
      lines.push('  新文档: ' + String(d.mergedDocument || '?') + out + '；母版 ' + String(d.sourceDocument || '?') + ' 未改动');
      lines.push('  数据行: ' + String(d.records) + '，替换字段: ' + ((d.fields || []).join(', ') || '无'));
    }
    if (d.preview) lines.push('  开头: ' + String(d.preview).replace(/[\r\n]+/g, ' / '));
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return longtailFail('邮件合并出错', errMsg);
  }
};

export const wordLongTailTools: RegisteredTool[] = [
  { definition: getContentControlsDefinition, handler: getContentControlsHandler },
  { definition: addContentControlDefinition, handler: addContentControlHandler },
  { definition: addFootnoteDefinition, handler: addFootnoteHandler },
  { definition: addEndnoteDefinition, handler: addEndnoteHandler },
  { definition: getNotesDefinition, handler: getNotesHandler },
  { definition: insertIndexDefinition, handler: insertIndexHandler },
  { definition: insertCrossReferenceDefinition, handler: insertCrossReferenceHandler },
  { definition: mailMergeDefinition, handler: mailMergeHandler },
];

export default wordLongTailTools;

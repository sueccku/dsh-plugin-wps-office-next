/**
 * Input: Word 文档生产族工具的调用参数
 * Output: 页码 / 分栏 / 修订 / 批注删除 的执行结果
 * Pos: Word 文档生产族（P3-3）。支持面已用裸 COM 量过（见 test/.artifacts/e2e/word-probe*.ps1）。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P3 状态。
 */
import { v4 as uuidv4 } from 'uuid';
import { impactText, type RangeImpact } from '../impact';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

const sectionParam = { type: 'number' as const, description: '第几节（从 1 开始）；不填则用第 1 节' };

/** 修订类型（Word 的 WdRevisionType），只翻译常见的几种，其余如实给数字 */
const REVISION_NAMES: Record<number, string> = {
  1: '插入', 2: '删除', 3: '属性', 4: '段落编号', 5: '域显示', 7: '冲突', 8: '样式',
  9: '替换', 10: '段落属性', 11: '表格属性', 12: '节属性', 16: '单元格插入', 17: '单元格删除', 18: '单元格合并',
};

function produceFail(prefix: string, error?: string): ToolCallResult {
  return { id: uuidv4(), success: false, content: [{ type: 'text', text: prefix + ': ' + error }], error: error };
}

/** 页码 */
export const insertPageNumbersDefinition: ToolDefinition = {
  name: 'wps_word_insert_page_numbers',
  description: `给**某一节**的页眉或页脚插入页码。

使用场景：
- "页脚右下角要有页码" —— 页码是域，页数变化会自动更新
- "目录那节用罗马数字、技术方案那节从 1 重新开始" —— numberStyle + restartSection
- "第 X 页 / 共 Y 页" —— includeTotal 打开后再插一个 NUMPAGES 域

注意：
- **只影响点名的那一节**；多节文档要逐节调用。
- numberStyle 支持 arabic / upperRoman / lowerRoman / upperLetter / lowerLetter。
- restartSection=1 表示"本节从 1 重新开始"，0 表示"接着上一节数"。**起始编号写不进去**
  （WPS 的 PageNumbers.StartingNumber 是只读的，写 1 回读仍是 0），所以只能"从 1 重开"，
  不能"从第 7 页起"——请求这种会被明确记进 warnings，不会假装成功。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      section: sectionParam,
      position: { type: 'string', enum: ['footer', 'header'], description: '放页脚还是页眉，默认 footer' },
      alignment: { type: 'string', enum: ['left', 'center', 'right'], description: '对齐方式，默认 right' },
      showFirstPage: { type: 'boolean', description: '是否在首页显示，默认 true' },
      includeTotal: { type: 'boolean', description: '是否同时插入总页数（NUMPAGES 域），形成"第 X 页 / 共 Y 页"，默认 false' },
      numberStyle: {
        type: 'string',
        enum: ['arabic', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter'],
        description: '页码编号格式，默认不动（arabic / upperRoman / lowerRoman / upperLetter / lowerLetter）',
      },
      restartSection: {
        type: 'boolean',
        description: 'true = 本节页码从 1 重新开始；false = 接着上一节数（续前节）。默认不动',
      },
    },
  },
};

export const insertPageNumbersHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      section?: number;
      position?: string;
      alignment?: string;
      pageNumbers?: number;
      includeTotal?: boolean;
      totalPagesFields?: number;
      numberStyle?: string | null;
      restartSection?: boolean | null;
    }>(
      'insertPageNumbers',
      { section: args.section, position: args.position, alignment: args.alignment, showFirstPage: args.showFirstPage, includeTotal: args.includeTotal, numberStyle: args.numberStyle, restartSection: args.restartSection },
      WpsAppType.WRITER
    );
    if (!response.success) return produceFail('插入页码失败', response.error);
    const d = response.data || {};
    const where = d.position === 'header' ? '页眉' : '页脚';
    const align = d.alignment === 'center' ? '居中' : d.alignment === 'left' ? '左对齐' : '右对齐';
    const total = d.includeTotal ? '，并插入了总页数域 ' + String(d.totalPagesFields ?? 0) + ' 个（NUMPAGES）' : '';
    const styleCn: Record<string, string> = {
      arabic: '阿拉伯数字', upperRoman: '大写罗马数字', lowerRoman: '小写罗马数字',
      upperLetter: '大写字母', lowerLetter: '小写字母',
    };
    const style = d.numberStyle ? '，编号格式 ' + (styleCn[d.numberStyle] || d.numberStyle) : '';
    const restart = d.restartSection === true ? '，本节从 1 重新开始'
      : d.restartSection === false ? '，接着上一节数（续前节）' : '';
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '第 ' + String(d.section) + ' 节' + where + '已插入页码（' + align + style + restart + '，该处共 ' + String(d.pageNumbers ?? 0) + ' 个页码域' + total + '）' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return produceFail('插入页码出错', errMsg);
  }
};

/** 读页眉页脚（含域） */
export const getHeaderFooterDefinition: ToolDefinition = {
  name: 'wps_word_get_header_footer',
  description: '读取某一节的页眉与页脚：文本内容与其中的域（页码、总页数等）。使用场景：改页眉页脚之前先看现在写的是什么；核对页码域到底进文件没有。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      section: sectionParam,
    },
  },
};

export const getHeaderFooterHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      section?: number;
      headerText?: string;
      footerText?: string;
      headerFields?: Array<{ type: string }>;
      footerFields?: Array<{ type: string }>;
    }>('getHeaderFooter', { section: args.section }, WpsAppType.WRITER);
    if (!response.success) return produceFail('读取页眉页脚失败', response.error);
    const d = response.data || {};
    const names = (list?: Array<{ type: string }>) =>
      (list || []).length ? (list || []).map((f) => f.type).join(', ') : '无';
    return {
      id: uuidv4(),
      success: true,
      content: [{
        type: 'text',
        text: '第 ' + String(d.section) + ' 节：\n页眉: ' + JSON.stringify(d.headerText ?? '') + '\n页眉域: ' + names(d.headerFields)
          + '\n页脚: ' + JSON.stringify(d.footerText ?? '') + '\n页脚域: ' + names(d.footerFields),
      }],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return produceFail('读取页眉页脚出错', errMsg);
  }
};

/** 分栏 */
export const setColumnsDefinition: ToolDefinition = {
  name: 'wps_word_set_columns',
  description: '设置分栏：栏数、栏间距、是否加分隔线。使用场景：把长文排成两栏。count 用 1 就是取消分栏。',
  category: ToolCategory.DOCUMENT,
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

export const setColumnsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ section?: number; applied?: string[]; count?: number; spacing?: number }>(
      'setColumns',
      { section: args.section, count: args.count, spacing: args.spacing, lineBetween: args.lineBetween },
      WpsAppType.WRITER
    );
    if (!response.success) return produceFail('设置分栏失败', response.error);
    const d = response.data || {};
    const applied = d.applied || [];
    // 单栏时 WPS 的 TextColumns.Spacing 是哨兵值 9999999，印出来像 bug，所以只在分栏时报告间距。
    const spacing = Number(d.count) > 1 ? '，栏间距 ' + String(d.spacing) + ' 磅' : '';
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '第 ' + String(d.section) + ' 节现在是 ' + String(d.count) + ' 栏（已应用 ' + applied.join(', ') + spacing + '）' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return produceFail('设置分栏出错', errMsg);
  }
};

/** 修订列表 */
export const getRevisionsDefinition: ToolDefinition = {
  name: 'wps_word_get_revisions',
  description: '列出文档里的修订（插入/删除/替换等），并报告「修订跟踪」当前是否打开。使用场景：接手别人的稿子，先看改了什么。',
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: {} },
};

export const getRevisionsHandler: ToolHandler = async (_args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ revisions?: Array<{ index?: number; type?: number; text?: string; author?: string }>; count?: number; trackChanges?: boolean }>(
      'getRevisions',
      {},
      WpsAppType.WRITER
    );
    if (!response.success) return produceFail('读取修订失败', response.error);
    const d = response.data || {};
    const revisions = d.revisions || [];
    const head = '修订跟踪: ' + (d.trackChanges ? '已打开' : '已关闭');
    if (!revisions.length) return { id: uuidv4(), success: true, content: [{ type: 'text', text: head + '；文档里没有修订。' }] };
    const lines = revisions.map((r) => '  ' + String(r.index) + '. ' + (REVISION_NAMES[Number(r.type)] || ('类型 ' + String(r.type))) + '：' + String(r.text || '') + (r.author ? '（' + r.author + '）' : ''));
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: head + '；修订（' + revisions.length + ' 处）:\n' + lines.join('\n') }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return produceFail('读取修订出错', errMsg);
  }
};

/** 接受修订 */
export const acceptRevisionsDefinition: ToolDefinition = {
  name: 'wps_word_accept_revisions',
  description: '接受修订：给 index 只接受那一处，不填则接受全部。使用场景：审阅通过，把改动定稿。接受后文字变成正文，修订记录消失。',
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: { index: { type: 'number', description: '只接受第几处修订（从 1 开始）；不填则全部接受' } } },
};

export const acceptRevisionsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ accepted?: number; remaining?: number }>(
      'acceptRevisions',
      { index: args.index },
      WpsAppType.WRITER
    );
    if (!response.success) return produceFail('接受修订失败', response.error);
    const d = response.data || {};
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已接受 ' + String(d.accepted ?? 0) + ' 处修订；还剩 ' + String(d.remaining ?? 0) + ' 处' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return produceFail('接受修订出错', errMsg);
  }
};

/** 拒绝修订 */
export const rejectRevisionsDefinition: ToolDefinition = {
  name: 'wps_word_reject_revisions',
  description: '拒绝修订：给 index 只拒绝那一处，不填则拒绝全部（回到改之前的原文）。使用场景：这版改动不要，恢复原样。',
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: { index: { type: 'number', description: '只拒绝第几处修订（从 1 开始）；不填则全部拒绝' } } },
};

export const rejectRevisionsHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ rejected?: number; remaining?: number }>(
      'rejectRevisions',
      { index: args.index },
      WpsAppType.WRITER
    );
    if (!response.success) return produceFail('拒绝修订失败', response.error);
    const d = response.data || {};
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已拒绝 ' + String(d.rejected ?? 0) + ' 处修订；还剩 ' + String(d.remaining ?? 0) + ' 处' }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return produceFail('拒绝修订出错', errMsg);
  }
};

/** 删除批注 */
export const deleteCommentDefinition: ToolDefinition = {
  name: 'wps_word_delete_comment',
  description: '删除批注：给 index 只删那一条（序号见 wps_word_get_comments），不填则全部删除。使用场景：意见处理完了，清掉批注再交付。',
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: { index: { type: 'number', description: '只删第几条批注（从 1 开始）；不填则全部删除' } } },
};

export const deleteCommentHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ deleted?: number; remaining?: number; impact?: RangeImpact }>(
      'deleteComment',
      { index: args.index },
      WpsAppType.WRITER
    );
    if (!response.success) return produceFail('删除批注失败', response.error);
    const d = response.data || {};
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: '已删除 ' + String(d.deleted ?? 0) + ' 条批注；还剩 ' + String(d.remaining ?? 0) + ' 条' + impactText(d.impact) }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return produceFail('删除批注出错', errMsg);
  }
};

export const wordProduceTools: RegisteredTool[] = [
  { definition: insertPageNumbersDefinition, handler: insertPageNumbersHandler },
  { definition: getHeaderFooterDefinition, handler: getHeaderFooterHandler },
  { definition: setColumnsDefinition, handler: setColumnsHandler },
  { definition: getRevisionsDefinition, handler: getRevisionsHandler },
  { definition: acceptRevisionsDefinition, handler: acceptRevisionsHandler },
  { definition: rejectRevisionsDefinition, handler: rejectRevisionsHandler },
  { definition: deleteCommentDefinition, handler: deleteCommentHandler },
];

export default wordProduceTools;

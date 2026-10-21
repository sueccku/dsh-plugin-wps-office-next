/**
 * Input: 文档管理工具参数
 * Output: 文档操作结果
 * Pos: Word 文档管理工具实现。一旦我被修改，请更新我的头部注释，并在 docs/FIXES.md 记账。
 * Word文档管理Tools
 * 处理文档的打开、切换、获取列表等管理操作
 *
 * 包含（10 个）：
 * - wps_word_get_open_documents: 获取所有已打开的文档列表
 * - wps_word_switch_document: 切换到指定文档
 * - wps_word_open_document: 打开指定路径的文档
 * - wps_word_get_document_text: 获取文档文本内容
 * - wps_word_insert_header: 设置页眉内容
 * - wps_word_insert_footer: 设置页脚内容
 * - wps_word_insert_section_break: 插入分节符
 * - wps_word_set_line_spacing: 设置行距
 * - wps_word_create_document: 新建空白文档
 * - wps_word_close_document: 关闭文档（可选保存）
 *
 * 目录生成在 format.ts（generate_toc）。
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

/**
 * 获取所有已打开的文档列表
 */
export const getOpenDocumentsDefinition: ToolDefinition = {
  name: 'wps_word_get_open_documents',
  description: `获取当前WPS Writer中所有已打开的文档列表。

使用场景：
- "看看现在打开了哪些文档"
- "列出所有打开的Word文件"
- "查看当前文档列表"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const getOpenDocumentsHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      documents: Array<{ name: string; path: string; active: boolean }>;
    }>(
      'getOpenDocuments',
      {},
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const docs = response.data.documents;
      if (!docs || docs.length === 0) {
        return {
          id: uuidv4(),
          success: true,
          content: [{ type: 'text', text: '当前没有打开任何文档' }],
        };
      }

      const docList = docs
        .map((doc, i) => `${i + 1}. ${doc.name}${doc.active ? ' (当前活动)' : ''}\n   路径: ${doc.path}`)
        .join('\n');

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已打开的文档列表 (共${docs.length}个):\n${docList}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取文档列表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档列表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 切换到指定文档
 */
export const switchDocumentDefinition: ToolDefinition = {
  name: 'wps_word_switch_document',
  description: `切换到指定名称的文档。

使用场景：
- "切换到报告.docx"
- "打开另一个文档窗口"
- "切换到那个合同文档"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要切换到的文档名称',
      },
    },
    required: ['name'],
  },
};

export const switchDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as { name: string };

  if (!name || name.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '文档名称不能为空！' }],
      error: '文档名称为空',
    };
  }

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'switchDocument',
      { name },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已切换到文档: ${name}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换文档失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换文档出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 打开指定路径的文档
 */
export const openDocumentDefinition: ToolDefinition = {
  name: 'wps_word_open_document',
  description: `打开指定路径的Word文档。默认以只读打开（readOnly=true），防止读来源文档时误改原件；要编辑时显式传 readOnly:false。

使用场景：
- "打开桌面上的报告.docx"
- "帮我打开这个文件路径的文档"
- "加载指定位置的Word文件"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      path: {
        type: 'string',
        description: '要打开的文档文件路径',
      },
      readOnly: {
        type: 'boolean',
        description: '以只读方式打开（默认 true，安全默认）；要保存改动必须显式传 false。',
      },
    },
    required: ['path'],
  },
};

export const openDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { path, readOnly } = args as { path: string; readOnly?: boolean };

  if (!path || path.trim() === '') {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '文件路径不能为空！' }],
      error: '文件路径为空',
    };
  }

  try {
    const params = {
      path: path,
      readOnly: readOnly === undefined ? true : readOnly,
    };

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      documentName: string;
    }>(
      'openDocument',
      params,
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文档打开成功！\n文件: ${response.data.documentName || path}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `打开文档失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `打开文档出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 获取文档文本内容
 */
export const getDocumentTextDefinition: ToolDefinition = {
  name: 'wps_word_get_document_text',
  description: `获取当前Word文档的文本内容。

使用场景：
- "读取文档内容"
- "获取文档的前100个字符"
- "查看文档从第50到第200个字符的内容"

可指定起始和结束位置来获取部分文本（0 基字符索引）；默认只回前 10,000 字符，超出会标 truncated。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      start: {
        type: 'number',
        description: '起始位置（字符索引），默认从头开始',
      },
      end: {
        type: 'number',
        description: '结束位置（字符索引），默认到文档末尾',
      },
    },
  },
};

export const getDocumentTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { start, end } = args as { start?: number; end?: number };

  try {
    const params: Record<string, unknown> = {};
    if (start !== undefined) params.start = start;
    if (end !== undefined) params.end = end;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      text: string;
      length: number;
      totalLength?: number;
      start?: number;
      end?: number;
      truncated?: boolean;
    }>(
      'getDocumentText',
      params,
      WpsAppType.WRITER
    );

    if (response.success && response.data) {
      const d = response.data;
      const rangeInfo = start !== undefined || end !== undefined
        ? `\n范围: ${d.start ?? start ?? 0} - ${d.end ?? end ?? '末尾'}`
        : '';
      const truncInfo = d.truncated ? `\n注意: 只返回前 ${d.length} 字符（全文 ${d.totalLength ?? '?'} 字符），请用 start/end 分页读取。` : '';

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文档文本内容 (${d.length}字符${d.totalLength !== undefined ? `，全文 ${d.totalLength} 字符` : ''})${rangeInfo}${truncInfo}:\n\n${d.text}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取文档文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取文档文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置页眉内容
 */
export const insertHeaderDefinition: ToolDefinition = {
  name: 'wps_word_insert_header',
  description: `设置页眉内容（传空 text 可清空）。

使用场景：
- "给文档加个页眉"
- "设置页眉为公司名称"
- "修改第2节的页眉"
- "删掉页眉"（text:""）`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '页眉文本内容；传空字符串表示清空页眉',
      },
      section: {
        type: 'number',
        description: '节编号（从1开始），默认1',
        default: 1,
      },
      alignment: {
        type: 'string',
        enum: ['left', 'center', 'right'],
        description: '页眉对齐方式；不传则保持原样',
      },
      firstPageDifferent: { type: 'boolean', description: '首页页眉是否不同' },
      oddEvenDifferent: { type: 'boolean', description: '奇偶页页眉是否不同' },
      linkToPrevious: { type: 'boolean', description: '是否链接到前一条页眉（false=断开）' },
    },
    required: [],
  },
};

export const insertHeaderHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, section = 1, alignment, firstPageDifferent, oddEvenDifferent, linkToPrevious } = args as {
    text?: string; section?: number; alignment?: string;
    firstPageDifferent?: boolean; oddEvenDifferent?: boolean; linkToPrevious?: boolean;
  };

  try {
    const params: Record<string, unknown> = { text: text ?? '', section };
    if (alignment !== undefined) params.alignment = alignment;
    if (firstPageDifferent !== undefined) params.firstPageDifferent = firstPageDifferent;
    if (oddEvenDifferent !== undefined) params.oddEvenDifferent = oddEvenDifferent;
    if (linkToPrevious !== undefined) params.linkToPrevious = linkToPrevious;
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertHeader',
      params,
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: text ? `页眉已设置: "${text}" (第${section}节)` : `页眉已清空 (第${section}节)` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页眉失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页眉出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置页脚内容
 */
export const insertFooterDefinition: ToolDefinition = {
  name: 'wps_word_insert_footer',
  description: `设置页脚内容（传空 text 可清空；includePageNumber 可同时插入页码域）。

使用场景：
- "给文档加个页脚"
- "设置页脚为页码"
- "修改第1节的页脚"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '页脚文本内容；传空字符串表示清空页脚',
      },
      section: {
        type: 'number',
        description: '节编号（从1开始），默认1',
        default: 1,
      },
      alignment: {
        type: 'string',
        enum: ['left', 'center', 'right'],
        description: '页脚对齐方式；不传则保持原样',
      },
      includePageNumber: {
        type: 'boolean',
        description: '是否在页脚追加 PAGE 域（第 X 页）',
      },
      firstPageDifferent: { type: 'boolean', description: '首页页脚是否不同' },
      oddEvenDifferent: { type: 'boolean', description: '奇偶页页脚是否不同' },
      linkToPrevious: { type: 'boolean', description: '是否链接到前一条页脚（false=断开）' },
    },
    required: [],
  },
};

export const insertFooterHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, section = 1, alignment, includePageNumber, firstPageDifferent, oddEvenDifferent, linkToPrevious } = args as {
    text?: string;
    section?: number;
    alignment?: string;
    includePageNumber?: boolean;
    firstPageDifferent?: boolean;
    oddEvenDifferent?: boolean;
    linkToPrevious?: boolean;
  };

  try {
    const params: Record<string, unknown> = { text: text ?? '', section };
    if (alignment !== undefined) params.alignment = alignment;
    if (includePageNumber !== undefined) params.includePageNumber = includePageNumber;
    if (firstPageDifferent !== undefined) params.firstPageDifferent = firstPageDifferent;
    if (oddEvenDifferent !== undefined) params.oddEvenDifferent = oddEvenDifferent;
    if (linkToPrevious !== undefined) params.linkToPrevious = linkToPrevious;
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertFooter',
      params,
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: text ? `页脚已设置: "${text}" (第${section}节)` : `页脚已清空 (第${section}节)` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页脚失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页脚出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};



/**
 * 插入分节符
 */
export const insertSectionBreakDefinition: ToolDefinition = {
  name: 'wps_word_insert_section_break',
  description: `插入分节符，用于将文档分为不同的节，以便对各节应用不同的页面设置。

使用场景：
- "插入一个分节符"
- "从下一页开始新的一节"
- "在这里分节"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      breakType: {
        type: 'string',
        description: '分节符类型：nextPage(下一页)、continuous(连续)、evenPage(偶数页)、oddPage(奇数页)，默认nextPage',
        default: 'nextPage',
      },
    },
  },
};

export const insertSectionBreakHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { breakType = 'nextPage' } = args as { breakType?: string };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertSectionBreak',
      { breakType },
      WpsAppType.WRITER
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `分节符已插入（类型: ${breakType}）` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入分节符失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入分节符出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置行距
 */
export const setLineSpacingDefinition: ToolDefinition = {
  name: 'wps_word_set_line_spacing',
  description: `设置段落行距。

使用场景：
- "把行距设为1.5倍"
- "设置第3段的行距为2倍"
- "调整行距"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      lineSpacing: {
        type: 'number',
        description: '行距值（如1.0、1.5、2.0等）',
      },
      paragraphIndex: {
        type: 'number',
        description: '段落索引（从0开始），不指定则应用于当前段落或全文',
      },
    },
    required: ['lineSpacing'],
  },
};

export const setLineSpacingHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { lineSpacing, paragraphIndex } = args as { lineSpacing: number; paragraphIndex?: number };

  if (lineSpacing === undefined || lineSpacing <= 0) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '行距值必须为正数！' }],
      error: '行距值无效',
    };
  }

  try {
    const params: Record<string, unknown> = { lineSpacing };
    if (paragraphIndex !== undefined) params.paragraphIndex = paragraphIndex;

    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setLineSpacing',
      params,
      WpsAppType.WRITER
    );

    if (response.success) {
      const target = paragraphIndex !== undefined ? `第${paragraphIndex}段` : '当前段落';
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `行距已设置为 ${lineSpacing} 倍（${target}）` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置行距失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置行距出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 新建空白文档
 */
export const createDocumentDefinition: ToolDefinition = {
  name: 'wps_word_create_document',
  description: `新建一个空白 Word 文档（不是打开已有文件）。

使用场景：
- "新建一个 Word 文档"
- "把结论写进一个新文档里"
- 需要从零起草，而不是编辑现有文档时

新建后用 wps_word_insert_text 写内容，用 wps_common_save_as 保存到磁盘。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const createDocumentHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{ name?: string }>(
      'createDocument',
      {},
      WpsAppType.WRITER
    );
    if (response.success) {
      const name = response.data?.name;
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `已新建空白文档${name ? ': ' + name : ''}` }],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `新建文档失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `新建文档出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 关闭文档
 */
export const closeDocumentDefinition: ToolDefinition = {
  name: 'wps_word_close_document',
  description: `关闭 Word 文档。默认不保存（save=false），避免批量收尾时回写源文档；要保存改动必须显式传 save:true。

使用场景：
- "关掉这个文档，别留着"
- 一批任务收尾时清理打开的文档

从未保存到磁盘的文档不会被强制保存（不会弹出保存对话框），此时结果里会带 warning 说明。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      name: { type: 'string', description: '要关闭的文档名称；不填则关闭当前活动文档' },
      save: { type: 'boolean', description: '是否保存后关闭，默认 false；要保存改动必须显式传 true' },
    },
  },
};

export const closeDocumentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, save } = args as { name?: string; save?: boolean };
  try {
    const response = await wpsClient.executeMethod<{
      closed: string;
      saved: boolean;
      saveRequested?: boolean;
      warning?: string;
    }>('closeDocument', { name, save: save === true }, WpsAppType.WRITER);
    if (response.success && response.data) {
      const d = response.data;
      const lines = [`文档已关闭: ${d.closed}（${d.saved ? '已保存' : '未保存'}）`];
      if (d.warning) lines.push('注意: ' + d.warning);
      return { id: uuidv4(), success: true, content: [{ type: 'text', text: lines.join('\n') }] };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `关闭文档失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `关闭文档出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 导出所有文档管理相关的Tools
 */
export const documentTools: RegisteredTool[] = [
  { definition: getOpenDocumentsDefinition, handler: getOpenDocumentsHandler },
  { definition: switchDocumentDefinition, handler: switchDocumentHandler },
  { definition: openDocumentDefinition, handler: openDocumentHandler },
  { definition: createDocumentDefinition, handler: createDocumentHandler },
  { definition: closeDocumentDefinition, handler: closeDocumentHandler },
  { definition: getDocumentTextDefinition, handler: getDocumentTextHandler },
  { definition: insertHeaderDefinition, handler: insertHeaderHandler },
  { definition: insertFooterDefinition, handler: insertFooterHandler },
  { definition: insertSectionBreakDefinition, handler: insertSectionBreakHandler },
  { definition: setLineSpacingDefinition, handler: setLineSpacingHandler },
];

export default documentTools;

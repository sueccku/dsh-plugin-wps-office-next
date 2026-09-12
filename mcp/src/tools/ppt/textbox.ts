/**
 * Input: 文本框与标题工具参数
 * Output: 文本框操作结果
 * Pos: PPT 文本框与标题工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
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
 * 删除文本框
 */
export const deleteTextboxDefinition: ToolDefinition = {
  name: 'wps_ppt_delete_textbox',
  description: '删除幻灯片上指定的文本框。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      textboxIndex: { type: 'number', description: '文本框索引' },
    },
    required: ['slideIndex', 'textboxIndex'],
  },
};

export const deleteTextboxHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, textboxIndex } = args as { slideIndex: number; textboxIndex: number };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'deleteTextBox',
      { slideIndex, textboxIndex },
      WpsAppType.PRESENTATION
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除文本框失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `第${slideIndex}页的文本框${textboxIndex}已删除` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `删除文本框出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 获取文本框列表
 */
export const getTextboxesDefinition: ToolDefinition = {
  name: 'wps_ppt_get_textboxes',
  description: '获取幻灯片上所有文本框的列表。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
    },
    required: ['slideIndex'],
  },
};

export const getTextboxesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };
  try {
    const response = await wpsClient.executeMethod<{ textboxes: unknown[] }>(
      'getTextBoxes',
      { slideIndex },
      WpsAppType.PRESENTATION
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取文本框列表失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `第${slideIndex}页文本框:\n${JSON.stringify(response.data?.textboxes || [], null, 2)}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取文本框列表出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 设置文本框文本
 */
export const setTextboxTextDefinition: ToolDefinition = {
  name: 'wps_ppt_set_textbox_text',
  description: '设置指定文本框的文本内容。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      textboxIndex: { type: 'number', description: '文本框索引' },
      text: { type: 'string', description: '文本内容' },
    },
    required: ['slideIndex', 'textboxIndex', 'text'],
  },
};

export const setTextboxTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, textboxIndex, text } = args as {
    slideIndex: number; textboxIndex: number; text: string;
  };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'setTextBoxText',
      { slideIndex, textboxIndex, text },
      WpsAppType.PRESENTATION
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置文本框内容失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `文本框内容已更新` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置文本框内容出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 设置文本框样式
 */
export const setTextboxStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_textbox_style',
  description: '设置文本框样式（字体大小、颜色、粗体等）。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      textboxIndex: { type: 'number', description: '文本框索引' },
      style: {
        type: 'object',
        description: '样式对象，包含 fontSize/fontColor/bold/italic/align 等',
      },
    },
    required: ['slideIndex', 'textboxIndex', 'style'],
  },
};

export const setTextboxStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, textboxIndex, style } = args as {
    slideIndex: number; textboxIndex: number; style: Record<string, unknown>;
  };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'setTextBoxStyle',
      { slideIndex, textboxIndex, style },
      WpsAppType.PRESENTATION
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置文本框样式失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `文本框样式已更新` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置文本框样式出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 获取幻灯片标题
 */
export const getSlideTitleDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_title',
  description: '获取指定幻灯片的标题。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
    },
    required: ['slideIndex'],
  },
};

export const getSlideTitleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };
  try {
    const response = await wpsClient.executeMethod<{ title: string }>(
      'getSlideTitle',
      { slideIndex },
      WpsAppType.PRESENTATION
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取标题失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `第${slideIndex}页标题: ${response.data?.title || '无标题'}` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `获取标题出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 设置幻灯片副标题
 */
export const setSlideSubtitleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_subtitle',
  description: '设置幻灯片的副标题。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      subtitle: { type: 'string', description: '副标题内容' },
    },
    required: ['slideIndex', 'subtitle'],
  },
};

export const setSlideSubtitleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, subtitle } = args as { slideIndex: number; subtitle: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'setSlideSubtitle',
      { slideIndex, subtitle },
      WpsAppType.PRESENTATION
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置副标题失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `第${slideIndex}页副标题已设置` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置副标题出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 设置幻灯片正文内容
 */
export const setSlideContentDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_content',
  description: '设置幻灯片的正文内容区域。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      content: { type: 'string', description: '正文内容' },
    },
    required: ['slideIndex', 'content'],
  },
};

export const setSlideContentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, content } = args as { slideIndex: number; content: string };
  try {
    const response = await wpsClient.executeMethod<{ message: string }>(
      'setSlideContent',
      { slideIndex, content },
      WpsAppType.PRESENTATION
    );
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置正文内容失败: ${response.error}` }], error: response.error };
    }
    return { id: uuidv4(), success: true, content: [{ type: 'text', text: `第${slideIndex}页正文内容已设置` }] };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `设置正文内容出错: ${errMsg}` }], error: errMsg };
  }
};

/**
 * 导出所有文本框与标题相关的Tools
 */
export const textboxTools: RegisteredTool[] = [
  { definition: deleteTextboxDefinition, handler: deleteTextboxHandler },
  { definition: getTextboxesDefinition, handler: getTextboxesHandler },
  { definition: setTextboxTextDefinition, handler: setTextboxTextHandler },
  { definition: setTextboxStyleDefinition, handler: setTextboxStyleHandler },
  { definition: getSlideTitleDefinition, handler: getSlideTitleHandler },
  { definition: setSlideSubtitleDefinition, handler: setSlideSubtitleHandler },
  { definition: setSlideContentDefinition, handler: setSlideContentHandler },
];

export default textboxTools;

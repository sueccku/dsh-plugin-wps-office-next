/**
 * Input: 演示文稿管理工具参数
 * Output: 演示文稿操作结果
 * Pos: PPT 演示文稿管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 包含：
 * - wps_ppt_create_presentation: 新建空白演示文稿
 * - wps_ppt_open_presentation: 打开指定路径的演示文稿
 * - wps_ppt_close_presentation: 关闭演示文稿
 * - wps_ppt_get_open_presentations: 获取所有已打开的演示文稿列表
 * - wps_ppt_switch_presentation: 切换到指定演示文稿
 * - wps_ppt_set_slide_theme: 设置演示文稿主题
 * - wps_ppt_copy_slide: 复制幻灯片
 * - wps_ppt_insert_slide_image: 在幻灯片中插入图片
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient, setPptTarget } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

/**
 * 新建空白演示文稿
 */
export const createPresentationDefinition: ToolDefinition = {
  name: 'wps_ppt_create_presentation',
  description: `新建空白演示文稿。

使用场景：
- "新建一个PPT"
- "创建一个演示文稿"
- "打开一个新的PPT"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const createPresentationHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
    }>(
      'createPresentation',
      {},
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `新建演示文稿成功！\n名称: ${response.data.name}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `新建演示文稿失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `新建演示文稿出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 打开指定路径的演示文稿
 */
export const openPresentationDefinition: ToolDefinition = {
  name: 'wps_ppt_open_presentation',
  description: `打开指定路径的演示文稿文件。

支持的文件格式：
- .pptx: PowerPoint 演示文稿
- .ppt: 旧版 PowerPoint 格式
- .dps: WPS 演示格式

使用场景：
- "打开桌面上的演示文稿"
- "打开这个PPT文件"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      filePath: {
        type: 'string',
        description: '演示文稿文件的完整路径',
      },
    },
    required: ['filePath'],
  },
};

export const openPresentationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { filePath } = args as {
    filePath: string;
  };

  try {
    // 跨平台参数对齐：macOS/Windows 底层均读取 params.path，需同时发送 path 别名
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
      filePath: string;
    }>(
      'openPresentation',
      { filePath, path: filePath },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `演示文稿打开成功！\n名称: ${response.data.name}\n路径: ${response.data.filePath}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `打开演示文稿失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `打开演示文稿出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 关闭演示文稿
 */
export const closePresentationDefinition: ToolDefinition = {
  name: 'wps_ppt_close_presentation',
  description: `关闭演示文稿。可指定文稿名称，不指定则关闭当前活动文稿。

使用场景：
- "关闭这个PPT"
- "关闭演示文稿不保存"
- "关闭指定的PPT"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要关闭的演示文稿名称，不填则关闭当前活动文稿',
      },
      save: {
        type: 'boolean',
        description: '关闭前是否保存，默认true',
      },
    },
    required: [],
  },
};

export const closePresentationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, save } = args as {
    name?: string;
    save?: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
    }>(
      'closePresentation',
      {
        name,
        save: save !== false,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const saveStatus = save !== false ? '已保存' : '未保存';
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `演示文稿已关闭！\n名称: ${response.data.name}\n保存状态: ${saveStatus}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `关闭演示文稿失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `关闭演示文稿出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 获取所有已打开的演示文稿列表
 */
export const getOpenPresentationsDefinition: ToolDefinition = {
  name: 'wps_ppt_get_open_presentations',
  description: `获取当前所有已打开的演示文稿列表。

使用场景：
- "有哪些PPT打开着"
- "列出所有演示文稿"
- "查看打开的PPT"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getOpenPresentationsHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      presentations: Array<{
        name: string;
        path: string;
        slideCount: number;
        isActive: boolean;
      }>;
      count: number;
    }>(
      'getOpenPresentations',
      {},
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const { presentations, count } = response.data;
      if (count === 0) {
        return {
          id: uuidv4(),
          success: true,
          content: [{ type: 'text', text: '当前没有打开的演示文稿。' }],
        };
      }

      let output = `当前打开 ${count} 个演示文稿：\n\n`;
      presentations.forEach((p, i) => {
        const activeTag = p.isActive ? ' [当前活动]' : '';
        output += `${i + 1}. ${p.name}${activeTag}\n`;
        output += `   路径: ${p.path}\n`;
        output += `   页数: ${p.slideCount}\n`;
      });

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取演示文稿列表失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取演示文稿列表出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 切换到指定演示文稿
 */
export const switchPresentationDefinition: ToolDefinition = {
  name: 'wps_ppt_switch_presentation',
  description: `切换到指定名称的演示文稿。

使用场景：
- "切换到另一个PPT"
- "打开那个叫xxx的演示文稿"
- "切换演示文稿"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '要切换到的演示文稿名称',
      },
    },
    required: ['name'],
  },
};

export const switchPresentationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name } = args as {
    name: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
    }>(
      'switchPresentation',
      { name },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已切换到演示文稿: ${response.data.name}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换演示文稿失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换演示文稿出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置演示文稿主题
 */
export const setSlideThemeDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_theme',
  description: `设置演示文稿主题。

使用场景：
- "切换PPT主题"
- "应用商务主题"
- "更换演示文稿风格"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      theme: {
        type: 'string',
        description: '主题名称',
      },
    },
    required: ['theme'],
  },
};

export const setSlideThemeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { theme } = args as { theme: string };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideTheme', // NOTE: macOS未实现，仅Windows支持
      { theme },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `主题已设置为: ${theme}` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置主题失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置主题出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 复制幻灯片
 */
export const copySlideDefinition: ToolDefinition = {
  name: 'wps_ppt_copy_slide',
  description: `复制幻灯片到指定位置。

使用场景：
- "复制第2页幻灯片"
- "把这页复制到第5页后面"
- "克隆当前幻灯片"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '要复制的幻灯片索引（从1开始）',
      },
      targetIndex: {
        type: 'number',
        description: '目标位置索引（从1开始），不填则在原位置后插入',
      },
    },
    required: ['slideIndex'],
  },
};

export const copySlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, targetIndex } = args as {
    slideIndex: number;
    targetIndex?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      newIndex: number;
    }>(
      'duplicateSlide',
      { slideIndex, targetIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片已复制！\n源页: 第${slideIndex}页\n新页位置: 第${response.data.newIndex}页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `复制幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `复制幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 在幻灯片中插入图片
 */
export const insertSlideImageDefinition: ToolDefinition = {
  name: 'wps_ppt_insert_slide_image',
  description: `在幻灯片中插入图片。

使用场景：
- "在第1页插入一张图片"
- "添加图片到幻灯片"
- "把这个图片放到PPT里"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
      imagePath: {
        type: 'string',
        description: '图片文件的完整路径',
      },
      left: {
        type: 'number',
        description: '左边距（像素），默认100',
      },
      top: {
        type: 'number',
        description: '上边距（像素），默认100',
      },
    },
    required: ['slideIndex', 'imagePath'],
  },
};

export const insertSlideImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, imagePath, left, top } = args as {
    slideIndex: number;
    imagePath: string;
    left?: number;
    top?: number;
  };

  try {
    // 跨平台参数对齐：底层 insertImage handler 读取 path/filePath，需同时发送别名
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'insertImage',
      { slideIndex, imagePath, path: imagePath, filePath: imagePath, left, top },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图片已插入到第${slideIndex}页幻灯片！\n图片路径: ${imagePath}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入图片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 从其它演示文稿导入整页幻灯片（跨PPT整合，保留来源格式）
 */
export const insertSlidesFromFileDefinition: ToolDefinition = {
  name: 'wps_ppt_insert_slides_from_file',
  description: `从另一个 PPT 文件把整页幻灯片插入到【当前活动演示文稿】，并保留来源幻灯片的原始格式（字体/配色/版式/图片）。用于把多个 PPT 整合成一个。

使用场景：
- "把可行性报告.pptx 的第3到5页插到当前PPT第10页后面"
- "整合多个PPT：把另一个演示文稿的所有幻灯片合并进来"
- "从某个PPT复制整页过来，保持原样式"

说明：
- 先用 wps_ppt_switch_presentation 切换到【目标/接收页】演示文稿，再调用本工具
- afterIndex 表示插入到第几页之后（0=插到最前，不填=追加到末尾）
- slideStart/slideEnd 指定只导入来源文件的某段页码范围，不填则导入全部
- 底层调用 WPS COM Slides.InsertFromFile，原样保留来源格式，避免AI重排导致的版式失真`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      filePath: {
        type: 'string',
        description: '来源 PPT 文件的完整路径',
      },
      afterIndex: {
        type: 'number',
        description: '插入到当前演示文稿第几页之后（0=最前，不填=末尾追加）',
      },
      slideStart: {
        type: 'number',
        description: '来源文件起始页码（从1开始，可选）',
      },
      slideEnd: {
        type: 'number',
        description: '来源文件结束页码（可选，与slideStart配合）',
      },
    },
    required: ['filePath'],
  },
};

export const insertSlidesFromFileHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { filePath, afterIndex, slideStart, slideEnd } = args as {
    filePath: string;
    afterIndex?: number;
    slideStart?: number;
    slideEnd?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      inserted: number;
      afterIndex: number;
      totalSlides: number;
      source: string;
    }>(
      'insertSlidesFromFile',
      { filePath, path: filePath, afterIndex, slideStart, slideEnd },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已从来源文件导入 ${response.data.inserted} 页幻灯片（保留来源格式）！\n来源: ${filePath}\n插入位置: 第 ${response.data.afterIndex} 页之后\n当前演示文稿总页数: ${response.data.totalSlides}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `导入幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `导入幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 锁定后续所有 PPT 操作的目标演示文稿（避免多文稿打开时活动文稿漂移改错文件）
 */
export const setActiveTargetDefinition: ToolDefinition = {
  name: 'wps_ppt_set_active_target',
  description: `锁定后续所有 PPT 操作的【目标演示文稿】（按文件名）。锁定后，本服务器会给所有演示类调用自动注入该文稿名，底层精确定位，**彻底避免同时打开多个 PPT 时"活动文稿漂移"导致改错文件**。

使用场景：
- 批量改某个 PPT 前先锁定："把目标锁定为 关节模组申报PPT-生成版.pptx"
- 传 clear=true 或留空 name 取消锁定，恢复使用当前活动文稿

强烈建议：在对一个特定 PPT 做成批增改前，先调用本工具锁定它。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '目标演示文稿文件名（如 关节模组申报PPT-生成版.pptx）；留空或 clear=true 则取消锁定',
      },
      clear: {
        type: 'boolean',
        description: '为 true 时清除锁定，恢复使用当前活动文稿',
      },
    },
    required: [],
  },
};

export const setActiveTargetHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, clear } = args as { name?: string; clear?: boolean };

  if (clear || !name) {
    setPptTarget(undefined);
    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: '已取消目标文稿锁定，后续 PPT 操作使用当前活动文稿。' }],
    };
  }

  try {
    const resp = await wpsClient.executeMethod<{ presentations: Array<{ name: string }> }>(
      'getOpenPresentations',
      {}
    );
    const names = resp.success && resp.data ? resp.data.presentations.map((x) => x.name) : [];
    const matched = names.includes(name);
    setPptTarget(name);
    const warn = matched
      ? ''
      : `\n⚠️ 当前打开的文稿中没有完全同名的「${name}」。已打开：${names.join(' / ') || '(无)'}。若名称不符，后续操作会报"未找到目标文稿"。`;
    return {
      id: uuidv4(),
      success: true,
      content: [
        {
          type: 'text',
          text: `已锁定目标演示文稿：${name}\n后续所有 PPT 操作都会定位到该文稿（不受活动窗口切换影响）。${warn}`,
        },
      ],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    setPptTarget(name);
    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: `已锁定目标演示文稿：${name}（打开列表校验失败：${errMsg}）` }],
    };
  }
};

/**
 * 导出所有演示文稿管理相关的Tools
 */
export const presentationTools: RegisteredTool[] = [
  { definition: createPresentationDefinition, handler: createPresentationHandler },
  { definition: openPresentationDefinition, handler: openPresentationHandler },
  { definition: closePresentationDefinition, handler: closePresentationHandler },
  { definition: getOpenPresentationsDefinition, handler: getOpenPresentationsHandler },
  { definition: switchPresentationDefinition, handler: switchPresentationHandler },
  { definition: setSlideThemeDefinition, handler: setSlideThemeHandler },
  { definition: copySlideDefinition, handler: copySlideHandler },
  { definition: insertSlideImageDefinition, handler: insertSlideImageHandler },
  { definition: insertSlidesFromFileDefinition, handler: insertSlidesFromFileHandler },
  { definition: setActiveTargetDefinition, handler: setActiveTargetHandler },
];

export default presentationTools;

/**
 * Input: PPT 杂项操作参数（母版、3D、超链接、搜索、放映）
 * Output: 杂项操作结果
 * Pos: PPT 杂项工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT杂项Tools - 母版/3D/超链接/搜索/放映模块
 *
 * 包含：
 * - wps_ppt_get_slide_master: 获取母版信息
 * - wps_ppt_set_master_background: 设置母版背景
 * - wps_ppt_add_master_element: 添加母版元素
 * - wps_ppt_set_3d_rotation: 设置3D旋转
 * - wps_ppt_set_3d_depth: 设置3D深度
 * - wps_ppt_set_3d_material: 设置3D材质
 * - wps_ppt_create_3d_text: 创建3D文字
 * - wps_ppt_add_ppt_hyperlink: 添加超链接
 * - wps_ppt_remove_ppt_hyperlink: 移除超链接
 * - wps_ppt_find_ppt_text: 搜索文本
 * - wps_ppt_replace_ppt_text: 替换文本
 * - wps_ppt_start_slide_show: 开始放映
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

// ==================== 母版操作 ====================

/**
 * 获取母版信息
 * 获取当前演示文稿的母版布局信息
 */
export const getSlideMasterDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_master',
  description: `获取当前演示文稿的母版信息。

返回母版的布局、背景、元素等详细信息。

使用场景：
- "查看母版信息"
- "获取母版布局"
- "看看母版有什么元素"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getSlideMasterHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      master: Record<string, unknown>;
    }>(
      'getSlideMaster',
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
            text: `母版信息获取成功！\n${JSON.stringify(response.data.master, null, 2)}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取母版信息失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取母版信息出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置母版背景
 * 修改母版的背景样式
 */
export const setMasterBackgroundDefinition: ToolDefinition = {
  name: 'wps_ppt_set_master_background',
  description: `设置母版背景样式。

支持纯色、渐变、图片等背景类型。

使用场景：
- "修改母版背景为蓝色"
- "设置母版背景渐变"
- "给母版换个背景图片"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      background: {
        type: 'object',
        description: '背景配置对象，支持 {type:"solid",color:"#xxx"}, {type:"gradient",colors:[...]}, {type:"image",path:"..."}',
      },
    },
    required: ['background'],
  },
};

export const setMasterBackgroundHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { background } = args as {
    background: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setMasterBackground',
      { background },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `母版背景设置成功！\n背景类型: ${(background as Record<string, unknown>).type || '自定义'}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置母版背景失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置母版背景出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 添加母版元素
 * 向母版中添加新的元素（文本框、形状、图片等）
 */
export const addMasterElementDefinition: ToolDefinition = {
  name: 'wps_ppt_add_master_element',
  description: `向母版中添加新元素。

支持添加文本框、形状、图片、Logo等母版级元素。

使用场景：
- "在母版上添加公司Logo"
- "给母版加个页脚"
- "在母版添加水印"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      element: {
        type: 'object',
        description: '元素配置对象，如 {type:"textbox",text:"...",left:0,top:0,width:100,height:50} 或 {type:"image",path:"...",left:0,top:0}',
      },
    },
    required: ['element'],
  },
};

export const addMasterElementHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { element } = args as {
    element: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      elementId?: number;
    }>(
      'addMasterElement',
      { element },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `母版元素添加成功！\n元素类型: ${(element as Record<string, unknown>).type || '未知'}${response.data.elementId ? `\n元素ID: ${response.data.elementId}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加母版元素失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加母版元素出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 3D 操作 ====================

/**
 * 设置3D旋转
 * 为形状设置3D旋转效果
 */
export const set3DRotationDefinition: ToolDefinition = {
  name: 'wps_ppt_set_3d_rotation',
  description: `设置幻灯片中形状的3D旋转效果。

通过调整X/Y/Z轴旋转角度实现3D透视效果。

使用场景：
- "给形状添加3D旋转效果"
- "设置3D透视角度"
- "让形状有立体感"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      rotation: {
        type: 'object',
        description: '旋转参数对象，如 {rotX:30,rotY:45,rotZ:0,perspective:50}',
      },
    },
    required: ['slideIndex', 'shapeIndex', 'rotation'],
  },
};

export const set3DRotationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, rotation } = args as {
    slideIndex: number;
    shapeIndex: number;
    rotation: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'set3DRotation',
      { slideIndex, shapeIndex, rotation },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `3D旋转设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n旋转参数: ${JSON.stringify(rotation)}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置3D旋转失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置3D旋转出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置3D深度
 * 为形状设置3D挤出深度
 */
export const set3DDepthDefinition: ToolDefinition = {
  name: 'wps_ppt_set_3d_depth',
  description: `设置幻灯片中形状的3D挤出深度。

通过调整深度值使形状产生立体挤出效果。

使用场景：
- "给形状添加3D深度"
- "设置立体厚度"
- "增加形状深度效果"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      depth: {
        type: 'number',
        description: '挤出深度值（磅），如 20、50、100',
      },
    },
    required: ['slideIndex', 'shapeIndex', 'depth'],
  },
};

export const set3DDepthHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, depth } = args as {
    slideIndex: number;
    shapeIndex: number;
    depth: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'set3DDepth',
      { slideIndex, shapeIndex, depth },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `3D深度设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n深度: ${depth} 磅`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置3D深度失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置3D深度出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置3D材质
 * 为形状设置3D材质效果
 */
export const set3DMaterialDefinition: ToolDefinition = {
  name: 'wps_ppt_set_3d_material',
  description: `设置幻灯片中形状的3D材质效果。

支持的材质类型：
- matte: 哑光
- plastic: 塑料
- metal: 金属
- wireframe: 线框
- soft_edge: 柔化边缘
- flat: 平面
- dark_edge: 暗边

使用场景：
- "给形状设置金属材质"
- "改成塑料质感"
- "用哑光效果"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      material: {
        type: 'string',
        description: '材质类型：matte(哑光)、plastic(塑料)、metal(金属)、wireframe(线框)、soft_edge(柔化)、flat(平面)、dark_edge(暗边)',
      },
    },
    required: ['slideIndex', 'shapeIndex', 'material'],
  },
};

export const set3DMaterialHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, material } = args as {
    slideIndex: number;
    shapeIndex: number;
    material: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'set3DMaterial',
      { slideIndex, shapeIndex, material },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const materialName: Record<string, string> = {
        matte: '哑光',
        plastic: '塑料',
        metal: '金属',
        wireframe: '线框',
        soft_edge: '柔化边缘',
        flat: '平面',
        dark_edge: '暗边',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `3D材质设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n材质: ${materialName[material] || material}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置3D材质失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置3D材质出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 创建3D文字
 * 在幻灯片中创建带有3D效果的文字
 */
export const create3DTextDefinition: ToolDefinition = {
  name: 'wps_ppt_create_3d_text',
  description: `在幻灯片中创建带有3D效果的文字。

可以指定文字内容和3D样式参数。

使用场景：
- "创建3D标题文字"
- "添加立体文字效果"
- "做一个炫酷的3D文字"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      text: {
        type: 'string',
        description: '文字内容',
      },
      style: {
        type: 'object',
        description: '可选的3D样式参数，如 {depth:30,material:"metal",color:"#FFD700",fontSize:48,rotation:{rotX:20,rotY:30}}',
      },
    },
    required: ['slideIndex', 'text'],
  },
};

export const create3DTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, text, style } = args as {
    slideIndex: number;
    text: string;
    style?: Record<string, unknown>;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      shapeId?: number;
    }>(
      'create3DText',
      { slideIndex, text, style: style || {} },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `3D文字创建成功！\n幻灯片: 第 ${slideIndex} 页\n文字: "${text}"${response.data.shapeId ? `\n形状ID: ${response.data.shapeId}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `创建3D文字失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `创建3D文字出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 超链接操作 ====================

/**
 * 添加超链接
 * 为形状添加超链接
 */
export const addPptHyperlinkDefinition: ToolDefinition = {
  name: 'wps_ppt_add_ppt_hyperlink',
  description: `为幻灯片中的形状添加超链接。

支持网页链接、邮件链接、幻灯片内部跳转等。

使用场景：
- "给按钮添加链接"
- "设置点击跳转到网页"
- "添加邮件链接"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      url: {
        type: 'string',
        description: '超链接地址，如 "https://example.com" 或 "mailto:test@example.com"',
      },
    },
    required: ['slideIndex', 'shapeIndex', 'url'],
  },
};

export const addPptHyperlinkHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, url } = args as {
    slideIndex: number;
    shapeIndex: number;
    url: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'addPptHyperlink',
      { slideIndex, shapeIndex, url },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `超链接添加成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n链接: ${url}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加超链接失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加超链接出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 移除超链接
 * 删除形状上的超链接
 */
export const removePptHyperlinkDefinition: ToolDefinition = {
  name: 'wps_ppt_remove_ppt_hyperlink',
  description: `移除幻灯片中形状的超链接。

使用场景：
- "删除按钮的链接"
- "移除超链接"
- "取消链接"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
    },
    required: ['slideIndex', 'shapeIndex'],
  },
};

export const removePptHyperlinkHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex } = args as {
    slideIndex: number;
    shapeIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'removePptHyperlink',
      { slideIndex, shapeIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `超链接移除成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `移除超链接失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `移除超链接出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 搜索操作 ====================

/**
 * 搜索文本
 * 在演示文稿中搜索指定文本
 */
export const findPptTextDefinition: ToolDefinition = {
  name: 'wps_ppt_find_ppt_text',
  description: `在演示文稿中搜索指定文本。

返回包含目标文本的幻灯片页码和形状信息。

使用场景：
- "搜索PPT中的某段文字"
- "查找包含关键词的幻灯片"
- "找到所有提到xxx的位置"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '要搜索的文本内容',
      },
    },
    required: ['text'],
  },
};

export const findPptTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text } = args as {
    text: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      results: Array<{
        slideIndex: number;
        shapeIndex: number;
        text: string;
      }>;
      count: number;
    }>(
      'findPptText',
      { text },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const result = response.data;
      if (result.count === 0) {
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `未找到包含 "${text}" 的内容。`,
            },
          ],
        };
      }

      let output = `搜索完成！找到 ${result.count} 处匹配：\n`;
      result.results.forEach((r, i) => {
        output += `${i + 1}. 第 ${r.slideIndex} 页, 形状 ${r.shapeIndex}: "${r.text}"\n`;
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
        content: [{ type: 'text', text: `搜索文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `搜索文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 替换文本
 * 在演示文稿中查找并替换文本
 */
export const replacePptTextDefinition: ToolDefinition = {
  name: 'wps_ppt_replace_ppt_text',
  description: `在演示文稿中查找并替换文本。

批量替换所有匹配的文本内容。

使用场景：
- "把PPT中所有的A替换成B"
- "批量替换公司名称"
- "修改所有页面的标题"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      find: {
        type: 'string',
        description: '要查找的文本',
      },
      replace: {
        type: 'string',
        description: '替换为的文本',
      },
    },
    required: ['find', 'replace'],
  },
};

export const replacePptTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { find, replace } = args as {
    find: string;
    replace: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      count: number;
    }>(
      'replacePptText',
      { find, replace },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文本替换完成！\n查找: "${find}"\n替换为: "${replace}"\n替换数量: ${response.data.count} 处`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `替换文本失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `替换文本出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 放映操作 ====================

/**
 * 开始放映
 * 从指定页开始幻灯片放映
 */
export const startSlideShowDefinition: ToolDefinition = {
  name: 'wps_ppt_start_slide_show',
  description: `开始幻灯片放映。

可以从指定页面开始放映，默认从第1页开始。

使用场景：
- "放映幻灯片"
- "从第3页开始演示"
- "开始PPT放映"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      fromSlide: {
        type: 'number',
        description: '从第几页开始放映（从1开始），默认从第1页开始',
      },
    },
    required: [],
  },
};

export const startSlideShowHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { fromSlide } = args as {
    fromSlide?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'startSlideShow',
      { fromSlide: fromSlide || 1 },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片放映已开始！\n起始页: 第 ${fromSlide || 1} 页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `开始放映失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `开始放映出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有杂项Tools
 */
export const miscTools: RegisteredTool[] = [
  // 母版操作
  { definition: getSlideMasterDefinition, handler: getSlideMasterHandler },
  { definition: setMasterBackgroundDefinition, handler: setMasterBackgroundHandler },
  { definition: addMasterElementDefinition, handler: addMasterElementHandler },
  // 3D操作
  { definition: set3DRotationDefinition, handler: set3DRotationHandler },
  { definition: set3DDepthDefinition, handler: set3DDepthHandler },
  { definition: set3DMaterialDefinition, handler: set3DMaterialHandler },
  { definition: create3DTextDefinition, handler: create3DTextHandler },
  // 超链接操作
  { definition: addPptHyperlinkDefinition, handler: addPptHyperlinkHandler },
  { definition: removePptHyperlinkDefinition, handler: removePptHyperlinkHandler },
  // 搜索操作
  { definition: findPptTextDefinition, handler: findPptTextHandler },
  { definition: replacePptTextDefinition, handler: replacePptTextHandler },
  // 放映操作
  { definition: startSlideShowDefinition, handler: startSlideShowHandler },
];

export default miscTools;

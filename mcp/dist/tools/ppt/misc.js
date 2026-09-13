"use strict";
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
 * - wps_ppt_add_ppt_hyperlink: 添加超链接
 * - wps_ppt_remove_ppt_hyperlink: 移除超链接
 * - wps_ppt_find_ppt_text: 搜索文本
 * - wps_ppt_replace_ppt_text: 替换文本
 * - wps_ppt_start_slide_show: 开始放映
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.miscTools = exports.startSlideShowHandler = exports.startSlideShowDefinition = exports.replacePptTextHandler = exports.replacePptTextDefinition = exports.findPptTextHandler = exports.findPptTextDefinition = exports.removePptHyperlinkHandler = exports.removePptHyperlinkDefinition = exports.addPptHyperlinkHandler = exports.addPptHyperlinkDefinition = exports.set3DMaterialHandler = exports.set3DMaterialDefinition = exports.set3DDepthHandler = exports.set3DDepthDefinition = exports.set3DRotationHandler = exports.set3DRotationDefinition = exports.addMasterElementHandler = exports.addMasterElementDefinition = exports.setMasterBackgroundHandler = exports.setMasterBackgroundDefinition = exports.getSlideMasterHandler = exports.getSlideMasterDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
// ==================== 母版操作 ====================
/**
 * 获取母版信息
 * 获取当前演示文稿的母版布局信息
 */
exports.getSlideMasterDefinition = {
    name: 'wps_ppt_get_slide_master',
    description: `获取当前演示文稿的母版信息。

返回母版的布局、背景、元素等详细信息。

使用场景：
- "查看母版信息"
- "获取母版布局"
- "看看母版有什么元素"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const getSlideMasterHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSlideMaster', {}, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `母版信息获取成功！\n${JSON.stringify(response.data.master, null, 2)}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取母版信息失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取母版信息出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getSlideMasterHandler = getSlideMasterHandler;
/**
 * 设置母版背景
 * 修改母版的背景样式
 */
exports.setMasterBackgroundDefinition = {
    name: 'wps_ppt_set_master_background',
    description: `设置母版背景样式。

支持纯色、渐变、图片等背景类型。

使用场景：
- "修改母版背景为蓝色"
- "设置母版背景渐变"
- "给母版换个背景图片"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const setMasterBackgroundHandler = async (args) => {
    const { background } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setMasterBackground', { background }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `母版背景设置成功！\n背景类型: ${background.type || '自定义'}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置母版背景失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置母版背景出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setMasterBackgroundHandler = setMasterBackgroundHandler;
/**
 * 添加母版元素
 * 向母版中添加新的元素（文本框、形状、图片等）
 */
exports.addMasterElementDefinition = {
    name: 'wps_ppt_add_master_element',
    description: `向母版中添加新元素。

支持添加文本框、形状、图片、Logo等母版级元素。

使用场景：
- "在母版上添加公司Logo"
- "给母版加个页脚"
- "在母版添加水印"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const addMasterElementHandler = async (args) => {
    const { element } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addMasterElement', { element }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `母版元素添加成功！\n元素类型: ${element.type || '未知'}${response.data.elementId ? `\n元素ID: ${response.data.elementId}` : ''}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `添加母版元素失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `添加母版元素出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.addMasterElementHandler = addMasterElementHandler;
// ==================== 3D 操作 ====================
/**
 * 设置3D旋转
 * 为形状设置3D旋转效果
 */
exports.set3DRotationDefinition = {
    name: 'wps_ppt_set_3d_rotation',
    description: `设置幻灯片中形状的3D旋转效果。

通过调整X/Y/Z轴旋转角度实现3D透视效果。

使用场景：
- "给形状添加3D旋转效果"
- "设置3D透视角度"
- "让形状有立体感"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const set3DRotationHandler = async (args) => {
    const { slideIndex, shapeIndex, rotation } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('set3DRotation', { slideIndex, shapeIndex, rotation }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `3D旋转设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n旋转参数: ${JSON.stringify(rotation)}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置3D旋转失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置3D旋转出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.set3DRotationHandler = set3DRotationHandler;
/**
 * 设置3D深度
 * 为形状设置3D挤出深度
 */
exports.set3DDepthDefinition = {
    name: 'wps_ppt_set_3d_depth',
    description: `设置幻灯片中形状的3D挤出深度。

通过调整深度值使形状产生立体挤出效果。

使用场景：
- "给形状添加3D深度"
- "设置立体厚度"
- "增加形状深度效果"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const set3DDepthHandler = async (args) => {
    const { slideIndex, shapeIndex, depth } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('set3DDepth', { slideIndex, shapeIndex, depth }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `3D深度设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n深度: ${depth} 磅`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置3D深度失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置3D深度出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.set3DDepthHandler = set3DDepthHandler;
/**
 * 设置3D材质
 * 为形状设置3D材质效果
 */
exports.set3DMaterialDefinition = {
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
    category: tools_1.ToolCategory.PRESENTATION,
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
const set3DMaterialHandler = async (args) => {
    const { slideIndex, shapeIndex, material } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('set3DMaterial', { slideIndex, shapeIndex, material }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const materialName = {
                matte: '哑光',
                plastic: '塑料',
                metal: '金属',
                wireframe: '线框',
                soft_edge: '柔化边缘',
                flat: '平面',
                dark_edge: '暗边',
            };
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `3D材质设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n材质: ${materialName[material] || material}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置3D材质失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置3D材质出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.set3DMaterialHandler = set3DMaterialHandler;
exports.addPptHyperlinkDefinition = {
    name: 'wps_ppt_add_ppt_hyperlink',
    description: `为幻灯片中的形状添加超链接。

支持网页链接、邮件链接、幻灯片内部跳转等。

使用场景：
- "给按钮添加链接"
- "设置点击跳转到网页"
- "添加邮件链接"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const addPptHyperlinkHandler = async (args) => {
    const { slideIndex, shapeIndex, url } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addPptHyperlink', { slideIndex, shapeIndex, url }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `超链接添加成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n链接: ${url}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `添加超链接失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `添加超链接出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.addPptHyperlinkHandler = addPptHyperlinkHandler;
/**
 * 移除超链接
 * 删除形状上的超链接
 */
exports.removePptHyperlinkDefinition = {
    name: 'wps_ppt_remove_ppt_hyperlink',
    description: `移除幻灯片中形状的超链接。

使用场景：
- "删除按钮的链接"
- "移除超链接"
- "取消链接"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const removePptHyperlinkHandler = async (args) => {
    const { slideIndex, shapeIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('removePptHyperlink', { slideIndex, shapeIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `超链接移除成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `移除超链接失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `移除超链接出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.removePptHyperlinkHandler = removePptHyperlinkHandler;
// ==================== 搜索操作 ====================
/**
 * 搜索文本
 * 在演示文稿中搜索指定文本
 */
exports.findPptTextDefinition = {
    name: 'wps_ppt_find_ppt_text',
    description: `在演示文稿中搜索指定文本。

返回包含目标文本的幻灯片页码和形状信息。

使用场景：
- "搜索PPT中的某段文字"
- "查找包含关键词的幻灯片"
- "找到所有提到xxx的位置"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const findPptTextHandler = async (args) => {
    const { text } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('findPptText', { text }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const result = response.data;
            if (result.count === 0) {
                return {
                    id: (0, uuid_1.v4)(),
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
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: output }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `搜索文本失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `搜索文本出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.findPptTextHandler = findPptTextHandler;
/**
 * 替换文本
 * 在演示文稿中查找并替换文本
 */
exports.replacePptTextDefinition = {
    name: 'wps_ppt_replace_ppt_text',
    description: `在演示文稿中查找并替换文本。

批量替换所有匹配的文本内容。

使用场景：
- "把PPT中所有的A替换成B"
- "批量替换公司名称"
- "修改所有页面的标题"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const replacePptTextHandler = async (args) => {
    const { find, replace } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('replacePptText', { find, replace }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `文本替换完成！\n查找: "${find}"\n替换为: "${replace}"\n替换数量: ${response.data.count} 处`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `替换文本失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `替换文本出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.replacePptTextHandler = replacePptTextHandler;
// ==================== 放映操作 ====================
/**
 * 开始放映
 * 从指定页开始幻灯片放映
 */
exports.startSlideShowDefinition = {
    name: 'wps_ppt_start_slide_show',
    description: `开始幻灯片放映。

可以从指定页面开始放映，默认从第1页开始。

使用场景：
- "放映幻灯片"
- "从第3页开始演示"
- "开始PPT放映"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const startSlideShowHandler = async (args) => {
    const { fromSlide } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('startSlideShow', { fromSlide: fromSlide || 1 }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片放映已开始！\n起始页: 第 ${fromSlide || 1} 页`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `开始放映失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `开始放映出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.startSlideShowHandler = startSlideShowHandler;
/**
 * 导出所有杂项Tools
 */
exports.miscTools = [
    // 母版操作
    { definition: exports.getSlideMasterDefinition, handler: exports.getSlideMasterHandler },
    { definition: exports.setMasterBackgroundDefinition, handler: exports.setMasterBackgroundHandler },
    { definition: exports.addMasterElementDefinition, handler: exports.addMasterElementHandler },
    // 3D操作
    { definition: exports.set3DRotationDefinition, handler: exports.set3DRotationHandler },
    { definition: exports.set3DDepthDefinition, handler: exports.set3DDepthHandler },
    { definition: exports.set3DMaterialDefinition, handler: exports.set3DMaterialHandler },
    // 超链接操作
    { definition: exports.addPptHyperlinkDefinition, handler: exports.addPptHyperlinkHandler },
    { definition: exports.removePptHyperlinkDefinition, handler: exports.removePptHyperlinkHandler },
    // 搜索操作
    { definition: exports.findPptTextDefinition, handler: exports.findPptTextHandler },
    { definition: exports.replacePptTextDefinition, handler: exports.replacePptTextHandler },
    // 放映操作
    { definition: exports.startSlideShowDefinition, handler: exports.startSlideShowHandler },
];
exports.default = exports.miscTools;
//# sourceMappingURL=misc.js.map
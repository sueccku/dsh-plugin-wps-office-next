"use strict";
/**
 * Input: PPT 形状操作参数（删除、获取、位置、阴影、渐变、边框、透明度、对齐、分布、组合）
 * Output: 形状操作结果
 * Pos: PPT 形状基础工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT形状基础Tools - 形状管理模块
 * 处理形状的删除、获取、定位、样式设置、对齐分布和组合操作
 *
 * 包含：
 * - wps_ppt_delete_shape: 删除形状
 * - wps_ppt_get_shapes: 获取幻灯片中的形状列表
 * - wps_ppt_set_shape_position: 设置形状位置和大小
 * - wps_ppt_set_shape_shadow: 设置形状阴影
 * - wps_ppt_set_shape_gradient: 设置形状渐变填充
 * - wps_ppt_set_shape_border: 设置形状边框
 * - wps_ppt_set_shape_transparency: 设置形状透明度
 * - wps_ppt_align_shapes: 对齐多个形状
 * - wps_ppt_distribute_shapes: 等距分布多个形状
 * - wps_ppt_group_shapes: 组合多个形状
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.shapeBasicTools = exports.groupShapesHandler = exports.groupShapesDefinition = exports.distributeShapesHandler = exports.distributeShapesDefinition = exports.alignShapesHandler = exports.alignShapesDefinition = exports.setShapeTransparencyHandler = exports.setShapeTransparencyDefinition = exports.setShapeBorderHandler = exports.setShapeBorderDefinition = exports.setShapeGradientHandler = exports.setShapeGradientDefinition = exports.setShapeShadowHandler = exports.setShapeShadowDefinition = exports.setShapePositionHandler = exports.setShapePositionDefinition = exports.getShapesHandler = exports.getShapesDefinition = exports.deleteShapeHandler = exports.deleteShapeDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
// ==================== 1. 删除形状 ====================
exports.deleteShapeDefinition = {
    name: 'wps_ppt_delete_shape',
    description: `删除幻灯片中指定的形状。

使用场景：
- "删除第2页的第3个形状"
- "移除这个形状"`,
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
const deleteShapeHandler = async (args) => {
    const { slideIndex, shapeIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteShape', { slideIndex, shapeIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `形状删除成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `删除形状失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `删除形状出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.deleteShapeHandler = deleteShapeHandler;
// ==================== 2. 获取形状列表 ====================
exports.getShapesDefinition = {
    name: 'wps_ppt_get_shapes',
    description: `获取幻灯片中所有形状的列表信息。

返回每个形状的类型、位置、大小等属性。

使用场景：
- "这页PPT有哪些形状"
- "列出第3页的所有元素"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
        },
        required: ['slideIndex'],
    },
};
const getShapesHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getShapes', { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const shapes = response.data.shapes;
            let output = `幻灯片第 ${slideIndex} 页的形状列表（共 ${shapes.length} 个）：\n\n`;
            shapes.forEach((s) => {
                output += `[${s.index}] ${s.name} (${s.type})\n`;
                output += `    位置: (${s.left}, ${s.top}) 大小: ${s.width} x ${s.height}\n`;
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
                content: [{ type: 'text', text: `获取形状列表失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取形状列表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getShapesHandler = getShapesHandler;
// ==================== 3. 设置形状位置 ====================
exports.setShapePositionDefinition = {
    name: 'wps_ppt_set_shape_position',
    description: `设置幻灯片中指定形状的位置和大小。

使用场景：
- "把这个形状移到左上角"
- "调整形状大小"
- "设置形状位置为(100, 200)"`,
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
            left: {
                type: 'number',
                description: '左边距（磅）',
            },
            top: {
                type: 'number',
                description: '上边距（磅）',
            },
            width: {
                type: 'number',
                description: '宽度（磅），可选',
            },
            height: {
                type: 'number',
                description: '高度（磅），可选',
            },
        },
        required: ['slideIndex', 'shapeIndex', 'left', 'top'],
    },
};
const setShapePositionHandler = async (args) => {
    const { slideIndex, shapeIndex, left, top, width, height } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setShapePosition', { slideIndex, shapeIndex, left, top, width, height }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            let text = `形状位置设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n位置: (${left}, ${top})`;
            if (width !== undefined)
                text += `\n宽度: ${width}`;
            if (height !== undefined)
                text += `\n高度: ${height}`;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置形状位置失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置形状位置出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setShapePositionHandler = setShapePositionHandler;
// ==================== 4. 设置形状阴影 ====================
exports.setShapeShadowDefinition = {
    name: 'wps_ppt_set_shape_shadow',
    description: `设置幻灯片中指定形状的阴影效果。

shadow对象属性：
- enabled: 是否启用阴影 (boolean)
- color: 阴影颜色，如 "#000000"
- blur: 模糊半径（磅）
- offsetX: 水平偏移（磅）
- offsetY: 垂直偏移（磅）
- opacity: 透明度 (0-1)

使用场景：
- "给这个形状加阴影"
- "设置阴影效果"`,
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
            shadow: {
                type: 'object',
                description: '阴影配置对象',
                properties: {
                    enabled: { type: 'boolean', description: '是否启用阴影' },
                    color: { type: 'string', description: '阴影颜色' },
                    blur: { type: 'number', description: '模糊半径（磅）' },
                    offsetX: { type: 'number', description: '水平偏移（磅）' },
                    offsetY: { type: 'number', description: '垂直偏移（磅）' },
                    opacity: { type: 'number', description: '透明度 (0-1)' },
                },
            },
        },
        required: ['slideIndex', 'shapeIndex', 'shadow'],
    },
};
const setShapeShadowHandler = async (args) => {
    const { slideIndex, shapeIndex, shadow } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setShapeShadow', { slideIndex, shapeIndex, shadow }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `形状阴影设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置形状阴影失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置形状阴影出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setShapeShadowHandler = setShapeShadowHandler;
// ==================== 5. 设置形状渐变填充 ====================
exports.setShapeGradientDefinition = {
    name: 'wps_ppt_set_shape_gradient',
    description: `设置幻灯片中指定形状的渐变填充效果。

gradient对象属性：
gradient对象属性：
- stops: 渐变色标数组 [{color: "#FF0000", position: 0}, {color: "#0000FF", position: 1}]（当前仅支持两个色标）
注意：渐变角度与类型在 WPS 上不可设置（会挂起 COM 调用），因此不再提供 angle/type 参数。

使用场景：
- "给形状加渐变色"
- "设置从红到蓝的渐变"`,
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
            gradient: {
                type: 'object',
                description: '渐变配置对象',
                properties: {
                    stops: {
                        type: 'array',
                        description: '渐变色标数组',
                        items: {
                            type: 'object',
                            properties: {
                                color: { type: 'string', description: '颜色值' },
                                position: { type: 'number', description: '位置 (0-1)' },
                            },
                        },
                    },
                },
            },
        },
        required: ['slideIndex', 'shapeIndex', 'gradient'],
    },
};
const setShapeGradientHandler = async (args) => {
    const { slideIndex, shapeIndex, gradient } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setShapeGradient', { slideIndex, shapeIndex, gradient }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `形状渐变填充设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置形状渐变失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置形状渐变出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setShapeGradientHandler = setShapeGradientHandler;
// ==================== 6. 设置形状边框 ====================
exports.setShapeBorderDefinition = {
    name: 'wps_ppt_set_shape_border',
    description: `设置幻灯片中指定形状的边框样式。

border对象属性：
- enabled: 是否启用边框 (boolean)
- color: 边框颜色，如 "#000000"
- weight: 边框粗细（磅）
- style: 边框样式，如 "solid"(实线)、"dash"(虚线)、"dot"(点线)

使用场景：
- "给形状加边框"
- "设置红色虚线边框"`,
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
            border: {
                type: 'object',
                description: '边框配置对象',
                properties: {
                    enabled: { type: 'boolean', description: '是否启用边框' },
                    color: { type: 'string', description: '边框颜色' },
                    weight: { type: 'number', description: '边框粗细（磅）' },
                    style: { type: 'string', description: '边框样式', enum: ['solid', 'dash', 'dot', 'dash_dot', 'dash_dot_dot'] },
                },
            },
        },
        required: ['slideIndex', 'shapeIndex', 'border'],
    },
};
const setShapeBorderHandler = async (args) => {
    const { slideIndex, shapeIndex, border } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setShapeBorder', { slideIndex, shapeIndex, border }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `形状边框设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置形状边框失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置形状边框出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setShapeBorderHandler = setShapeBorderHandler;
// ==================== 7. 设置形状透明度 ====================
exports.setShapeTransparencyDefinition = {
    name: 'wps_ppt_set_shape_transparency',
    description: `设置幻灯片中指定形状的透明度。

使用场景：
- "把这个形状设为半透明"
- "设置形状透明度为50%"`,
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
            transparency: {
                type: 'number',
                description: '透明度值 (0-100)，0为完全不透明，100为完全透明',
            },
        },
        required: ['slideIndex', 'shapeIndex', 'transparency'],
    },
};
const setShapeTransparencyHandler = async (args) => {
    const { slideIndex, shapeIndex, transparency } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setShapeTransparency', { slideIndex, shapeIndex, transparency }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `形状透明度设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n透明度: ${transparency}%`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置形状透明度失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置形状透明度出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setShapeTransparencyHandler = setShapeTransparencyHandler;
// ==================== 8. 对齐多个形状 ====================
exports.alignShapesDefinition = {
    name: 'wps_ppt_align_shapes',
    description: `对齐幻灯片中的多个形状。

支持的对齐方式：
- left: 左对齐
- center: 水平居中
- right: 右对齐
- top: 顶部对齐
- middle: 垂直居中
- bottom: 底部对齐

使用场景：
- "把这几个形状左对齐"
- "让第1、3、5个形状垂直居中"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            shapeIndices: {
                type: 'array',
                items: { type: 'number' },
                description: '要对齐的形状索引数组（从1开始）',
            },
            alignment: {
                type: 'string',
                description: '对齐方式',
                enum: ['left', 'center', 'right', 'top', 'middle', 'bottom'],
            },
        },
        // shapeIndices is optional: without it every shape on the slide is aligned.
        required: ['slideIndex', 'alignment'],
    },
};
const alignShapesHandler = async (args) => {
    const { slideIndex, shapeIndices, alignment } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('alignShapes', { slideIndex, shapeIndices, alignment }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const alignName = {
                left: '左对齐',
                center: '水平居中',
                right: '右对齐',
                top: '顶部对齐',
                middle: '垂直居中',
                bottom: '底部对齐',
            };
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        // shapeIndices is optional: without it the action aligns every shape on the slide, and
                        // dereferencing it here used to fail with "Cannot read properties of undefined".
                        text: `形状对齐完成！\n幻灯片: 第 ${slideIndex} 页\n对齐方式: ${alignName[alignment] || alignment}\n形状数量: ${shapeIndices ? shapeIndices.length + ' 个' : (response.data?.shapes ?? '全部')}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `对齐形状失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `对齐形状出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.alignShapesHandler = alignShapesHandler;
// ==================== 9. 等距分布多个形状 ====================
exports.distributeShapesDefinition = {
    name: 'wps_ppt_distribute_shapes',
    description: `等距分布幻灯片中的多个形状。

支持的分布方向：
- horizontal: 水平等距分布
- vertical: 垂直等距分布

使用场景：
- "让这些形状水平等距排列"
- "垂直均匀分布这些元素"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            shapeIndices: {
                type: 'array',
                items: { type: 'number' },
                description: '要分布的形状索引数组（从1开始）',
            },
            direction: {
                type: 'string',
                description: '分布方向',
                enum: ['horizontal', 'vertical'],
            },
        },
        required: ['slideIndex', 'shapeIndices', 'direction'],
    },
};
const distributeShapesHandler = async (args) => {
    const { slideIndex, shapeIndices, direction } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('distributeShapes', { slideIndex, shapeIndices, direction }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const dirName = direction === 'horizontal' ? '水平等距' : '垂直等距';
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `形状分布完成！\n幻灯片: 第 ${slideIndex} 页\n分布方向: ${dirName}\n形状数量: ${shapeIndices.length} 个`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `分布形状失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `分布形状出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.distributeShapesHandler = distributeShapesHandler;
// ==================== 10. 组合多个形状 ====================
exports.groupShapesDefinition = {
    name: 'wps_ppt_group_shapes',
    description: `将幻灯片中的多个形状组合为一个组。

使用场景：
- "把这几个形状组合在一起"
- "将第1、2、3个形状编组"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片页码（从1开始）',
            },
            shapeIndices: {
                type: 'array',
                items: { type: 'number' },
                description: '要组合的形状索引数组（从1开始，至少2个）',
            },
        },
        required: ['slideIndex', 'shapeIndices'],
    },
};
const groupShapesHandler = async (args) => {
    const { slideIndex, shapeIndices } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('groupShapes', { slideIndex, shapeIndices }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            let text = `形状组合成功！\n幻灯片: 第 ${slideIndex} 页\n组合形状: ${shapeIndices.length} 个`;
            if (response.data?.groupIndex) {
                text += `\n组合索引: ${response.data.groupIndex}`;
            }
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `组合形状失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `组合形状出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.groupShapesHandler = groupShapesHandler;
/**
 * 导出所有形状基础相关的Tools
 */
exports.shapeBasicTools = [
    { definition: exports.deleteShapeDefinition, handler: exports.deleteShapeHandler },
    { definition: exports.getShapesDefinition, handler: exports.getShapesHandler },
    { definition: exports.setShapePositionDefinition, handler: exports.setShapePositionHandler },
    { definition: exports.setShapeShadowDefinition, handler: exports.setShapeShadowHandler },
    { definition: exports.setShapeGradientDefinition, handler: exports.setShapeGradientHandler },
    { definition: exports.setShapeBorderDefinition, handler: exports.setShapeBorderHandler },
    { definition: exports.setShapeTransparencyDefinition, handler: exports.setShapeTransparencyHandler },
    { definition: exports.alignShapesDefinition, handler: exports.alignShapesHandler },
    { definition: exports.distributeShapesDefinition, handler: exports.distributeShapesHandler },
    { definition: exports.groupShapesDefinition, handler: exports.groupShapesHandler },
];
exports.default = exports.shapeBasicTools;
//# sourceMappingURL=shape-basic.js.map
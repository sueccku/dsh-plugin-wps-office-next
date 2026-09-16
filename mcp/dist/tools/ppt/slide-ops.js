"use strict";
/**
 * Input: 幻灯片操作工具参数
 * Output: 幻灯片操作结果
 * Pos: PPT 幻灯片操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 幻灯片操作Tools - 删除/复制/移动/查询/切换/布局/备注
 *
 * 包含：
 * - wps_ppt_delete_slide: 删除指定幻灯片
 * - wps_ppt_duplicate_slide: 复制幻灯片
 * - wps_ppt_move_slide: 移动幻灯片到指定位置
 * - wps_ppt_get_slide_count: 获取幻灯片总数
 * - wps_ppt_get_slide_info: 获取指定幻灯片的详细信息
 * - wps_ppt_switch_slide: 切换到指定幻灯片
 * - wps_ppt_set_slide_layout: 设置幻灯片版式布局
 * - wps_ppt_get_slide_notes: 获取幻灯片备注内容
 * - wps_ppt_set_slide_notes: 设置幻灯片备注
 * - wps_ppt_add_shape: 添加形状
 * - wps_ppt_set_shape_style: 设置形状样式
 * - wps_ppt_add_textbox: 添加文本框
 * - wps_ppt_set_slide_title: 设置幻灯片标题
 * - wps_ppt_insert_image: 插入图片
 * - wps_ppt_set_shape_text: 设置形状文字
 * - wps_ppt_set_animation: 设置元素动画
 * - wps_ppt_set_background: 设置幻灯片背景
 * - wps_ppt_set_slide_size: 设置幻灯片尺寸
 * - wps_ppt_set_transition: 设置幻灯片切换效果
 * - wps_ppt_add_chart: 在幻灯片中插入图表
 * - wps_ppt_set_shape_fill: 设置形状填充颜色
 * - wps_ppt_add_speaker_notes: 添加演讲者备注
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.slideOpsTools = exports.setShapeFillHandler = exports.setShapeFillDefinition = exports.setSlideSizeHandler = exports.setSlideSizeDefinition = exports.setShapeTextHandler = exports.setShapeTextDefinition = exports.setSlideTitleHandler = exports.setSlideTitleDefinition = exports.addTextboxHandler = exports.addTextboxDefinition = exports.setShapeStyleHandler = exports.setShapeStyleDefinition = exports.addShapeHandler = exports.addShapeDefinition = exports.setSlideNotesHandler = exports.setSlideNotesDefinition = exports.getSlideNotesHandler = exports.getSlideNotesDefinition = exports.setSlideLayoutHandler = exports.setSlideLayoutDefinition = exports.switchSlideHandler = exports.switchSlideDefinition = exports.getSlideInfoHandler = exports.getSlideInfoDefinition = exports.getSlideCountHandler = exports.getSlideCountDefinition = exports.moveSlideHandler = exports.moveSlideDefinition = exports.deleteSlideHandler = exports.deleteSlideDefinition = void 0;
const uuid_1 = require("uuid");
const impact_1 = require("../impact");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
// ============================================================
// 1. wps_ppt_delete_slide - 删除指定幻灯片
// ============================================================
exports.deleteSlideDefinition = {
    name: 'wps_ppt_delete_slide',
    description: `删除指定的幻灯片。

使用场景：
- "删除第3页幻灯片"
- "把最后一页删掉"
- "移除多余的页面"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '要删除的幻灯片索引（从1开始）',
            },
        },
        required: ['slideIndex'],
    },
};
const deleteSlideHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('deleteSlide', { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片删除成功！\n已删除: 第 ${slideIndex} 页${(0, impact_1.impactText)(response.data?.impact)}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `删除幻灯片失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `删除幻灯片出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.deleteSlideHandler = deleteSlideHandler;
// ============================================================
// 2. wps_ppt_duplicate_slide - 复制幻灯片
// ============================================================
// ============================================================
// 3. wps_ppt_move_slide - 移动幻灯片到指定位置
// ============================================================
exports.moveSlideDefinition = {
    name: 'wps_ppt_move_slide',
    description: `移动幻灯片到指定位置。

使用场景：
- "把第5页移到第2页"
- "把最后一页移到开头"
- "调整幻灯片顺序"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            fromIndex: {
                type: 'number',
                description: '原位置索引（从1开始）',
            },
            toIndex: {
                type: 'number',
                description: '目标位置索引（从1开始）',
            },
        },
        required: ['fromIndex', 'toIndex'],
    },
};
const moveSlideHandler = async (args) => {
    const { fromIndex, toIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('moveSlide', { fromIndex, toIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片移动成功！\n从第 ${fromIndex} 页 → 移到第 ${toIndex} 页`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `移动幻灯片失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `移动幻灯片出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.moveSlideHandler = moveSlideHandler;
// ============================================================
// 4. wps_ppt_get_slide_count - 获取幻灯片总数
// ============================================================
exports.getSlideCountDefinition = {
    name: 'wps_ppt_get_slide_count',
    description: `获取演示文稿中的幻灯片总数。

使用场景：
- "一共有多少页幻灯片"
- "PPT有几页"
- "查看幻灯片数量"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const getSlideCountHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSlideCount', {}, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `当前演示文稿共有 ${response.data.count} 页幻灯片`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取幻灯片总数失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取幻灯片总数出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getSlideCountHandler = getSlideCountHandler;
// ============================================================
// 5. wps_ppt_get_slide_info - 获取指定幻灯片的详细信息
// ============================================================
exports.getSlideInfoDefinition = {
    name: 'wps_ppt_get_slide_info',
    description: `获取指定幻灯片的详细信息，包括布局、元素列表等。

使用场景：
- "查看第3页的信息"
- "这页有什么内容"
- "获取幻灯片详情"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片索引（从1开始）',
            },
        },
        required: ['slideIndex'],
    },
};
const getSlideInfoHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSlideInfo', { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const info = response.data;
            let output = `第 ${info.slideIndex} 页幻灯片信息：\n`;
            output += `布局: ${info.layout}\n`;
            output += `元素数量: ${info.shapesCount}\n`;
            if (info.shapes && info.shapes.length > 0) {
                output += `\n元素列表：\n`;
                info.shapes.forEach((shape, i) => {
                    output += `  ${i + 1}. [${shape.type}] ${shape.name}`;
                    if (shape.text) {
                        output += ` - "${shape.text.substring(0, 50)}${shape.text.length > 50 ? '...' : ''}"`;
                    }
                    output += '\n';
                });
            }
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
                content: [{ type: 'text', text: `获取幻灯片信息失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取幻灯片信息出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getSlideInfoHandler = getSlideInfoHandler;
// ============================================================
// 6. wps_ppt_switch_slide - 切换到指定幻灯片
// ============================================================
exports.switchSlideDefinition = {
    name: 'wps_ppt_switch_slide',
    description: `切换到指定的幻灯片页面。

使用场景：
- "切换到第5页"
- "跳到最后一页"
- "显示第1页"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '目标幻灯片索引（从1开始）',
            },
        },
        required: ['slideIndex'],
    },
};
const switchSlideHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('switchSlide', { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `已切换到第 ${slideIndex} 页幻灯片`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `切换幻灯片失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `切换幻灯片出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.switchSlideHandler = switchSlideHandler;
// ============================================================
// 7. wps_ppt_set_slide_layout - 设置幻灯片版式布局
// ============================================================
exports.setSlideLayoutDefinition = {
    name: 'wps_ppt_set_slide_layout',
    description: `设置幻灯片的版式布局。

支持的布局类型：
- title: 标题页
- title_content: 标题+内容
- blank: 空白页
- two_column: 两栏内容
- comparison: 对比布局
- section_header: 节标题
- title_only: 仅标题

使用场景：
- "把这页改成空白布局"
- "设置为两栏内容"
- "改成标题页版式"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片索引（从1开始）',
            },
            layout: {
                type: 'string',
                description: '布局名称',
            },
        },
        required: ['slideIndex', 'layout'],
    },
};
const setSlideLayoutHandler = async (args) => {
    const { slideIndex, layout } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSlideLayout', { slideIndex, layout }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const layoutName = {
                title: '标题页',
                title_content: '标题+内容',
                blank: '空白页',
                two_column: '两栏内容',
                comparison: '对比布局',
                section_header: '节标题',
                title_only: '仅标题',
            };
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片布局设置成功！\n幻灯片: 第 ${slideIndex} 页\n布局: ${layoutName[layout] || layout}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置幻灯片布局失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置幻灯片布局出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setSlideLayoutHandler = setSlideLayoutHandler;
// ============================================================
// 8. wps_ppt_get_slide_notes - 获取幻灯片备注内容
// ============================================================
exports.getSlideNotesDefinition = {
    name: 'wps_ppt_get_slide_notes',
    description: `获取指定幻灯片的备注内容。

使用场景：
- "查看第3页的备注"
- "读取演讲备注"
- "这页有什么备注"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片索引（从1开始）',
            },
        },
        required: ['slideIndex'],
    },
};
const getSlideNotesHandler = async (args) => {
    const { slideIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSlideNotes', { slideIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const notes = response.data.notes;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: notes
                            ? `第 ${slideIndex} 页备注内容：\n${notes}`
                            : `第 ${slideIndex} 页没有备注内容`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取幻灯片备注失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取幻灯片备注出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getSlideNotesHandler = getSlideNotesHandler;
// ============================================================
// 9. wps_ppt_set_slide_notes - 设置幻灯片备注
// ============================================================
exports.setSlideNotesDefinition = {
    name: 'wps_ppt_set_slide_notes',
    description: `设置幻灯片的备注内容，用于演讲提示。

使用场景：
- "给第1页添加备注"
- "写一些演讲提示"
- "修改备注内容"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片索引（从1开始）',
            },
            notes: {
                type: 'string',
                description: '备注内容',
            },
        },
        required: ['slideIndex', 'notes'],
    },
};
const setSlideNotesHandler = async (args) => {
    const { slideIndex, notes } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSlideNotes', { slideIndex, notes }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片备注设置成功！\n幻灯片: 第 ${slideIndex} 页\n备注内容: "${notes.substring(0, 80)}${notes.length > 80 ? '...' : ''}"`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置幻灯片备注失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置幻灯片备注出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setSlideNotesHandler = setSlideNotesHandler;
// ============================================================
// 10. wps_ppt_add_shape - 添加形状
// ============================================================
exports.addShapeDefinition = {
    name: 'wps_ppt_add_shape',
    description: `在幻灯片中添加形状。

支持的形状类型：
- rectangle: 矩形
- oval: 椭圆
- triangle: 三角形
- diamond: 菱形
- pentagon: 五边形
- hexagon: 六边形
- arrow: 箭头
- star: 星形
- heart: 心形
- cloud: 云形

使用场景：
- "在第1页添加一个矩形"
- "插入一个蓝色的圆形"
- "添加一个带文字的箭头"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片索引（从1开始），默认1',
            },
            type: {
                type: 'string',
                description: '形状类型',
                enum: ['rectangle', 'oval', 'triangle', 'diamond', 'pentagon', 'hexagon', 'arrow', 'star', 'heart', 'cloud'],
            },
            left: {
                type: 'number',
                description: '左边距（像素），默认100',
            },
            top: {
                type: 'number',
                description: '上边距（像素），默认100',
            },
            width: {
                type: 'number',
                description: '宽度（像素），默认100',
            },
            height: {
                type: 'number',
                description: '高度（像素），默认100',
            },
            text: {
                type: 'string',
                description: '形状内的文本',
            },
            fillColor: {
                type: 'string',
                description: '填充颜色，十六进制如 #FF0000',
            },
        },
        required: [],
    },
};
const addShapeHandler = async (args) => {
    const { slideIndex, type, left, top, width, height, text, fillColor } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addShape', {
            slideIndex: slideIndex || 1,
            type: type || 'rectangle',
            left: left || 100,
            top: top || 100,
            width: width || 100,
            height: height || 100,
            text,
            fillColor,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const shapeNameMap = {
                rectangle: '矩形', oval: '椭圆', triangle: '三角形', diamond: '菱形',
                pentagon: '五边形', hexagon: '六边形', arrow: '箭头', star: '星形',
                heart: '心形', cloud: '云形',
            };
            const typeName = shapeNameMap[type || 'rectangle'] || type;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `形状添加成功！\n幻灯片: 第 ${response.data.slideIndex} 页\n形状类型: ${typeName}\n形状名称: ${response.data.name}${text ? `\n文本: ${text}` : ''}${fillColor ? `\n填充色: ${fillColor}` : ''}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `添加形状失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `添加形状出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.addShapeHandler = addShapeHandler;
// ============================================================
// 11. wps_ppt_set_shape_style - 设置形状样式
// ============================================================
exports.setShapeStyleDefinition = {
    name: 'wps_ppt_set_shape_style',
    description: `设置幻灯片中形状的样式，包括填充颜色、边框颜色和边框粗细。

使用场景：
- "把矩形改成红色"
- "设置形状的边框为蓝色"
- "修改形状样式"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片索引（从1开始），默认1',
            },
            name: {
                type: 'string',
                description: '形状名称（通过getSlideInfo获取）',
            },
            shapeIndex: {
                type: 'number',
                description: '形状索引（与name二选一）',
            },
            fillColor: {
                type: 'string',
                description: '填充颜色，十六进制如 #FF0000',
            },
            lineColor: {
                type: 'string',
                description: '边框颜色，十六进制如 #000000',
            },
            lineWidth: {
                type: 'number',
                description: '边框粗细（磅）',
            },
        },
        required: [],
    },
};
const setShapeStyleHandler = async (args) => {
    const { slideIndex, name, shapeIndex, fillColor, lineColor, lineWidth } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setShapeStyle', {
            slideIndex: slideIndex || 1,
            name,
            shapeIndex,
            fillColor,
            lineColor,
            lineWidth,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            let output = `形状样式设置成功！\n形状: ${response.data.name}`;
            if (fillColor)
                output += `\n填充色: ${fillColor}`;
            if (lineColor)
                output += `\n边框色: ${lineColor}`;
            if (lineWidth)
                output += `\n边框粗细: ${lineWidth}pt`;
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
                content: [{ type: 'text', text: `设置形状样式失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置形状样式出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setShapeStyleHandler = setShapeStyleHandler;
// ============================================================
// 12. wps_ppt_add_textbox - 添加文本框
// ============================================================
exports.addTextboxDefinition = {
    name: 'wps_ppt_add_textbox',
    description: `在幻灯片中添加文本框。

使用场景：
- "在第1页添加一个文本框"
- "插入一个写着标题的文本框"
- "添加文本框并设置字号"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片索引（从1开始），默认1',
            },
            left: {
                type: 'number',
                description: '左边距（像素），默认100',
            },
            top: {
                type: 'number',
                description: '上边距（像素），默认100',
            },
            width: {
                type: 'number',
                description: '宽度（像素），默认200',
            },
            height: {
                type: 'number',
                description: '高度（像素），默认50',
            },
            text: {
                type: 'string',
                description: '文本框内容',
            },
            fontSize: {
                type: 'number',
                description: '字号大小',
            },
            fontName: {
                type: 'string',
                description: '字体名称，如 "微软雅黑"',
            },
        },
        required: [],
    },
};
const addTextboxHandler = async (args) => {
    const { slideIndex, left, top, width, height, text, fontSize, fontName } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('addTextBox', {
            slideIndex: slideIndex || 1,
            left: left || 100,
            top: top || 100,
            width: width || 200,
            height: height || 50,
            text,
            fontSize,
            fontName,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            // Report what was asked for: the action's payload does not always carry these back, and the
            // message used to print "第 undefined 页 / 名称: undefined".
            const targetSlide = response.data?.slideIndex ?? slideIndex ?? 1;
            const shapeName = response.data?.name ? `\n名称: ${response.data.name}` : '';
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `文本框添加成功！\n幻灯片: 第 ${targetSlide} 页${shapeName}${text ? `\n内容: "${text}"` : ''}${fontSize ? `\n字号: ${fontSize}` : ''}${fontName ? `\n字体: ${fontName}` : ''}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `添加文本框失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `添加文本框出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.addTextboxHandler = addTextboxHandler;
// ============================================================
// 13. wps_ppt_set_slide_title - 设置幻灯片标题
// ============================================================
exports.setSlideTitleDefinition = {
    name: 'wps_ppt_set_slide_title',
    description: `设置幻灯片的标题文本。

使用场景：
- "把第1页标题改成'年度总结'"
- "设置标题为'项目进展'"
- "修改幻灯片标题"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片索引（从1开始），默认1',
            },
            title: {
                type: 'string',
                description: '标题文本',
            },
        },
        required: ['title'],
    },
};
const setSlideTitleHandler = async (args) => {
    const { slideIndex, title } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSlideTitle', {
            slideIndex: slideIndex || 1,
            title,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片标题设置成功！\n幻灯片: 第 ${response.data.slideIndex} 页\n标题: "${response.data.title}"`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置幻灯片标题失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置幻灯片标题出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setSlideTitleHandler = setSlideTitleHandler;
// ============================================================
// 14. wps_ppt_insert_image - 插入图片
// ============================================================
// ============================================================
// 15. wps_ppt_set_shape_text - 设置形状文字
// ============================================================
exports.setShapeTextDefinition = {
    name: 'wps_ppt_set_shape_text',
    description: `设置幻灯片中指定形状的文字内容。

使用场景：
- "把第1页的第2个形状文字改成'销售额'"
- "修改形状里的文字"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
            shapeIndex: { type: 'number', description: '形状索引（从1开始）' },
            text: { type: 'string', description: '要设置的文字内容' },
        },
        required: ['slideIndex', 'shapeIndex', 'text'],
    },
};
const setShapeTextHandler = async (args) => {
    const { slideIndex, shapeIndex, text } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setShapeText', { slideIndex, shapeIndex, text }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: `形状文字设置成功！\n形状: ${response.data.name}\n文字: "${text}"` }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置形状文字失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置形状文字出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setShapeTextHandler = setShapeTextHandler;
// ============================================================
// 16. wps_ppt_set_animation - 设置元素动画
// ============================================================
// ============================================================
// 17. wps_ppt_set_background - 设置幻灯片背景
// ============================================================
// ============================================================
// 18. wps_ppt_set_slide_size - 设置幻灯片尺寸
// ============================================================
exports.setSlideSizeDefinition = {
    name: 'wps_ppt_set_slide_size',
    description: `设置演示文稿的幻灯片尺寸。

常用尺寸：
- 标准(4:3): 宽960, 高720
- 宽屏(16:9): 宽960, 高540
- 宽屏(16:10): 宽960, 高600
- A4横版: 宽1123, 高794
- A4竖版: 宽794, 高1123

使用场景：
- "把PPT改成16:9宽屏"
- "设置幻灯片为A4尺寸"
- "修改幻灯片大小"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            width: {
                type: 'number',
                description: '幻灯片宽度（像素）',
            },
            height: {
                type: 'number',
                description: '幻灯片高度（像素）',
            },
        },
        required: ['width', 'height'],
    },
};
const setSlideSizeHandler = async (args) => {
    const { width, height } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSlideSize', // NOTE: macOS未实现，仅Windows支持
        { width, height }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片尺寸设置成功！\n宽度: ${width}px\n高度: ${height}px`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置幻灯片尺寸失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置幻灯片尺寸出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setSlideSizeHandler = setSlideSizeHandler;
// ============================================================
// 19. wps_ppt_set_transition - 设置幻灯片切换效果
// ============================================================
// ============================================================
// 20. wps_ppt_add_chart - 在幻灯片中插入图表
// ============================================================
// ============================================================
// 21. wps_ppt_set_shape_fill - 设置形状填充颜色
// ============================================================
exports.setShapeFillDefinition = {
    name: 'wps_ppt_set_shape_fill',
    description: `设置幻灯片中指定形状的填充颜色。

使用场景：
- "把第1页的第2个形状填充为红色"
- "修改形状背景色为#00FF00"
- "设置形状的填充颜色"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
            shapeIndex: { type: 'number', description: '形状索引（从1开始）' },
            color: { type: 'string', description: '填充颜色，十六进制如 #FF0000' },
        },
        required: ['slideIndex', 'shapeIndex', 'color'],
    },
};
const setShapeFillHandler = async (args) => {
    const { slideIndex, shapeIndex, color } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setShapeFill', // NOTE: macOS未实现，仅Windows支持
        { slideIndex, shapeIndex, color }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `形状填充颜色设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n填充颜色: ${color}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `设置形状填充失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `设置形状填充出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.setShapeFillHandler = setShapeFillHandler;
// ============================================================
// 22. wps_ppt_add_speaker_notes - 添加演讲者备注
// ============================================================
// ============================================================
// 导出所有幻灯片操作相关的Tools
// ============================================================
exports.slideOpsTools = [
    { definition: exports.deleteSlideDefinition, handler: exports.deleteSlideHandler },
    { definition: exports.moveSlideDefinition, handler: exports.moveSlideHandler },
    { definition: exports.getSlideCountDefinition, handler: exports.getSlideCountHandler },
    { definition: exports.getSlideInfoDefinition, handler: exports.getSlideInfoHandler },
    { definition: exports.switchSlideDefinition, handler: exports.switchSlideHandler },
    { definition: exports.setSlideLayoutDefinition, handler: exports.setSlideLayoutHandler },
    { definition: exports.getSlideNotesDefinition, handler: exports.getSlideNotesHandler },
    { definition: exports.setSlideNotesDefinition, handler: exports.setSlideNotesHandler },
    { definition: exports.addShapeDefinition, handler: exports.addShapeHandler },
    { definition: exports.setShapeStyleDefinition, handler: exports.setShapeStyleHandler },
    { definition: exports.addTextboxDefinition, handler: exports.addTextboxHandler },
    { definition: exports.setSlideTitleDefinition, handler: exports.setSlideTitleHandler },
    { definition: exports.setShapeTextDefinition, handler: exports.setShapeTextHandler },
    { definition: exports.setSlideSizeDefinition, handler: exports.setSlideSizeHandler },
    { definition: exports.setShapeFillDefinition, handler: exports.setShapeFillHandler },
];
exports.default = exports.slideOpsTools;
//# sourceMappingURL=slide-ops.js.map
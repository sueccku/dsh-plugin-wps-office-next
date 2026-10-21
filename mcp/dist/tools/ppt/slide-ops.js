"use strict";
/**
 * Input: 幻灯片操作工具参数
 * Output: 幻灯片操作结果
 * Pos: PPT 幻灯片操作工具实现。一旦我被修改，请更新我的头部注释，并在 docs/FIXES.md 记账。
 *
 * 幻灯片操作Tools - 幻灯片生命周期（删除/移动/查询/切换/布局/备注/尺寸）
 * 与形状、文本框的创建（形状的读取/修改族在 shape-basic.ts / textbox.ts）
 *
 * 包含（15 个）：
 * - wps_ppt_delete_slide: 删除指定幻灯片
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
 * - wps_ppt_set_shape_text: 设置形状文字
 * - wps_ppt_set_slide_size: 设置幻灯片尺寸
 * - wps_ppt_set_shape_fill: 设置形状填充颜色
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
// 2. wps_ppt_move_slide - 移动幻灯片到指定位置
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
// 3. wps_ppt_get_slide_count - 获取幻灯片总数
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
// 4. wps_ppt_get_slide_info - 获取指定幻灯片的详细信息
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
            let output = '';
            // A4（改进指南 §9 P1 + §7.2）：目标身份与画布尺寸都要出现在模型看得到的那段文本里 ——
            // 桥返回了但 handler 不打印，等于这个能力对模型不存在。画布用「磅（像素）」同时给出：
            // 形状几何走磅，set_slide_size / export_slide_as_image 收像素，差 0.75 倍。
            if (info.presentationName) {
                const onDisk = typeof info.presentationPath === 'string' && info.presentationPath.includes('\\');
                output += `演示文稿: ${info.presentationName}\n`;
                output += `演示文稿路径: ${onDisk ? info.presentationPath : '(尚未保存到磁盘)'}\n`;
            }
            if (typeof info.slideWidth === 'number' && typeof info.slideHeight === 'number') {
                const px = typeof info.slideWidthPx === 'number' && typeof info.slideHeightPx === 'number'
                    ? `（${info.slideWidthPx} × ${info.slideHeightPx} 像素）` : '';
                output += `画布尺寸: ${info.slideWidth} × ${info.slideHeight} 磅${px}\n`;
            }
            output += `第 ${info.slideIndex} 页幻灯片信息：\n`;
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
// 5. wps_ppt_switch_slide - 切换到指定幻灯片
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
// 6. wps_ppt_set_slide_layout - 设置幻灯片版式布局
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
// 7. wps_ppt_get_slide_notes - 获取幻灯片备注内容
// ============================================================
exports.getSlideNotesDefinition = {
    name: 'wps_ppt_get_slide_notes',
    description: `获取幻灯片备注内容。传 all:true 返回全部页与 hasNotes，用于"40 页里哪些没备注"的体检。

使用场景：
- "查看第3页的备注"
- "读取演讲备注"
- "这页有什么备注"
- "哪些页没有备注"（all:true）`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            slideIndex: {
                type: 'number',
                description: '幻灯片索引（从1开始）；all:true 时忽略',
            },
            all: {
                type: 'boolean',
                description: 'true 时返回全部幻灯片的备注与 hasNotes',
            },
        },
        required: [],
    },
};
const getSlideNotesHandler = async (args) => {
    const { slideIndex, all } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSlideNotes', all ? { all: true } : { slideIndex: slideIndex ?? 1 }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            if (all) {
                const d = response.data;
                return {
                    id: (0, uuid_1.v4)(),
                    success: true,
                    content: [{
                            type: 'text',
                            text: `备注体检: ${d.withNotes ?? 0}/${d.total ?? 0} 页有备注\n${(d.slides || []).map((s) => `第${s.slideIndex}页: ${s.hasNotes ? s.notes : '(无备注)'}`).join('\n')}`,
                        }],
                };
            }
            const notes = response.data.notes || '';
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: notes
                            ? `第 ${slideIndex ?? 1} 页备注内容：\n${notes}`
                            : `第 ${slideIndex ?? 1} 页没有备注内容`,
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
// 8. wps_ppt_set_slide_notes - 设置幻灯片备注
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
// 9. wps_ppt_add_shape - 添加形状
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
// 10. wps_ppt_set_shape_style - 设置形状样式
// ============================================================
exports.setShapeStyleDefinition = {
    name: 'wps_ppt_set_shape_style',
    description: `设置幻灯片中形状的样式，包括填充颜色、边框颜色、边框粗细和圆角半径。

使用场景：
- "把矩形改成红色"
- "设置形状的边框为蓝色"
- "把圆角调大一点"（roundness，仅带调整柄的形状有效；结果会回传实际生效值）
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
            roundness: {
                type: 'number',
                description: '圆角半径（0-0.5，越大越圆；仅圆角矩形这类带调整柄的形状支持，其它形状会明确报错）',
            },
        },
        required: [],
    },
};
const setShapeStyleHandler = async (args) => {
    const { slideIndex, name, shapeIndex, fillColor, lineColor, lineWidth, roundness } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setShapeStyle', {
            slideIndex: slideIndex || 1,
            name,
            shapeIndex,
            fillColor,
            lineColor,
            lineWidth,
            roundness,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            let output = `形状样式设置成功！\n形状: ${response.data.name}`;
            if (fillColor)
                output += `\n填充色: ${fillColor}`;
            if (lineColor)
                output += `\n边框色: ${lineColor}`;
            if (lineWidth)
                output += `\n边框粗细: ${lineWidth}pt`;
            if (roundness !== undefined)
                output += `\n圆角半径: ${roundness}`;
            if (response.data.roundness !== undefined && response.data.roundness !== null) {
                output += `\n实际圆角半径: ${response.data.roundness}`;
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
// 11. wps_ppt_add_textbox - 添加文本框
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
// 12. wps_ppt_set_slide_title - 设置幻灯片标题
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
// 13. wps_ppt_set_shape_text - 设置形状文字
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
// 14. wps_ppt_set_slide_size - 设置幻灯片尺寸
// ============================================================
exports.setSlideSizeDefinition = {
    name: 'wps_ppt_set_slide_size',
    description: `设置演示文稿的幻灯片尺寸。width / height 的单位是**像素**，桥内部乘以 0.75 换成磅（COM 的 PageSetup.SlideWidth 是磅）。

常用尺寸（像素 → 实际磅值）：
- 标准 4:3: 960 × 720 像素（= 720 × 540 磅）
- 宽屏 16:9: 1280 × 720 像素（= 960 × 540 磅）—— 注意不是 960 × 540：那组数字是**磅**，按像素传会小一圈
- 宽屏 16:10: 1280 × 800 像素（= 960 × 600 磅）
- A4 横版: 1123 × 794 像素（= 842 × 595 磅）
- A4 竖版: 794 × 1123 像素（= 595 × 842 磅）

改完用 wps_ppt_get_slide_info 回读 slideWidth / slideHeight（磅）与 slideWidthPx / slideHeightPx（像素）核对。

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
            scaleContent: {
                type: 'boolean',
                description: '改尺寸时是否按比例缩放页面上的形状（PT3：确保内容不跑出画布）；默认 false',
            },
        },
        required: ['width', 'height'],
    },
};
const setSlideSizeHandler = async (args) => {
    const { width, height, scaleContent } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('setSlideSize', { width, height, scaleContent }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片尺寸设置成功！\n宽度: ${width}px\n高度: ${height}px${response.data?.scaleContent ? `\n已按比例缩放 ${response.data?.affectedSlides ?? 0} 页上的形状` : '\n未缩放页面内容（scaleContent=false）'}`,
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
// 15. wps_ppt_set_shape_fill - 设置形状填充颜色
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
        const response = await wps_client_1.wpsClient.executeMethod('setShapeFill', { slideIndex, shapeIndex, color }, wps_1.WpsAppType.PRESENTATION);
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
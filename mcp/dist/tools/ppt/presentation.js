"use strict";
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
 * - wps_ppt_copy_slide: 复制幻灯片
 * - wps_ppt_insert_slide_image: 在幻灯片中插入图片
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.presentationTools = exports.setActiveTargetHandler = exports.setActiveTargetDefinition = exports.insertSlidesFromFileHandler = exports.insertSlidesFromFileDefinition = exports.insertSlideImageHandler = exports.insertSlideImageDefinition = exports.copySlideHandler = exports.copySlideDefinition = exports.switchPresentationHandler = exports.switchPresentationDefinition = exports.getOpenPresentationsHandler = exports.getOpenPresentationsDefinition = exports.closePresentationHandler = exports.closePresentationDefinition = exports.openPresentationHandler = exports.openPresentationDefinition = exports.createPresentationHandler = exports.createPresentationDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 新建空白演示文稿
 */
exports.createPresentationDefinition = {
    name: 'wps_ppt_create_presentation',
    description: `新建空白演示文稿。

使用场景：
- "新建一个PPT"
- "创建一个演示文稿"
- "打开一个新的PPT"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const createPresentationHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('createPresentation', {}, wps_1.WpsAppType.PRESENTATION);
        // A fresh WPS deck has zero slides and the action reports no name, so the message must not
        // depend on a payload field: it used to print "名称: undefined". Success is what matters.
        if (response.success) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: response.data?.slideCount
                            ? `新建演示文稿成功！\n幻灯片数: ${response.data.slideCount}`
                            : '新建演示文稿成功！（新文稿暂无幻灯片，请先 add_slide）',
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `新建演示文稿失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `新建演示文稿出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.createPresentationHandler = createPresentationHandler;
/**
 * 打开指定路径的演示文稿
 */
exports.openPresentationDefinition = {
    name: 'wps_ppt_open_presentation',
    description: `打开指定路径的演示文稿文件。

支持的文件格式：
- .pptx: PowerPoint 演示文稿
- .ppt: 旧版 PowerPoint 格式
- .dps: WPS 演示格式

使用场景：
- "打开桌面上的演示文稿"
- "打开这个PPT文件"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const openPresentationHandler = async (args) => {
    const { filePath } = args;
    try {
        // 跨平台参数对齐：macOS/Windows 底层均读取 params.path，需同时发送 path 别名
        const response = await wps_client_1.wpsClient.executeMethod('openPresentation', { filePath, path: filePath }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `演示文稿打开成功！\n名称: ${response.data.name}\n路径: ${response.data.filePath}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `打开演示文稿失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `打开演示文稿出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.openPresentationHandler = openPresentationHandler;
/**
 * 关闭演示文稿
 */
exports.closePresentationDefinition = {
    name: 'wps_ppt_close_presentation',
    description: `关闭演示文稿。可指定文稿名称，不指定则关闭当前活动文稿。

使用场景：
- "关闭这个PPT"
- "关闭演示文稿不保存"
- "关闭指定的PPT"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const closePresentationHandler = async (args) => {
    const { name, save } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('closePresentation', {
            name,
            save: save !== false,
        }, wps_1.WpsAppType.PRESENTATION);
        if (response.success) {
            const closed = response.data?.closed ?? name ?? '(当前演示文稿)';
            const saveStatus = response.data?.saved ? '已保存' : '未保存';
            const warnNote = response.data?.warning ? `\n注意: ${response.data.warning}` : '';
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `演示文稿已关闭！\n名称: ${closed}\n保存状态: ${saveStatus}${warnNote}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `关闭演示文稿失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `关闭演示文稿出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.closePresentationHandler = closePresentationHandler;
/**
 * 获取所有已打开的演示文稿列表
 */
exports.getOpenPresentationsDefinition = {
    name: 'wps_ppt_get_open_presentations',
    description: `获取当前所有已打开的演示文稿列表。

使用场景：
- "有哪些PPT打开着"
- "列出所有演示文稿"
- "查看打开的PPT"`,
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {},
        required: [],
    },
};
const getOpenPresentationsHandler = async (_args) => {
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getOpenPresentations', {}, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            const { presentations, count } = response.data;
            if (count === 0) {
                return {
                    id: (0, uuid_1.v4)(),
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
                id: (0, uuid_1.v4)(),
                success: true,
                content: [{ type: 'text', text: output }],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `获取演示文稿列表失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `获取演示文稿列表出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.getOpenPresentationsHandler = getOpenPresentationsHandler;
/**
 * 切换到指定演示文稿
 */
exports.switchPresentationDefinition = {
    name: 'wps_ppt_switch_presentation',
    description: `切换到指定名称的演示文稿。

使用场景：
- "切换到另一个PPT"
- "打开那个叫xxx的演示文稿"
- "切换演示文稿"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const switchPresentationHandler = async (args) => {
    const { name } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('switchPresentation', { name }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `已切换到演示文稿: ${response.data.name}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `切换演示文稿失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `切换演示文稿出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.switchPresentationHandler = switchPresentationHandler;
exports.copySlideDefinition = {
    name: 'wps_ppt_copy_slide',
    description: `复制幻灯片到指定位置。

使用场景：
- "复制第2页幻灯片"
- "把这页复制到第5页后面"
- "克隆当前幻灯片"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const copySlideHandler = async (args) => {
    const { slideIndex, targetIndex } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('duplicateSlide', { slideIndex, targetIndex }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `幻灯片已复制！\n源页: 第${slideIndex}页\n新页位置: 第${response.data.newIndex}页`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `复制幻灯片失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `复制幻灯片出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.copySlideHandler = copySlideHandler;
/**
 * 在幻灯片中插入图片
 */
exports.insertSlideImageDefinition = {
    name: 'wps_ppt_insert_slide_image',
    description: `在幻灯片中插入图片。

使用场景：
- "在第1页插入一张图片"
- "添加图片到幻灯片"
- "把这个图片放到PPT里"`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const insertSlideImageHandler = async (args) => {
    const { slideIndex, imagePath, left, top } = args;
    try {
        // "insertImage" is the Word action, so this used to insert the picture into the Word document
        // instead of the slide. insertPptImage is the presentation one.
        const response = await wps_client_1.wpsClient.executeMethod('insertPptImage', { slideIndex, path: imagePath, left, top }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `图片已插入到第${slideIndex}页幻灯片！\n图片路径: ${imagePath}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `插入图片失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `插入图片出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertSlideImageHandler = insertSlideImageHandler;
/**
 * 从其它演示文稿导入整页幻灯片（跨PPT整合，保留来源格式）
 */
exports.insertSlidesFromFileDefinition = {
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
    category: tools_1.ToolCategory.PRESENTATION,
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
const insertSlidesFromFileHandler = async (args) => {
    const { filePath, afterIndex, slideStart, slideEnd } = args;
    try {
        const response = await wps_client_1.wpsClient.executeMethod('insertSlidesFromFile', { filePath, path: filePath, afterIndex, slideStart, slideEnd }, wps_1.WpsAppType.PRESENTATION);
        if (response.success && response.data) {
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `已从来源文件导入 ${response.data.inserted} 页幻灯片（保留来源格式）！\n来源: ${filePath}\n插入位置: 第 ${response.data.afterIndex} 页之后\n当前演示文稿总页数: ${response.data.totalSlides}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `导入幻灯片失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `导入幻灯片出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.insertSlidesFromFileHandler = insertSlidesFromFileHandler;
/**
 * 锁定后续所有 PPT 操作的目标演示文稿（避免多文稿打开时活动文稿漂移改错文件）
 */
exports.setActiveTargetDefinition = {
    name: 'wps_ppt_set_active_target',
    description: `锁定后续所有 PPT 操作的【目标演示文稿】（按文件名）。锁定后，本服务器会给所有演示类调用自动注入该文稿名，底层精确定位，**彻底避免同时打开多个 PPT 时"活动文稿漂移"导致改错文件**。

使用场景：
- 批量改某个 PPT 前先锁定："把目标锁定为 关节模组申报PPT-生成版.pptx"
- 传 clear=true 或留空 name 取消锁定，恢复使用当前活动文稿

强烈建议：在对一个特定 PPT 做成批增改前，先调用本工具锁定它。`,
    category: tools_1.ToolCategory.PRESENTATION,
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
const setActiveTargetHandler = async (args) => {
    const { name, clear } = args;
    if (clear || !name) {
        (0, wps_client_1.setPptTarget)(undefined);
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [{ type: 'text', text: '已取消目标文稿锁定，后续 PPT 操作使用当前活动文稿。' }],
        };
    }
    try {
        const resp = await wps_client_1.wpsClient.executeMethod('getOpenPresentations', {});
        const names = resp.success && resp.data ? resp.data.presentations.map((x) => x.name) : [];
        const matched = names.includes(name);
        (0, wps_client_1.setPptTarget)(name);
        const warn = matched
            ? ''
            : `\n⚠️ 当前打开的文稿中没有完全同名的「${name}」。已打开：${names.join(' / ') || '(无)'}。若名称不符，后续操作会报"未找到目标文稿"。`;
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [
                {
                    type: 'text',
                    text: `已锁定目标演示文稿：${name}\n后续所有 PPT 操作都会定位到该文稿（不受活动窗口切换影响）。${warn}`,
                },
            ],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        (0, wps_client_1.setPptTarget)(name);
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [{ type: 'text', text: `已锁定目标演示文稿：${name}（打开列表校验失败：${errMsg}）` }],
        };
    }
};
exports.setActiveTargetHandler = setActiveTargetHandler;
/**
 * 导出所有演示文稿管理相关的Tools
 */
exports.presentationTools = [
    { definition: exports.createPresentationDefinition, handler: exports.createPresentationHandler },
    { definition: exports.openPresentationDefinition, handler: exports.openPresentationHandler },
    { definition: exports.closePresentationDefinition, handler: exports.closePresentationHandler },
    { definition: exports.getOpenPresentationsDefinition, handler: exports.getOpenPresentationsHandler },
    { definition: exports.switchPresentationDefinition, handler: exports.switchPresentationHandler },
    { definition: exports.copySlideDefinition, handler: exports.copySlideHandler },
    { definition: exports.insertSlideImageDefinition, handler: exports.insertSlideImageHandler },
    { definition: exports.insertSlidesFromFileDefinition, handler: exports.insertSlidesFromFileHandler },
    { definition: exports.setActiveTargetDefinition, handler: exports.setActiveTargetHandler },
];
exports.default = exports.presentationTools;
//# sourceMappingURL=presentation.js.map
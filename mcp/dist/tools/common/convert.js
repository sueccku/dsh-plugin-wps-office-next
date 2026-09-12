"use strict";
/**
 * Input: 转换工具参数
 * Output: 文档转换结果
 * Pos: 通用文档转换工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 文档转换Tools - 跨应用格式转换模块
 * 负责Word/Excel/PPT文档的格式转换操作
 *
 * 包含：
 * - wps_convert_to_pdf: 转换为PDF格式
 * - wps_convert_format: 格式互转（docx<->doc, xlsx<->xls, pptx<->ppt等）
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.getFormatCode = exports.getAppTypeByExtension = exports.convertTools = exports.convertFormatHandler = exports.convertFormatDefinition = exports.convertToPdfHandler = exports.convertToPdfDefinition = void 0;
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
/**
 * 根据文件扩展名判断应该用哪个WPS应用类型
 */
const getAppTypeByExtension = (filePath) => {
    const ext = filePath.toLowerCase().split('.').pop();
    // Word文档
    if (['doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'rtf', 'wps', 'wpt'].includes(ext || '')) {
        return wps_1.WpsAppType.WRITER;
    }
    // Excel表格
    if (['xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm', 'csv', 'et', 'ett'].includes(ext || '')) {
        return wps_1.WpsAppType.SPREADSHEET;
    }
    // PPT演示
    if (['ppt', 'pptx', 'pptm', 'pot', 'potx', 'potm', 'pps', 'ppsx', 'ppsm', 'dps', 'dpt'].includes(ext || '')) {
        return wps_1.WpsAppType.PRESENTATION;
    }
    return null;
};
exports.getAppTypeByExtension = getAppTypeByExtension;
/**
 * 根据输出格式获取对应的文件格式代码
 * WPS的格式代码和微软Office基本兼容，但也有自己的一套
 */
const getFormatCode = (format, appType) => {
    const formatLower = format.toLowerCase();
    switch (appType) {
        case wps_1.WpsAppType.WRITER:
            // Word文档格式代码
            // wdFormatDocument = 0 (.doc)
            // wdFormatDocumentDefault = 16 (.docx)
            // wdFormatPDF = 17
            // wdFormatRTF = 6
            // wdFormatXPS = 18
            // wdFormatHTML = 8
            const wordFormats = {
                'doc': 0,
                'docx': 16,
                'pdf': 17,
                'rtf': 6,
                'xps': 18,
                'html': 8,
                'htm': 8,
                'txt': 2, // wdFormatText
                'xml': 11, // wdFormatXML
            };
            return wordFormats[formatLower] ?? 16; // 默认docx
        case wps_1.WpsAppType.SPREADSHEET:
            // Excel工作簿格式代码
            // xlWorkbookNormal = -4143 (.xls)
            // xlOpenXMLWorkbook = 51 (.xlsx)
            // xlTypePDF = 0 (导出PDF时用)
            // xlCSV = 6
            // xlHtml = 44
            const excelFormats = {
                'xls': -4143,
                'xlsx': 51,
                'xlsm': 52,
                'xlsb': 50,
                'csv': 6,
                'html': 44,
                'htm': 44,
                'pdf': 0, // 特殊处理
                'xps': 1, // 特殊处理
            };
            return excelFormats[formatLower] ?? 51; // 默认xlsx
        case wps_1.WpsAppType.PRESENTATION:
            // PPT演示格式代码
            // ppSaveAsPresentation = 1 (.ppt)
            // ppSaveAsOpenXMLPresentation = 24 (.pptx)
            // ppSaveAsPDF = 32
            // ppSaveAsHTML = 12
            const pptFormats = {
                'ppt': 1,
                'pptx': 24,
                'pptm': 25,
                'pdf': 32,
                'xps': 33,
                'html': 12,
                'htm': 12,
                'png': 18,
                'jpg': 17,
                'jpeg': 17,
                'gif': 16,
                'bmp': 19,
            };
            return pptFormats[formatLower] ?? 24; // 默认pptx
        default:
            return -1;
    }
};
exports.getFormatCode = getFormatCode;
/**
 * 转换为PDF格式
 * 支持Word、Excel、PPT文档一键转PDF
 */
exports.convertToPdfDefinition = {
    name: 'wps_convert_to_pdf',
    description: `将当前文档转换为PDF格式。支持Word、Excel、PPT文档。

使用场景：
- "把这个Word转成PDF"
- "导出PDF给客户看"
- "把表格保存成PDF格式"
- "PPT转PDF方便打印"

特点：
- 自动检测当前打开的文档类型
- 可以指定输出路径，不指定则使用原文件名.pdf
- 保持原文档格式和排版`,
    category: tools_1.ToolCategory.COMMON,
    inputSchema: {
        type: 'object',
        properties: {
            outputPath: {
                type: 'string',
                description: 'PDF输出路径（包含文件名），如不指定则使用原文件路径，把扩展名改为.pdf',
            },
            openAfterExport: {
                type: 'boolean',
                description: '导出后是否自动打开PDF，默认false',
            },
            app_type: {
                type: 'string',
                enum: ['excel', 'word', 'ppt'],
                description: '要导出的应用；不填则按 Excel→Word→PPT 选第一个正在运行的文档',
            },
        },
        required: [],
    },
};
const convertToPdfHandler = async (args) => {
    const { outputPath, openAfterExport, app_type } = args;
    try {
        // 调用WPS加载项执行转换
        const response = await wps_client_1.wpsClient.executeMethod('convertToPDF', {
            // The bridge reads outputPath. The path/filePath aliases carried the same value and were
            // never read; the parameter guard now rejects them outright.
            outputPath: outputPath || '',
            openAfterExport: openAfterExport || false,
            appType: app_type,
        }
        // 不指定appType，让WPS加载项自动检测当前活动的应用
        );
        if (response.success && response.data) {
            const result = response.data;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `PDF导出成功！
源文件: ${result.sourcePath}
输出路径: ${result.outputPath}
文档类型: ${result.appType}${result.pageCount ? `\n页数: ${result.pageCount}` : ''}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `PDF导出失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `PDF导出出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.convertToPdfHandler = convertToPdfHandler;
/**
 * 格式互转
 * 支持docx<->doc, xlsx<->xls, pptx<->ppt等多种格式转换
 */
exports.convertFormatDefinition = {
    name: 'wps_convert_format',
    description: `将当前文档转换为其他格式。

使用场景：
- "把docx转成doc格式"
- "保存为旧版Excel格式兼容老系统"
- "把PPT转成pptx"
- "导出为RTF格式"
- "转换成HTML网页格式"

支持的格式：
- Word: doc, docx, rtf, txt, html, xml
- Excel: xls, xlsx, xlsm, xlsb, csv, html
- PPT: ppt, pptx, pptm, html, png, jpg, gif, bmp

注意：
- 转换时会保留原文档，另存为新格式
- 某些格式转换可能会丢失部分效果（如doc不支持的新特性）`,
    category: tools_1.ToolCategory.COMMON,
    inputSchema: {
        type: 'object',
        properties: {
            targetFormat: {
                type: 'string',
                description: '目标格式扩展名，如 doc, xlsx, ppt, rtf, csv, html 等',
            },
            outputPath: {
                type: 'string',
                description: '输出路径（包含文件名），如不指定则使用原文件名改为新扩展名',
            },
            app_type: {
                type: 'string',
                enum: ['excel', 'word', 'ppt'],
                description: '要转换的应用；不填则按 Excel→Word→PPT 选第一个正在运行的文档',
            },
        },
        required: ['targetFormat'],
    },
};
const convertFormatHandler = async (args) => {
    const { targetFormat, outputPath, app_type } = args;
    if (!targetFormat || targetFormat.trim() === '') {
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: '目标格式不能为空，请指定目标格式（如 doc、xlsx、pdf 等）' }],
            error: '目标格式为空',
        };
    }
    try {
        // 调用WPS加载项执行格式转换
        const response = await wps_client_1.wpsClient.executeMethod('convertFormat', {
            targetFormat: targetFormat.toLowerCase().replace(/^\./, ''), // 去掉开头的点
            outputPath: outputPath || '',
            appType: app_type,
        }
        // 不指定appType，让WPS加载项自动检测
        );
        if (response.success && response.data) {
            const result = response.data;
            return {
                id: (0, uuid_1.v4)(),
                success: true,
                content: [
                    {
                        type: 'text',
                        text: `格式转换成功！
源文件: ${result.sourcePath}
源格式: ${result.sourceFormat}
目标格式: ${result.targetFormat}
输出路径: ${result.outputPath}`,
                    },
                ],
            };
        }
        else {
            return {
                id: (0, uuid_1.v4)(),
                success: false,
                content: [{ type: 'text', text: `格式转换失败: ${response.error}` }],
                error: response.error,
            };
        }
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return {
            id: (0, uuid_1.v4)(),
            success: false,
            content: [{ type: 'text', text: `格式转换出错: ${errMsg}` }],
            error: errMsg,
        };
    }
};
exports.convertFormatHandler = convertFormatHandler;
/**
 * 导出所有转换相关的Tools
 */
exports.convertTools = [
    { definition: exports.convertToPdfDefinition, handler: exports.convertToPdfHandler },
    { definition: exports.convertFormatDefinition, handler: exports.convertFormatHandler },
];
exports.default = exports.convertTools;
//# sourceMappingURL=convert.js.map
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.pptExportTools = exports.exportNotesHandler = exports.exportNotesDefinition = void 0;
/**
 * Input: 导出备注的路径/格式
 * Output: 备注文本文件
 * Pos: P1-7（2026-10-21）备注/讲义门面的第一步：把全部备注导出成 text/markdown。
 *      不新增桥 action：复用 getSlideNotes({all:true})，文件写入由 MCP 进程完成。
 */
const node_fs_1 = require("node:fs");
const node_path_1 = require("node:path");
const uuid_1 = require("uuid");
const tools_1 = require("../../types/tools");
const wps_client_1 = require("../../client/wps-client");
const wps_1 = require("../../types/wps");
exports.exportNotesDefinition = {
    name: 'wps_ppt_export_notes',
    description: '把演示文稿全部演讲者备注导出为 text/markdown 文件（每页一节，空备注跳过）。用于交付讲稿/提词稿；不做 Word 排版。',
    category: tools_1.ToolCategory.PRESENTATION,
    inputSchema: {
        type: 'object',
        properties: {
            outputPath: { type: 'string', description: '输出文件路径（.md 或 .txt）' },
            format: { type: 'string', enum: ['markdown', 'text'], description: 'markdown（默认）或 text' },
        },
        required: ['outputPath'],
    },
};
const exportNotesHandler = async (args) => {
    const { outputPath, format } = args;
    if (!outputPath || !outputPath.trim()) {
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: 'outputPath 不能为空' }], error: '缺少输出路径' };
    }
    try {
        const response = await wps_client_1.wpsClient.executeMethod('getSlideNotes', { all: true }, wps_1.WpsAppType.PRESENTATION);
        if (!response.success) {
            return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `读取备注失败: ${response.error}` }], error: response.error };
        }
        const slides = response.data?.slides || [];
        const isMarkdown = (format || 'markdown') !== 'text';
        const lines = [];
        if (isMarkdown)
            lines.push('# 演讲备注\n');
        let exported = 0;
        for (const slide of slides) {
            if (!slide.hasNotes)
                continue;
            exported++;
            const title = isMarkdown ? `## 第 ${slide.slideIndex} 页` : `== 第 ${slide.slideIndex} 页 ==`;
            lines.push(title, '', slide.notes.trim(), '');
        }
        if (exported === 0) {
            return { id: (0, uuid_1.v4)(), success: true, content: [{ type: 'text', text: `没有可导出的备注（共 ${slides.length} 页）` }] };
        }
        const target = (0, node_path_1.resolve)(outputPath);
        (0, node_fs_1.mkdirSync)((0, node_path_1.dirname)(target), { recursive: true });
        (0, node_fs_1.writeFileSync)(target, lines.join('\n'), 'utf8');
        return {
            id: (0, uuid_1.v4)(),
            success: true,
            content: [{ type: 'text', text: `备注已导出: ${target}\n有备注 ${exported}/${slides.length} 页` }],
        };
    }
    catch (error) {
        const errMsg = error instanceof Error ? error.message : String(error);
        return { id: (0, uuid_1.v4)(), success: false, content: [{ type: 'text', text: `导出备注出错: ${errMsg}` }], error: errMsg };
    }
};
exports.exportNotesHandler = exportNotesHandler;
exports.pptExportTools = [
    { definition: exports.exportNotesDefinition, handler: exports.exportNotesHandler },
];
//# sourceMappingURL=export-notes.js.map
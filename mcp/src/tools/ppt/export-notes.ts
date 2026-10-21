/**
 * Input: 导出备注的路径/格式
 * Output: 备注文本文件
 * Pos: P1-7（2026-10-21）备注/讲义门面的第一步：把全部备注导出成 text/markdown。
 *      不新增桥 action：复用 getSlideNotes({all:true})，文件写入由 MCP 进程完成。
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
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

export const exportNotesDefinition: ToolDefinition = {
  name: 'wps_ppt_export_notes',
  description: '把演示文稿全部演讲者备注导出为 text/markdown 文件（每页一节，空备注跳过）。用于交付讲稿/提词稿；不做 Word 排版。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      outputPath: { type: 'string', description: '输出文件路径（.md 或 .txt）' },
      format: { type: 'string', enum: ['markdown', 'text'], description: 'markdown（默认）或 text' },
    },
    required: ['outputPath'],
  },
};

export const exportNotesHandler: ToolHandler = async (args: Record<string, unknown>): Promise<ToolCallResult> => {
  const { outputPath, format } = args as { outputPath?: string; format?: string };
  if (!outputPath || !outputPath.trim()) {
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: 'outputPath 不能为空' }], error: '缺少输出路径' };
  }
  try {
    const response = await wpsClient.executeMethod<{
      total?: number;
      withNotes?: number;
      slides?: Array<{ slideIndex: number; notes: string; hasNotes: boolean }>;
    }>('getSlideNotes', { all: true }, WpsAppType.PRESENTATION);
    if (!response.success) {
      return { id: uuidv4(), success: false, content: [{ type: 'text', text: `读取备注失败: ${response.error}` }], error: response.error };
    }
    const slides = response.data?.slides || [];
    const isMarkdown = (format || 'markdown') !== 'text';
    const lines: string[] = [];
    if (isMarkdown) lines.push('# 演讲备注\n');
    let exported = 0;
    for (const slide of slides) {
      if (!slide.hasNotes) continue;
      exported++;
      const title = isMarkdown ? `## 第 ${slide.slideIndex} 页` : `== 第 ${slide.slideIndex} 页 ==`;
      lines.push(title, '', slide.notes.trim(), '');
    }
    if (exported === 0) {
      return { id: uuidv4(), success: true, content: [{ type: 'text', text: `没有可导出的备注（共 ${slides.length} 页）` }] };
    }
    const target = resolve(outputPath);
    mkdirSync(dirname(target), { recursive: true });
    writeFileSync(target, lines.join('\n'), 'utf8');
    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: `备注已导出: ${target}\n有备注 ${exported}/${slides.length} 页` }],
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: `导出备注出错: ${errMsg}` }], error: errMsg };
  }
};

export const pptExportTools: RegisteredTool[] = [
  { definition: exportNotesDefinition, handler: exportNotesHandler },
];

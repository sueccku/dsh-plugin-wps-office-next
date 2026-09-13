/**
 * Input: 文档管理工具参数
 * Output: 文档操作结果
 * Pos: Word 文档管理工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word文档管理Tools
 * 处理文档的打开、切换、获取列表等管理操作
 *
 * 包含：
 * - wps_word_get_open_documents: 获取所有已打开的文档列表
 * - wps_word_switch_document: 切换到指定文档
 * - wps_word_open_document: 打开指定路径的文档
 * - wps_word_create_document: 新建空白文档
 * - wps_word_close_document: 关闭文档（可选保存）
 * - wps_word_get_document_text: 获取文档文本内容
 * - wps_word_insert_header: 设置页眉内容
 * - wps_word_insert_footer: 设置页脚内容
 * - wps_word_generate_doc_toc: 自动生成文档目录
 * - wps_word_insert_section_break: 插入分节符
 * - wps_word_set_line_spacing: 设置行距
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 获取所有已打开的文档列表
 */
export declare const getOpenDocumentsDefinition: ToolDefinition;
export declare const getOpenDocumentsHandler: ToolHandler;
/**
 * 切换到指定文档
 */
export declare const switchDocumentDefinition: ToolDefinition;
export declare const switchDocumentHandler: ToolHandler;
/**
 * 打开指定路径的文档
 */
export declare const openDocumentDefinition: ToolDefinition;
export declare const openDocumentHandler: ToolHandler;
/**
 * 获取文档文本内容
 */
export declare const getDocumentTextDefinition: ToolDefinition;
export declare const getDocumentTextHandler: ToolHandler;
/**
 * 设置页眉内容
 */
export declare const insertHeaderDefinition: ToolDefinition;
export declare const insertHeaderHandler: ToolHandler;
/**
 * 设置页脚内容
 */
export declare const insertFooterDefinition: ToolDefinition;
export declare const insertFooterHandler: ToolHandler;
/**
 * 自动生成文档目录
 */
export declare const generateDocTocDefinition: ToolDefinition;
export declare const generateDocTocHandler: ToolHandler;
/**
 * 插入分节符
 */
export declare const insertSectionBreakDefinition: ToolDefinition;
export declare const insertSectionBreakHandler: ToolHandler;
/**
 * 设置行距
 */
export declare const setLineSpacingDefinition: ToolDefinition;
export declare const setLineSpacingHandler: ToolHandler;
/**
 * 新建空白文档
 */
export declare const createDocumentDefinition: ToolDefinition;
export declare const createDocumentHandler: ToolHandler;
/**
 * 关闭文档
 */
export declare const closeDocumentDefinition: ToolDefinition;
export declare const closeDocumentHandler: ToolHandler;
/**
 * 导出所有文档管理相关的Tools
 */
export declare const documentTools: RegisteredTool[];
export default documentTools;
//# sourceMappingURL=document.d.ts.map
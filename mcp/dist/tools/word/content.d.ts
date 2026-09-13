/**
 * Input: Word 内容操作参数
 * Output: 文档内容变更结果
 * Pos: Word 内容工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word内容操作Tools - 文档内容编辑模块
 * 处理文档内容的插入、查找替换、表格、段落格式、页眉页脚等操作
 *
 * 包含：
 * - wps_word_insert_text: 插入文本到文档
 * - wps_word_find_replace: 查找替换功能
 * - wps_word_insert_table: 插入表格
 * - wps_word_set_paragraph: 设置段落格式
 * - wps_word_insert_header: 插入页眉
 * - wps_word_insert_footer: 插入页脚
 * - wps_word_get_active_document: 获取当前文档信息
 * - wps_word_insert_page_break: 插入分页符
 * - wps_word_insert_comment: 插入批注
 * - wps_word_set_font_style: 设置选中文字的字体样式属性
 * - wps_word_set_text_color: 设置文字颜色
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 插入文本到文档
 * 可以在光标位置、文档开头或结尾插入文本
 */
export declare const insertTextDefinition: ToolDefinition;
export declare const insertTextHandler: ToolHandler;
/**
 * 查找替换功能
 * 支持全文批量查找替换文本
 */
export declare const findReplaceDefinition: ToolDefinition;
export declare const findReplaceHandler: ToolHandler;
/**
 * 插入表格到文档光标位置
 */
export declare const insertTableDefinition: ToolDefinition;
export declare const insertTableHandler: ToolHandler;
/**
 * 设置当前段落格式
 */
export declare const setParagraphDefinition: ToolDefinition;
export declare const setParagraphHandler: ToolHandler;
/**
 * 获取当前活动文档信息
 */
export declare const getActiveDocumentDefinition: ToolDefinition;
export declare const getActiveDocumentHandler: ToolHandler;
/**
 * 在文档中插入图片
 */
export declare const insertImageDefinition: ToolDefinition;
export declare const insertImageHandler: ToolHandler;
/**
 * 导出所有内容操作相关的Tools
 */
export declare const insertPageBreakDefinition: ToolDefinition;
export declare const insertPageBreakHandler: ToolHandler;
/**
 * 插入批注到文档选中内容
 */
export declare const insertCommentDefinition: ToolDefinition;
export declare const insertCommentHandler: ToolHandler;
/**
 * 设置选中文字的颜色
 */
export declare const setTextColorDefinition: ToolDefinition;
export declare const setTextColorHandler: ToolHandler;
/**
 * 获取文档段落结构
 * 返回段落的文本、样式、位置信息，用于了解文档结构和识别填写位置
 */
export declare const getParagraphsDefinition: ToolDefinition;
export declare const getParagraphsHandler: ToolHandler;
/**
 * 查找文本返回位置（不替换）
 * 返回匹配文本的位置信息和上下文，供后续精确操作使用
 */
export declare const findInDocumentDefinition: ToolDefinition;
export declare const findInDocumentHandler: ToolHandler;
/**
 * 智能填写模板字段
 * 自动判断关键字附近的填写模式（下划线/冒号后/标签后/占位符），在正确位置插入内容并保持格式
 */
export declare const smartFillFieldDefinition: ToolDefinition;
export declare const smartFillFieldHandler: ToolHandler;
/**
 * 替换书签内容
 * 通过书签名定位并替换内容，保持书签位置格式
 */
export declare const replaceBookmarkContentDefinition: ToolDefinition;
export declare const replaceBookmarkContentHandler: ToolHandler;
export declare const contentTools: RegisteredTool[];
export default contentTools;
//# sourceMappingURL=content.d.ts.map
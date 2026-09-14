/**
 * Input: Word 工具定义
 * Output: Word 工具注册数组
 * Pos: Word Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word Tools入口 - Word工具汇总模块
 * 整合格式化、内容操作、文档管理和校对的所有Tools
 */
import { RegisteredTool } from '../../types/tools';
/**
 * 所有Word相关的Tools
 * 包含：
 * - 格式化Tools: apply_style, set_font, generate_toc, insert_bookmark, set_page_setup
 * - 内容Tools: insert_text, find_replace, insert_table, set_paragraph, get_active_document, insert_image, set_font_style
 * - 文档管理Tools: get_open_documents, switch_document, open_document, get_document_text, insert_header, insert_footer, generate_doc_toc
 * - 校对Tools: enable_track_changes, get_track_changes_status, replace_range, proofread_basic
 * - 深水区Tools（P3）: get_bookmarks, get_comments, get_document_stats, insert_hyperlink,
 *   get_tables, get_table_data, set_table_cell, add_table_lines, delete_table_line,
 *   merge_table_cells, split_table_cell, set_table_format, convert_table_to_text
 * - 文档生产Tools（P3-3）: insert_page_numbers, set_columns, get_revisions, accept_revisions,
 *   reject_revisions, delete_comment
 * - 长尾Tools（P3-4）: get_content_controls, add_content_control, add_footnote, add_endnote,
 *   get_notes, insert_index, insert_cross_reference, mail_merge
 */
export declare const wordTools: RegisteredTool[];
export { formatTools } from './format';
export { contentTools } from './content';
export { documentTools } from './document';
export { proofreadTools } from './proofread';
export { wordDeepTools } from './deep';
export { wordProduceTools } from './produce';
export { wordLongTailTools } from './longtail';
export { applyStyleDefinition, applyStyleHandler, setFontDefinition, setFontHandler, generateTocDefinition, generateTocHandler, insertBookmarkDefinition, insertBookmarkHandler, setPageSetupDefinition, setPageSetupHandler, } from './format';
export { insertTextDefinition, insertTextHandler, findReplaceDefinition, findReplaceHandler, insertTableDefinition, insertTableHandler, setParagraphDefinition, setParagraphHandler, getActiveDocumentDefinition, getActiveDocumentHandler, insertPageBreakDefinition, insertPageBreakHandler, insertCommentDefinition, insertCommentHandler, setTextColorDefinition, setTextColorHandler, getParagraphsDefinition, getParagraphsHandler, findInDocumentDefinition, findInDocumentHandler, smartFillFieldDefinition, smartFillFieldHandler, replaceBookmarkContentDefinition, replaceBookmarkContentHandler, } from './content';
export { getOpenDocumentsDefinition, getOpenDocumentsHandler, switchDocumentDefinition, switchDocumentHandler, openDocumentDefinition, openDocumentHandler, createDocumentDefinition, createDocumentHandler, closeDocumentDefinition, closeDocumentHandler, getDocumentTextDefinition, getDocumentTextHandler, insertHeaderDefinition, insertHeaderHandler, insertFooterDefinition, insertFooterHandler, insertSectionBreakDefinition, insertSectionBreakHandler, setLineSpacingDefinition, setLineSpacingHandler, } from './document';
export { enableTrackChangesDefinition, enableTrackChangesHandler, getTrackChangesStatusDefinition, getTrackChangesStatusHandler, replaceRangeDefinition, replaceRangeHandler, proofreadBasicDefinition, proofreadBasicHandler, } from './proofread';
export default wordTools;
//# sourceMappingURL=index.d.ts.map
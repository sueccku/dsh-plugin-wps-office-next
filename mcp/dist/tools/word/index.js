"use strict";
/**
 * Input: Word 工具定义
 * Output: Word 工具注册数组
 * Pos: Word Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word Tools入口 - Word工具汇总模块
 * 整合格式化、内容操作、文档管理和校对的所有Tools
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.closeDocumentDefinition = exports.createDocumentHandler = exports.createDocumentDefinition = exports.openDocumentHandler = exports.openDocumentDefinition = exports.switchDocumentHandler = exports.switchDocumentDefinition = exports.getOpenDocumentsHandler = exports.getOpenDocumentsDefinition = exports.replaceBookmarkContentHandler = exports.replaceBookmarkContentDefinition = exports.smartFillFieldHandler = exports.smartFillFieldDefinition = exports.findInDocumentHandler = exports.findInDocumentDefinition = exports.getParagraphsHandler = exports.getParagraphsDefinition = exports.setTextColorHandler = exports.setTextColorDefinition = exports.insertCommentHandler = exports.insertCommentDefinition = exports.insertPageBreakHandler = exports.insertPageBreakDefinition = exports.getActiveDocumentHandler = exports.getActiveDocumentDefinition = exports.setParagraphHandler = exports.setParagraphDefinition = exports.insertTableHandler = exports.insertTableDefinition = exports.findReplaceHandler = exports.findReplaceDefinition = exports.insertTextHandler = exports.insertTextDefinition = exports.setPageSetupHandler = exports.setPageSetupDefinition = exports.insertBookmarkHandler = exports.insertBookmarkDefinition = exports.generateTocHandler = exports.generateTocDefinition = exports.setFontHandler = exports.setFontDefinition = exports.applyStyleHandler = exports.applyStyleDefinition = exports.wordProduceTools = exports.wordDeepTools = exports.proofreadTools = exports.documentTools = exports.contentTools = exports.formatTools = exports.wordTools = void 0;
exports.proofreadBasicHandler = exports.proofreadBasicDefinition = exports.replaceRangeHandler = exports.replaceRangeDefinition = exports.getTrackChangesStatusHandler = exports.getTrackChangesStatusDefinition = exports.enableTrackChangesHandler = exports.enableTrackChangesDefinition = exports.setLineSpacingHandler = exports.setLineSpacingDefinition = exports.insertSectionBreakHandler = exports.insertSectionBreakDefinition = exports.insertFooterHandler = exports.insertFooterDefinition = exports.insertHeaderHandler = exports.insertHeaderDefinition = exports.getDocumentTextHandler = exports.getDocumentTextDefinition = exports.closeDocumentHandler = void 0;
const format_1 = require("./format");
const content_1 = require("./content");
const document_1 = require("./document");
const proofread_1 = require("./proofread");
const deep_1 = require("./deep");
const produce_1 = require("./produce");
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
 */
exports.wordTools = [
    ...format_1.formatTools,
    ...content_1.contentTools,
    ...document_1.documentTools,
    ...proofread_1.proofreadTools,
    ...deep_1.wordDeepTools,
    ...produce_1.wordProduceTools,
];
// 分别导出，方便按需使用
var format_2 = require("./format");
Object.defineProperty(exports, "formatTools", { enumerable: true, get: function () { return format_2.formatTools; } });
var content_2 = require("./content");
Object.defineProperty(exports, "contentTools", { enumerable: true, get: function () { return content_2.contentTools; } });
var document_2 = require("./document");
Object.defineProperty(exports, "documentTools", { enumerable: true, get: function () { return document_2.documentTools; } });
var proofread_2 = require("./proofread");
Object.defineProperty(exports, "proofreadTools", { enumerable: true, get: function () { return proofread_2.proofreadTools; } });
var deep_2 = require("./deep");
Object.defineProperty(exports, "wordDeepTools", { enumerable: true, get: function () { return deep_2.wordDeepTools; } });
var produce_2 = require("./produce");
Object.defineProperty(exports, "wordProduceTools", { enumerable: true, get: function () { return produce_2.wordProduceTools; } });
// 导出单独的定义和处理器，方便测试
var format_3 = require("./format");
Object.defineProperty(exports, "applyStyleDefinition", { enumerable: true, get: function () { return format_3.applyStyleDefinition; } });
Object.defineProperty(exports, "applyStyleHandler", { enumerable: true, get: function () { return format_3.applyStyleHandler; } });
Object.defineProperty(exports, "setFontDefinition", { enumerable: true, get: function () { return format_3.setFontDefinition; } });
Object.defineProperty(exports, "setFontHandler", { enumerable: true, get: function () { return format_3.setFontHandler; } });
Object.defineProperty(exports, "generateTocDefinition", { enumerable: true, get: function () { return format_3.generateTocDefinition; } });
Object.defineProperty(exports, "generateTocHandler", { enumerable: true, get: function () { return format_3.generateTocHandler; } });
Object.defineProperty(exports, "insertBookmarkDefinition", { enumerable: true, get: function () { return format_3.insertBookmarkDefinition; } });
Object.defineProperty(exports, "insertBookmarkHandler", { enumerable: true, get: function () { return format_3.insertBookmarkHandler; } });
Object.defineProperty(exports, "setPageSetupDefinition", { enumerable: true, get: function () { return format_3.setPageSetupDefinition; } });
Object.defineProperty(exports, "setPageSetupHandler", { enumerable: true, get: function () { return format_3.setPageSetupHandler; } });
var content_3 = require("./content");
Object.defineProperty(exports, "insertTextDefinition", { enumerable: true, get: function () { return content_3.insertTextDefinition; } });
Object.defineProperty(exports, "insertTextHandler", { enumerable: true, get: function () { return content_3.insertTextHandler; } });
Object.defineProperty(exports, "findReplaceDefinition", { enumerable: true, get: function () { return content_3.findReplaceDefinition; } });
Object.defineProperty(exports, "findReplaceHandler", { enumerable: true, get: function () { return content_3.findReplaceHandler; } });
Object.defineProperty(exports, "insertTableDefinition", { enumerable: true, get: function () { return content_3.insertTableDefinition; } });
Object.defineProperty(exports, "insertTableHandler", { enumerable: true, get: function () { return content_3.insertTableHandler; } });
Object.defineProperty(exports, "setParagraphDefinition", { enumerable: true, get: function () { return content_3.setParagraphDefinition; } });
Object.defineProperty(exports, "setParagraphHandler", { enumerable: true, get: function () { return content_3.setParagraphHandler; } });
Object.defineProperty(exports, "getActiveDocumentDefinition", { enumerable: true, get: function () { return content_3.getActiveDocumentDefinition; } });
Object.defineProperty(exports, "getActiveDocumentHandler", { enumerable: true, get: function () { return content_3.getActiveDocumentHandler; } });
Object.defineProperty(exports, "insertPageBreakDefinition", { enumerable: true, get: function () { return content_3.insertPageBreakDefinition; } });
Object.defineProperty(exports, "insertPageBreakHandler", { enumerable: true, get: function () { return content_3.insertPageBreakHandler; } });
Object.defineProperty(exports, "insertCommentDefinition", { enumerable: true, get: function () { return content_3.insertCommentDefinition; } });
Object.defineProperty(exports, "insertCommentHandler", { enumerable: true, get: function () { return content_3.insertCommentHandler; } });
Object.defineProperty(exports, "setTextColorDefinition", { enumerable: true, get: function () { return content_3.setTextColorDefinition; } });
Object.defineProperty(exports, "setTextColorHandler", { enumerable: true, get: function () { return content_3.setTextColorHandler; } });
Object.defineProperty(exports, "getParagraphsDefinition", { enumerable: true, get: function () { return content_3.getParagraphsDefinition; } });
Object.defineProperty(exports, "getParagraphsHandler", { enumerable: true, get: function () { return content_3.getParagraphsHandler; } });
Object.defineProperty(exports, "findInDocumentDefinition", { enumerable: true, get: function () { return content_3.findInDocumentDefinition; } });
Object.defineProperty(exports, "findInDocumentHandler", { enumerable: true, get: function () { return content_3.findInDocumentHandler; } });
Object.defineProperty(exports, "smartFillFieldDefinition", { enumerable: true, get: function () { return content_3.smartFillFieldDefinition; } });
Object.defineProperty(exports, "smartFillFieldHandler", { enumerable: true, get: function () { return content_3.smartFillFieldHandler; } });
Object.defineProperty(exports, "replaceBookmarkContentDefinition", { enumerable: true, get: function () { return content_3.replaceBookmarkContentDefinition; } });
Object.defineProperty(exports, "replaceBookmarkContentHandler", { enumerable: true, get: function () { return content_3.replaceBookmarkContentHandler; } });
var document_3 = require("./document");
Object.defineProperty(exports, "getOpenDocumentsDefinition", { enumerable: true, get: function () { return document_3.getOpenDocumentsDefinition; } });
Object.defineProperty(exports, "getOpenDocumentsHandler", { enumerable: true, get: function () { return document_3.getOpenDocumentsHandler; } });
Object.defineProperty(exports, "switchDocumentDefinition", { enumerable: true, get: function () { return document_3.switchDocumentDefinition; } });
Object.defineProperty(exports, "switchDocumentHandler", { enumerable: true, get: function () { return document_3.switchDocumentHandler; } });
Object.defineProperty(exports, "openDocumentDefinition", { enumerable: true, get: function () { return document_3.openDocumentDefinition; } });
Object.defineProperty(exports, "openDocumentHandler", { enumerable: true, get: function () { return document_3.openDocumentHandler; } });
Object.defineProperty(exports, "createDocumentDefinition", { enumerable: true, get: function () { return document_3.createDocumentDefinition; } });
Object.defineProperty(exports, "createDocumentHandler", { enumerable: true, get: function () { return document_3.createDocumentHandler; } });
Object.defineProperty(exports, "closeDocumentDefinition", { enumerable: true, get: function () { return document_3.closeDocumentDefinition; } });
Object.defineProperty(exports, "closeDocumentHandler", { enumerable: true, get: function () { return document_3.closeDocumentHandler; } });
Object.defineProperty(exports, "getDocumentTextDefinition", { enumerable: true, get: function () { return document_3.getDocumentTextDefinition; } });
Object.defineProperty(exports, "getDocumentTextHandler", { enumerable: true, get: function () { return document_3.getDocumentTextHandler; } });
Object.defineProperty(exports, "insertHeaderDefinition", { enumerable: true, get: function () { return document_3.insertHeaderDefinition; } });
Object.defineProperty(exports, "insertHeaderHandler", { enumerable: true, get: function () { return document_3.insertHeaderHandler; } });
Object.defineProperty(exports, "insertFooterDefinition", { enumerable: true, get: function () { return document_3.insertFooterDefinition; } });
Object.defineProperty(exports, "insertFooterHandler", { enumerable: true, get: function () { return document_3.insertFooterHandler; } });
Object.defineProperty(exports, "insertSectionBreakDefinition", { enumerable: true, get: function () { return document_3.insertSectionBreakDefinition; } });
Object.defineProperty(exports, "insertSectionBreakHandler", { enumerable: true, get: function () { return document_3.insertSectionBreakHandler; } });
Object.defineProperty(exports, "setLineSpacingDefinition", { enumerable: true, get: function () { return document_3.setLineSpacingDefinition; } });
Object.defineProperty(exports, "setLineSpacingHandler", { enumerable: true, get: function () { return document_3.setLineSpacingHandler; } });
var proofread_3 = require("./proofread");
Object.defineProperty(exports, "enableTrackChangesDefinition", { enumerable: true, get: function () { return proofread_3.enableTrackChangesDefinition; } });
Object.defineProperty(exports, "enableTrackChangesHandler", { enumerable: true, get: function () { return proofread_3.enableTrackChangesHandler; } });
Object.defineProperty(exports, "getTrackChangesStatusDefinition", { enumerable: true, get: function () { return proofread_3.getTrackChangesStatusDefinition; } });
Object.defineProperty(exports, "getTrackChangesStatusHandler", { enumerable: true, get: function () { return proofread_3.getTrackChangesStatusHandler; } });
Object.defineProperty(exports, "replaceRangeDefinition", { enumerable: true, get: function () { return proofread_3.replaceRangeDefinition; } });
Object.defineProperty(exports, "replaceRangeHandler", { enumerable: true, get: function () { return proofread_3.replaceRangeHandler; } });
Object.defineProperty(exports, "proofreadBasicDefinition", { enumerable: true, get: function () { return proofread_3.proofreadBasicDefinition; } });
Object.defineProperty(exports, "proofreadBasicHandler", { enumerable: true, get: function () { return proofread_3.proofreadBasicHandler; } });
exports.default = exports.wordTools;
//# sourceMappingURL=index.js.map
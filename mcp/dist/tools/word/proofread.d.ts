/**
 * Input: 文档校对工具参数
 * Output: 校对结果和修订跟踪
 * Pos: Word 校对工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word校对Tools - 文档校对与修订模块
 * 处理文档校对、修订模式控制、精确范围替换等操作
 *
 * 包含：
 * - wps_word_enable_track_changes: 开启/关闭修订模式
 * - wps_word_get_track_changes_status: 获取修订模式状态
 * - wps_word_replace_range: 按字符范围替换文本（修订模式下跟踪）
 * - wps_word_proofread_basic: 基础文本校对（正则检测错别字/语病）
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
export declare const enableTrackChangesDefinition: ToolDefinition;
export declare const enableTrackChangesHandler: ToolHandler;
export declare const getTrackChangesStatusDefinition: ToolDefinition;
export declare const getTrackChangesStatusHandler: ToolHandler;
export declare const replaceRangeDefinition: ToolDefinition;
export declare const replaceRangeHandler: ToolHandler;
export declare const proofreadBasicDefinition: ToolDefinition;
export declare const proofreadBasicHandler: ToolHandler;
export declare const proofreadTools: RegisteredTool[];
export default proofreadTools;
//# sourceMappingURL=proofread.d.ts.map
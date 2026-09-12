/**
 * Input: 文本框与标题工具参数
 * Output: 文本框操作结果
 * Pos: PPT 文本框与标题工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 删除文本框
 */
export declare const deleteTextboxDefinition: ToolDefinition;
export declare const deleteTextboxHandler: ToolHandler;
/**
 * 获取文本框列表
 */
export declare const getTextboxesDefinition: ToolDefinition;
export declare const getTextboxesHandler: ToolHandler;
/**
 * 设置文本框文本
 */
export declare const setTextboxTextDefinition: ToolDefinition;
export declare const setTextboxTextHandler: ToolHandler;
/**
 * 设置文本框样式
 */
export declare const setTextboxStyleDefinition: ToolDefinition;
export declare const setTextboxStyleHandler: ToolHandler;
/**
 * 获取幻灯片标题
 */
export declare const getSlideTitleDefinition: ToolDefinition;
export declare const getSlideTitleHandler: ToolHandler;
/**
 * 设置幻灯片副标题
 */
export declare const setSlideSubtitleDefinition: ToolDefinition;
export declare const setSlideSubtitleHandler: ToolHandler;
/**
 * 设置幻灯片正文内容
 */
export declare const setSlideContentDefinition: ToolDefinition;
export declare const setSlideContentHandler: ToolHandler;
/**
 * 导出所有文本框与标题相关的Tools
 */
export declare const textboxTools: RegisteredTool[];
export default textboxTools;
//# sourceMappingURL=textbox.d.ts.map
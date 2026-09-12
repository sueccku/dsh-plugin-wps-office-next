/**
 * Input: Word 格式化参数
 * Output: 样式与字体设置结果
 * Pos: Word 格式化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word格式化Tools - 排版格式化模块
 * 处理文档样式、字体、目录等格式化需求
 *
 * 包含：
 * - wps_word_apply_style: 应用样式到选中区域
 * - wps_word_set_font: 设置字体格式
 * - wps_word_generate_toc: 生成目录
 * - wps_word_insert_bookmark: 插入书签
 * - wps_word_set_page_setup: 设置页面布局
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
/**
 * 应用样式到选中区域
 * 快速应用Word内置样式，比如标题1、标题2、正文等
 */
export declare const applyStyleDefinition: ToolDefinition;
export declare const applyStyleHandler: ToolHandler;
/**
 * 设置字体格式
 */
export declare const setFontDefinition: ToolDefinition;
export declare const setFontHandler: ToolHandler;
/**
 * 生成目录
 * 自动根据文档中的标题样式生成目录
 */
export declare const generateTocDefinition: ToolDefinition;
export declare const generateTocHandler: ToolHandler;
/**
 * 插入书签
 * 在当前光标位置插入书签，方便交叉引用和导航
 */
export declare const insertBookmarkDefinition: ToolDefinition;
export declare const insertBookmarkHandler: ToolHandler;
/**
 * 设置页面布局
 * 调整页面方向、边距等页面设置
 */
export declare const setPageSetupDefinition: ToolDefinition;
export declare const setPageSetupHandler: ToolHandler;
/**
 * 导出所有格式化相关的Tools
 */
export declare const formatTools: RegisteredTool[];
export default formatTools;
//# sourceMappingURL=format.d.ts.map
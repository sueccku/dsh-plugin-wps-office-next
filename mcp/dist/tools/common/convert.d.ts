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
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
import { WpsAppType } from '../../types/wps';
/**
 * 根据文件扩展名判断应该用哪个WPS应用类型
 */
declare const getAppTypeByExtension: (filePath: string) => WpsAppType | null;
/**
 * 根据输出格式获取对应的文件格式代码
 * WPS的格式代码和微软Office基本兼容，但也有自己的一套
 */
declare const getFormatCode: (format: string, appType: WpsAppType) => number;
/**
 * 转换为PDF格式
 * 支持Word、Excel、PPT文档一键转PDF
 */
export declare const convertToPdfDefinition: ToolDefinition;
export declare const convertToPdfHandler: ToolHandler;
/**
 * 格式互转
 * 支持docx<->doc, xlsx<->xls, pptx<->ppt等多种格式转换
 */
export declare const convertFormatDefinition: ToolDefinition;
export declare const convertFormatHandler: ToolHandler;
/**
 * 导出所有转换相关的Tools
 */
export declare const convertTools: RegisteredTool[];
export { getAppTypeByExtension, getFormatCode };
export default convertTools;
//# sourceMappingURL=convert.d.ts.map
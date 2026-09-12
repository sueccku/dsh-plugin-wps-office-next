/**
 * Input: 通用操作工具参数
 * Output: 操作结果
 * Pos: 通用操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 通用操作Tools - 保存/连接检测/文本选取等基础模块
 *
 * 包含：
 * - wps_common_save: 保存当前文档
 * - wps_common_save_as: 另存为
 * - wps_common_ping: 检测WPS连接
 * - wps_common_wire_check: 检查通信线路
 * - wps_common_get_app_info: 获取WPS应用信息
 * - wps_common_get_selected_text: 获取选中文本
 * - wps_common_set_selected_text: 替换选中文本
 */
import { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';
export declare const saveDefinition: ToolDefinition;
export declare const saveHandler: ToolHandler;
export declare const saveAsDefinition: ToolDefinition;
export declare const saveAsHandler: ToolHandler;
export declare const pingDefinition: ToolDefinition;
export declare const pingHandler: ToolHandler;
export declare const wireCheckDefinition: ToolDefinition;
export declare const wireCheckHandler: ToolHandler;
export declare const getAppInfoDefinition: ToolDefinition;
export declare const getAppInfoHandler: ToolHandler;
export declare const getSelectedTextDefinition: ToolDefinition;
export declare const getSelectedTextHandler: ToolHandler;
export declare const setSelectedTextDefinition: ToolDefinition;
export declare const setSelectedTextHandler: ToolHandler;
export declare const generalTools: RegisteredTool[];
export default generalTools;
//# sourceMappingURL=general.d.ts.map
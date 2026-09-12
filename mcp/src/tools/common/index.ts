/**
 * Input: 通用工具定义
 * Output: 通用工具注册数组
 * Pos: Common Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Common Tools入口 - 通用工具汇总模块
 * 整合跨应用的通用Tools
 *
 * 导出所有通用的Tools，Word/Excel/PPT均可使用
 * 主要包括文档转换、格式互转等跨应用功能
 */

import { RegisteredTool } from '../../types/tools';
import { convertTools } from './convert';
import { generalTools } from './general';

/**
 * 所有通用的Tools
 * 包含：
 * - 转换Tools: convert_to_pdf, convert_format
 * - 通用Tools: save, save_as, ping, wire_check, get_app_info, get_selected_text, set_selected_text
 *
 * 后续可继续添加更多通用功能，如：
 * - 文档比较
 * - 批量处理
 * - 云端同步
 */
export const commonTools: RegisteredTool[] = [
  ...convertTools,
  ...generalTools,
];

// 分别导出，方便按需使用
export { convertTools } from './convert';
export { generalTools } from './general';

// 导出单独的定义和处理器，方便测试
export {
  convertToPdfDefinition,
  convertToPdfHandler,
  convertFormatDefinition,
  convertFormatHandler,
  getAppTypeByExtension,
  getFormatCode,
} from './convert';

// 导出通用操作工具的定义和处理器
export {
  saveDefinition,
  saveHandler,
  saveAsDefinition,
  saveAsHandler,
  pingDefinition,
  pingHandler,
  wireCheckDefinition,
  wireCheckHandler,
  getAppInfoDefinition,
  getAppInfoHandler,
  getSelectedTextDefinition,
  getSelectedTextHandler,
  setSelectedTextDefinition,
  setSelectedTextHandler,
} from './general';

export default commonTools;

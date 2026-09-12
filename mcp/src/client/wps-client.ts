/**
 * Input: 平台信息与WPS调用参数
 * Output: WPS API 调用结果
 * Pos: Windows COM 专用客户端。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS 通信客户端 - 仅 Windows，通过 PowerShell 调用 WPS COM 接口
 */

import {
  WpsEndpointConfig,
  WpsApiRequest,
  WpsApiResponse,
  WpsAppType,
  WpsClientStatus,
  DocumentInfo,
  WorkbookInfo,
  PresentationInfo,
} from '../types/wps';
import { log, logRequest, logResponse } from '../utils/logger';
import { errorUtils } from '../utils/error';
import { comHost } from './com-host';
import { appendFileSync } from 'node:fs';

// ==================== PPT 目标文稿锁定（避免多文稿打开时 ActivePresentation 漂移）====================
// 通过 wps_ppt_set_active_target 设置后，所有 PRESENTATION 类调用（executeMethod 带 WpsAppType.PRESENTATION）
// 自动注入 presentationName，由 wps-com.ps1 的 Get-TargetPres 精确定位目标文稿；单次调用显式传 presentationName 可覆盖。
let pptTargetName: string | undefined;
export function setPptTarget(name?: string): void {
  pptTargetName = name && name.trim() ? name.trim() : undefined;
}
export function getPptTarget(): string | undefined {
  return pptTargetName;
}

/**
 * 执行PowerShell命令 (Windows)
 */
/**
 * 统一执行接口 - 通过常驻 COM host 调用 WPS
 * 进程生命周期、串行化、超时与崩溃恢复由 com-host.ts 负责
 */
/**
 * 参数契约追踪：把每次调用实际发出的 action 与参数键写入 JSONL
 * 由 WPS_OFFICE_TRACE 环境变量开启，用于发现「工具发了参数但底层不读」这类静默失败
 */
function traceCall(action: string, params: Record<string, unknown>, success: boolean): void {
  const target = process.env.WPS_OFFICE_TRACE;
  if (!target) return;
  try {
    appendFileSync(target, JSON.stringify({ ts: Date.now(), action, keys: Object.keys(params || {}), success }) + '\n');
  } catch {
    // 追踪本身绝不能影响调用
  }
}

async function execWpsAction(action: string, params: Record<string, unknown> = {}): Promise<unknown> {
  return comHost.invoke(action, params);
}

/**
 * WPS客户端类 - Windows COM 通信
 */
export class WpsClient {
  private status: WpsClientStatus;

  constructor(_config?: Partial<WpsEndpointConfig>) {
    this.status = { connected: false };
    log.info('WPS Client initialized', { method: 'PowerShell COM' });
  }

  /**
   * 调用WPS接口（跨平台）
   */
  async invokeAction<T = unknown>(action: string, params: Record<string, unknown> = {}): Promise<WpsApiResponse<T>> {
    const startTime = Date.now();
    logRequest(action, params);

    try {
      const result = await execWpsAction(action, params) as WpsApiResponse<T>;
      const duration = Date.now() - startTime;
      logResponse(action, result.success, duration);
      traceCall(action, params, result.success === true);

      if (result.success) {
        this.status.connected = true;
        this.status.lastHeartbeat = new Date();
      }

      return result;
    } catch (error) {
      const duration = Date.now() - startTime;
      logResponse(action, false, duration);
      traceCall(action, params, false);
      this.status.connected = false;
      throw errorUtils.wrap(error, `WPS 调用失败（PowerShell COM）: ${action}`);
    }
  }

  /**
   * 兼容旧API
   */
  async callApi<T = unknown>(request: WpsApiRequest): Promise<WpsApiResponse<T>> {
    const actionMap: Record<string, string> = {
      // Excel 旧API
      'workbook.getActive': 'getActiveWorkbook',
      'cell.getValue': 'getCellValue',
      'cell.setValue': 'setCellValue',
      'range.getData': 'getRangeData',
      'range.setData': 'setRangeData',
      // Word 旧API
      'document.getActive': 'getActiveDocument',
      'document.getText': 'getDocumentText',
      'document.insertText': 'insertText',
      // PPT 旧API
      'presentation.getActive': 'getActivePresentation',
      'presentation.addSlide': 'addSlide',
      // 通用旧API
      'file.save': 'save',
      'file.saveAs': 'saveAs',
      'file.open': 'openFile',
      'ping': 'ping',
    };
    const action = actionMap[request.method] || request.method;
    return this.invokeAction<T>(action, request.params || {});
  }

  /**
   * 检查WPS连接状态
   */
  async checkConnection(): Promise<boolean> {
    try {
      const result = await this.invokeAction('ping');
      this.status.connected = result.success;
      return result.success;
    } catch {
      this.status.connected = false;
      this.status.error = 'Connection check failed';
      return false;
    }
  }

  /**
   * 获取客户端状态
   */
  getStatus(): WpsClientStatus {
    return { ...this.status };
  }

  // ==================== 表格操作 (WPS表格) ====================

  async getActiveWorkbook(): Promise<WorkbookInfo | null> {
    const response = await this.invokeAction<WorkbookInfo>('getActiveWorkbook');
    return response.success ? response.data || null : null;
  }

  async getCellValue(sheet: string | number, row: number, col: number): Promise<unknown> {
    const response = await this.invokeAction<{ value: unknown }>('getCellValue', { sheet, row, col });
    return response.data?.value;
  }

  async setCellValue(sheet: string | number, row: number, col: number, value: unknown): Promise<boolean> {
    const response = await this.invokeAction('setCellValue', { sheet, row, col, value });
    return response.success;
  }

  async getRangeData(sheet: string | number | undefined, range: string): Promise<unknown[][]> {
    const response = await this.invokeAction<{ data: unknown[][] }>('getRangeData', { sheet, range });
    return response.data?.data || [];
  }

  async setRangeData(sheet: string | number | undefined, range: string, data: unknown[][]): Promise<boolean> {
    const response = await this.invokeAction('setRangeData', { sheet, range, data });
    return response.success;
  }

  async setFormula(sheet: string | number, row: number, col: number, formula: string): Promise<boolean> {
    const response = await this.invokeAction('setFormula', { sheet, row, col, formula });
    return response.success;
  }

  // ==================== 文档操作 (WPS文字) ====================

  async getActiveDocument(): Promise<DocumentInfo | null> {
    const response = await this.invokeAction<DocumentInfo>('getActiveDocument');
    return response.success ? response.data || null : null;
  }

  async createDocument(): Promise<boolean> {
    const response = await this.invokeAction('createDocument');
    return response.success;
  }

  async insertText(text: string, position?: number): Promise<boolean> {
    const response = await this.invokeAction('insertText', { text, position });
    return response.success;
  }

  async getDocumentText(): Promise<string> {
    const response = await this.invokeAction<{ text: string }>('getDocumentText');
    return response.data?.text || '';
  }

  // ==================== 演示操作 (WPS演示) ====================

  async getActivePresentation(): Promise<PresentationInfo | null> {
    const response = await this.invokeAction<PresentationInfo>('getActivePresentation');
    return response.success ? response.data || null : null;
  }

  async createPresentation(): Promise<boolean> {
    const response = await this.invokeAction('createPresentation');
    return response.success;
  }

  async addSlide(layout?: string): Promise<boolean> {
    const response = await this.invokeAction('addSlide', { layout });
    return response.success;
  }

  // ==================== 通用操作 ====================

  async executeMethod<T = unknown>(
    method: string,
    params?: Record<string, unknown>,
    appType?: WpsAppType
  ): Promise<WpsApiResponse<T>> {
    let finalParams: Record<string, unknown> = params || {};
    // 锁定目标文稿后，为演示类调用自动注入 presentationName（已显式传入则不覆盖）
    if (appType === WpsAppType.PRESENTATION && pptTargetName && finalParams.presentationName === undefined) {
      finalParams = { ...finalParams, presentationName: pptTargetName };
    }
    return this.invokeAction<T>(method, finalParams);
  }

  async openFile(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('openFile', { path: filePath });
    return response.success;
  }

  async saveFile(_appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('save');
    return response.success;
  }

  async saveFileAs(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('saveAs', { path: filePath });
    return response.success;
  }
}

// 导出单例
export const wpsClient = new WpsClient();

export default WpsClient;
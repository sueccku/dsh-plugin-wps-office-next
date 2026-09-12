"use strict";
/**
 * Input: 平台信息与WPS调用参数
 * Output: WPS API 调用结果
 * Pos: Windows COM 专用客户端。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS 通信客户端 - 仅 Windows，通过 PowerShell 调用 WPS COM 接口
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.wpsClient = exports.WpsClient = void 0;
exports.setPptTarget = setPptTarget;
exports.getPptTarget = getPptTarget;
const wps_1 = require("../types/wps");
const logger_1 = require("../utils/logger");
const error_1 = require("../utils/error");
const com_host_1 = require("./com-host");
const node_fs_1 = require("node:fs");
// ==================== PPT 目标文稿锁定（避免多文稿打开时 ActivePresentation 漂移）====================
// 通过 wps_ppt_set_active_target 设置后，所有 PRESENTATION 类调用（executeMethod 带 WpsAppType.PRESENTATION）
// 自动注入 presentationName，由 wps-com.ps1 的 Get-TargetPres 精确定位目标文稿；单次调用显式传 presentationName 可覆盖。
let pptTargetName;
function setPptTarget(name) {
    pptTargetName = name && name.trim() ? name.trim() : undefined;
}
function getPptTarget() {
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
function traceCall(action, params, success) {
    const target = process.env.WPS_OFFICE_TRACE;
    if (!target)
        return;
    try {
        (0, node_fs_1.appendFileSync)(target, JSON.stringify({ ts: Date.now(), action, keys: Object.keys(params || {}), success }) + '\n');
    }
    catch {
        // 追踪本身绝不能影响调用
    }
}
async function execWpsAction(action, params = {}) {
    return com_host_1.comHost.invoke(action, params);
}
/**
 * WPS客户端类 - Windows COM 通信
 */
class WpsClient {
    status;
    constructor(_config) {
        this.status = { connected: false };
        logger_1.log.info('WPS Client initialized', { method: 'PowerShell COM' });
    }
    /**
     * 调用WPS接口（跨平台）
     */
    async invokeAction(action, params = {}) {
        const startTime = Date.now();
        (0, logger_1.logRequest)(action, params);
        try {
            const result = await execWpsAction(action, params);
            const duration = Date.now() - startTime;
            (0, logger_1.logResponse)(action, result.success, duration);
            traceCall(action, params, result.success === true);
            if (result.success) {
                this.status.connected = true;
                this.status.lastHeartbeat = new Date();
            }
            return result;
        }
        catch (error) {
            const duration = Date.now() - startTime;
            (0, logger_1.logResponse)(action, false, duration);
            traceCall(action, params, false);
            this.status.connected = false;
            throw error_1.errorUtils.wrap(error, `WPS 调用失败（PowerShell COM）: ${action}`);
        }
    }
    /**
     * 兼容旧API
     */
    async callApi(request) {
        const actionMap = {
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
        return this.invokeAction(action, request.params || {});
    }
    /**
     * 检查WPS连接状态
     */
    async checkConnection() {
        try {
            const result = await this.invokeAction('ping');
            this.status.connected = result.success;
            return result.success;
        }
        catch {
            this.status.connected = false;
            this.status.error = 'Connection check failed';
            return false;
        }
    }
    /**
     * 获取客户端状态
     */
    getStatus() {
        return { ...this.status };
    }
    // ==================== 表格操作 (WPS表格) ====================
    async getActiveWorkbook() {
        const response = await this.invokeAction('getActiveWorkbook');
        return response.success ? response.data || null : null;
    }
    async getCellValue(sheet, row, col) {
        const response = await this.invokeAction('getCellValue', { sheet, row, col });
        return response.data?.value;
    }
    async setCellValue(sheet, row, col, value) {
        const response = await this.invokeAction('setCellValue', { sheet, row, col, value });
        return response.success;
    }
    async getRangeData(sheet, range) {
        const response = await this.invokeAction('getRangeData', { sheet, range });
        return response.data?.data || [];
    }
    async setRangeData(sheet, range, data) {
        const response = await this.invokeAction('setRangeData', { sheet, range, data });
        return response.success;
    }
    async setFormula(sheet, row, col, formula) {
        const response = await this.invokeAction('setFormula', { sheet, row, col, formula });
        return response.success;
    }
    // ==================== 文档操作 (WPS文字) ====================
    async getActiveDocument() {
        const response = await this.invokeAction('getActiveDocument');
        return response.success ? response.data || null : null;
    }
    async createDocument() {
        const response = await this.invokeAction('createDocument');
        return response.success;
    }
    async insertText(text, position) {
        const response = await this.invokeAction('insertText', { text, position });
        return response.success;
    }
    async getDocumentText() {
        const response = await this.invokeAction('getDocumentText');
        return response.data?.text || '';
    }
    // ==================== 演示操作 (WPS演示) ====================
    async getActivePresentation() {
        const response = await this.invokeAction('getActivePresentation');
        return response.success ? response.data || null : null;
    }
    async createPresentation() {
        const response = await this.invokeAction('createPresentation');
        return response.success;
    }
    async addSlide(layout) {
        const response = await this.invokeAction('addSlide', { layout });
        return response.success;
    }
    // ==================== 通用操作 ====================
    async executeMethod(method, params, appType) {
        let finalParams = params || {};
        // 锁定目标文稿后，为演示类调用自动注入 presentationName（已显式传入则不覆盖）
        if (appType === wps_1.WpsAppType.PRESENTATION && pptTargetName && finalParams.presentationName === undefined) {
            finalParams = { ...finalParams, presentationName: pptTargetName };
        }
        return this.invokeAction(method, finalParams);
    }
    async openFile(filePath, _appType) {
        const response = await this.invokeAction('openFile', { path: filePath });
        return response.success;
    }
    async saveFile(_appType) {
        const response = await this.invokeAction('save');
        return response.success;
    }
    async saveFileAs(filePath, _appType) {
        const response = await this.invokeAction('saveAs', { path: filePath });
        return response.success;
    }
}
exports.WpsClient = WpsClient;
// 导出单例
exports.wpsClient = new WpsClient();
exports.default = WpsClient;
//# sourceMappingURL=wps-client.js.map
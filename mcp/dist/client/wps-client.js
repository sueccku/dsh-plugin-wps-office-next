"use strict";
/**
 * Input: 平台信息与WPS调用参数
 * Output: WPS API 调用结果
 * Pos: Windows COM 专用客户端。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS 通信客户端 - 仅 Windows，通过 PowerShell 调用 WPS COM 接口
 */
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.wpsClient = exports.WpsClient = void 0;
exports.setPptTarget = setPptTarget;
exports.getPptTarget = getPptTarget;
const child_process_1 = require("child_process");
const path = __importStar(require("path"));
const wps_1 = require("../types/wps");
const logger_1 = require("../utils/logger");
const error_1 = require("../utils/error");
// PowerShell脚本路径 (Windows)
const PS_SCRIPT_PATH = path.join(__dirname, '../../scripts/wps-com.ps1');
// PowerShell 默认超时（毫秒）
const PS_TIMEOUT = 30000;
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
async function execPowerShell(action, params = {}) {
    return new Promise((resolve, reject) => {
        // JSON参数通过spawn args数组传递，Node自动处理Windows引号转义
        const paramsJson = JSON.stringify(params);
        const args = [
            '-ExecutionPolicy', 'Bypass',
            '-File', PS_SCRIPT_PATH,
            '-Action', action,
            '-Params', paramsJson
        ];
        logger_1.log.debug('Executing PowerShell', { action, params });
        const ps = (0, child_process_1.spawn)('powershell', args, {
            windowsHide: true,
            stdio: ['pipe', 'pipe', 'pipe']
        });
        let stdout = '';
        let stderr = '';
        let killed = false;
        // 超时保护：防止PowerShell进程挂起，超时时主动 SIGTERM 并记录 PID
        const timeoutHandle = setTimeout(() => {
            killed = true;
            logger_1.log.warn('PowerShell timeout, killing process', { pid: ps.pid, action, timeoutMs: PS_TIMEOUT });
            ps.kill('SIGTERM');
            reject(new Error(`PowerShell 执行超时（${PS_TIMEOUT}ms）: ${action} (PID: ${ps.pid})`));
        }, PS_TIMEOUT);
        ps.stdout.on('data', (data) => {
            stdout += data.toString();
        });
        ps.stderr.on('data', (data) => {
            stderr += data.toString();
        });
        ps.on('close', (code) => {
            clearTimeout(timeoutHandle);
            if (killed)
                return; // 已超时处理，忽略后续事件
            if (code !== 0) {
                // 非零退出码：分别处理 stderr 有内容 / 为空两种路径，避免空 stderr 误入 JSON.parse
                if (stderr) {
                    logger_1.log.error('PowerShell error', { stderr, code, pid: ps.pid, action });
                    reject(new Error(stderr));
                }
                else {
                    logger_1.log.error('PowerShell exited with non-zero code', { code, stdout: stdout.substring(0, 200), pid: ps.pid, action });
                    reject(new Error(`PowerShell 退出码 ${code}: ${stdout.substring(0, 200) || '(无输出)'}`));
                }
                return;
            }
            try {
                const result = JSON.parse(stdout.trim());
                resolve(result);
            }
            catch (_e) {
                logger_1.log.error('Failed to parse PowerShell output', { stdout: stdout.substring(0, 200), pid: ps.pid, action });
                reject(new Error(`PowerShell 输出解析失败（非有效JSON）: ${stdout.substring(0, 200)}`));
            }
        });
        ps.on('error', (err) => {
            clearTimeout(timeoutHandle);
            if (killed)
                return;
            logger_1.log.error('PowerShell spawn error', { error: err.message, pid: ps.pid, action });
            reject(new Error(`无法启动 PowerShell 进程: ${err.message}`));
        });
    });
}
/**
 * 统一执行接口 - 根据平台选择调用方式
 * Mac: 反向轮询模式（MCP Server是服务端，WPS加载项来取命令）
 * Windows: PowerShell调用COM接口
 */
async function execWpsAction(action, params = {}) {
    return execPowerShell(action, params);
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
            if (result.success) {
                this.status.connected = true;
                this.status.lastHeartbeat = new Date();
            }
            return result;
        }
        catch (error) {
            const duration = Date.now() - startTime;
            (0, logger_1.logResponse)(action, false, duration);
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
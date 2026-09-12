"use strict";
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
exports.comHost = exports.ComHost = exports.HOST_SCRIPT = void 0;
/**
 * Input: WPS action name and parameters
 * Output: WPS action result
 * Pos: Windows resident COM host client. Owns the long-lived PowerShell process,
 *      serializes calls, applies per-action timeouts and recovers from crashes.
 *      一旦我被修改，请更新我的头部注释。
 */
const child_process_1 = require("child_process");
const path = __importStar(require("path"));
const logger_1 = require("../utils/logger");
/** Path to the resident host script; overridable for tests and relocated installs. */
exports.HOST_SCRIPT = process.env.WPS_OFFICE_HOST_SCRIPT || path.join(__dirname, '../../..', 'host', 'wps-com-host.ps1');
/** Actions that legitimately run for minutes (export, pivot, beautify, batch insert). */
const LONG_ACTIONS = new Set([
    'convertToPDF',
    'convertFormat',
    'exportChartAsImage',
    'exportRangeAsImage',
    'exportSlideAsImage',
    'createPivotTable',
    'updatePivotTable',
    'beautify',
    'beautifySlide',
    'insertSlidesFromFile',
    'recalculate',
    'proofreadBasic',
    'saveAs',
]);
const DEFAULT_TIMEOUT_MS = Number(process.env.WPS_OFFICE_TIMEOUT_MS || 60000);
const LONG_TIMEOUT_MS = Number(process.env.WPS_OFFICE_LONG_TIMEOUT_MS || 300000);
const STARTUP_TIMEOUT_MS = Number(process.env.WPS_OFFICE_STARTUP_TIMEOUT_MS || 60000);
function timeoutFor(action) {
    return LONG_ACTIONS.has(action) ? LONG_TIMEOUT_MS : DEFAULT_TIMEOUT_MS;
}
/**
 * Resident COM host.
 *
 * One PowerShell process is kept alive for the lifetime of the MCP server. Requests are
 * serialized because WPS COM automation is single threaded; a hung or crashed host is
 * killed and respawned on the next call.
 */
class ComHost {
    child = null;
    stdoutBuffer = '';
    readyWaiters = [];
    inflight = null;
    sequence = 0;
    chain = Promise.resolve();
    stderrTail = [];
    /** True while a host process is attached. */
    get isRunning() {
        return this.child !== null && !this.child.killed;
    }
    /** Invoke one WPS action through the resident host. */
    invoke(action, params = {}) {
        const run = this.chain.then(() => this.dispatch(action, params));
        this.chain = run.catch(() => undefined);
        return run;
    }
    /** Stop the host; the next invoke starts a fresh one. */
    async stop() {
        const child = this.child;
        this.child = null;
        this.failInflight(new Error('COM host stopped'));
        if (child) {
            try {
                child.kill();
            }
            catch { /* already gone */ }
        }
    }
    async dispatch(action, params) {
        await this.ensureStarted();
        const id = ++this.sequence;
        return new Promise((resolve, reject) => {
            const child = this.child;
            if (!child || !child.stdin || child.killed) {
                reject(new Error('COM host is not available'));
                return;
            }
            const timer = setTimeout(() => {
                // A timed-out action means the host is stuck inside COM; the process is not reusable.
                logger_1.log.warn('COM host timeout, restarting host', { action, id });
                this.inflight = null;
                this.killChild(new Error('COM host timeout calling ' + action));
                reject(new Error('WPS 操作超时（' + timeoutFor(action) + 'ms）: ' + action));
            }, timeoutFor(action));
            this.inflight = { id, resolve, reject, timer };
            const frame = JSON.stringify({ id, action, params });
            child.stdin.write(frame + '\n', (error) => {
                if (error) {
                    clearTimeout(timer);
                    this.inflight = null;
                    reject(new Error('COM host write failed: ' + error.message));
                }
            });
        });
    }
    ensureStarted() {
        if (this.isRunning)
            return Promise.resolve();
        return new Promise((resolve, reject) => {
            const timer = setTimeout(() => {
                this.readyWaiters = this.readyWaiters.filter((w) => w.timer !== timer);
                this.killChild(new Error('COM host startup timed out'));
                reject(new Error('COM host 启动超时（' + STARTUP_TIMEOUT_MS + 'ms）: ' + exports.HOST_SCRIPT));
            }, STARTUP_TIMEOUT_MS);
            this.readyWaiters.push({ resolve, reject, timer });
            if (this.child)
                return; // a start is already in flight
            logger_1.log.info('Starting resident COM host', { script: exports.HOST_SCRIPT });
            this.stderrTail = [];
            this.stdoutBuffer = '';
            let child;
            try {
                child = (0, child_process_1.spawn)('powershell', ['-NoProfile', '-NoLogo', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', exports.HOST_SCRIPT], { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] });
            }
            catch (error) {
                this.failReady(new Error('COM host spawn failed: ' + error.message));
                return;
            }
            this.child = child;
            child.stdout?.on('data', (chunk) => this.onStdout(chunk.toString()));
            child.stderr?.on('data', (chunk) => this.onStderr(chunk.toString()));
            child.on('error', (error) => {
                this.killChild(new Error('COM host process error: ' + error.message));
            });
            child.on('exit', (code, signal) => {
                const detail = 'COM host exited (code=' + String(code) + ', signal=' + String(signal) + ')';
                this.killChild(new Error(detail + (this.stderrTail.length ? ' :: ' + this.stderrTail.join(' | ') : '')));
            });
        });
    }
    onStdout(chunk) {
        this.stdoutBuffer += chunk;
        let index = this.stdoutBuffer.indexOf('\n');
        while (index >= 0) {
            const line = this.stdoutBuffer.slice(0, index).trim();
            this.stdoutBuffer = this.stdoutBuffer.slice(index + 1);
            if (line.length > 0)
                this.onFrame(line);
            index = this.stdoutBuffer.indexOf('\n');
        }
    }
    onFrame(line) {
        let frame;
        try {
            frame = JSON.parse(line);
        }
        catch {
            // The protocol owns stdout; anything else is a defect worth surfacing.
            logger_1.log.warn('COM host wrote a non-JSON stdout line', { line: line.slice(0, 200) });
            return;
        }
        if (frame.ready) {
            logger_1.log.info('Resident COM host ready', { pid: frame.pid });
            this.failReady(null);
            return;
        }
        const inflight = this.inflight;
        if (!inflight || frame.id !== inflight.id) {
            logger_1.log.warn('COM host frame without a matching request', { id: frame.id });
            return;
        }
        clearTimeout(inflight.timer);
        this.inflight = null;
        inflight.resolve(frame.result || { success: false, error: 'COM host returned no result' });
    }
    onStderr(chunk) {
        const text = chunk.trim();
        if (!text)
            return;
        this.stderrTail.push(text);
        if (this.stderrTail.length > 8)
            this.stderrTail.shift();
        logger_1.log.debug('COM host stderr', { text: text.slice(0, 400) });
    }
    failReady(error) {
        const waiters = this.readyWaiters;
        this.readyWaiters = [];
        for (const waiter of waiters) {
            clearTimeout(waiter.timer);
            if (error)
                waiter.reject(error);
            else
                waiter.resolve();
        }
    }
    failInflight(error) {
        const inflight = this.inflight;
        if (!inflight)
            return;
        clearTimeout(inflight.timer);
        this.inflight = null;
        inflight.reject(error);
    }
    killChild(cause) {
        const child = this.child;
        this.child = null;
        this.stdoutBuffer = '';
        if (child && !child.killed) {
            try {
                child.kill();
            }
            catch { /* already gone */ }
        }
        this.failReady(cause);
        this.failInflight(cause);
    }
}
exports.ComHost = ComHost;
/** Shared host instance used by the WPS client. */
exports.comHost = new ComHost();
exports.default = exports.comHost;
//# sourceMappingURL=com-host.js.map
/**
 * Input: WPS action name and parameters
 * Output: WPS action result
 * Pos: Windows resident COM host client. Owns the long-lived PowerShell process,
 *      serializes calls, applies per-action timeouts and recovers from crashes.
 *      A timeout reports "state unknown" and shortens the next call instead of retrying blindly,
 *      because the host cannot tell a blocked dialog from a dead WPS.
 *      一旦我被修改，请更新我的头部注释。
 */
import { spawn, ChildProcess } from 'child_process';
import * as path from 'path';
import { log } from '../utils/logger';

/** Path to the resident host script; overridable for tests and relocated installs. */
export const HOST_SCRIPT =
  process.env.WPS_OFFICE_HOST_SCRIPT || path.join(__dirname, '../../..', 'host', 'wps-com-host.ps1');

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
/**
 * After a timeout the host is killed but WPS itself is deliberately left alone: it may still be
 * showing a modal dialog, and force-closing it would throw away whatever the user had unsaved.
 * So the next call runs under a short leash instead - it either proves WPS answers again (and the
 * normal timeout comes back) or fails fast with the same explanation, instead of hanging the
 * session for another full timeout.
 */
const SUSPECT_TIMEOUT_MS = Number(process.env.WPS_OFFICE_SUSPECT_TIMEOUT_MS || 15000);

/**
 * Absolute path to Windows PowerShell 5.1. The action layer depends on 5.1 COM adapter
 * semantics, so the host is pinned to the in-box interpreter instead of whatever 'powershell'
 * resolves to on PATH.
 */
export const POWERSHELL_EXE =
  process.env.WPS_OFFICE_POWERSHELL ||
  path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');

/** Action result as produced by the COM layer: success plus data or error. */
export interface WpsActionOutcome {
  success: boolean;
  data?: unknown;
  error?: string;
  [key: string]: unknown;
}

interface HostFrame {
  id?: number;
  ok?: boolean;
  result?: WpsActionOutcome;
  ms?: number;
  ready?: boolean;
  pid?: number;
  clientPid?: number;
  psVersion?: string;
  error?: string;
}

function timeoutFor(action: string, suspect = false): number {
  const normal = LONG_ACTIONS.has(action) ? LONG_TIMEOUT_MS : DEFAULT_TIMEOUT_MS;
  return suspect ? Math.min(normal, SUSPECT_TIMEOUT_MS) : normal;
}

/**
 * Resident COM host.
 *
 * One PowerShell process is kept alive for the lifetime of the MCP server. Requests are
 * serialized because WPS COM automation is single threaded; a hung or crashed host is
 * killed and respawned on the next call.
 */
export class ComHost {
  private child: ChildProcess | null = null;
  private stdoutBuffer = '';
  private readyWaiters: Array<{ resolve: () => void; reject: (error: Error) => void; timer: NodeJS.Timeout }> = [];
  private inflight: { id: number; resolve: (value: WpsActionOutcome) => void; reject: (error: Error) => void; timer: NodeJS.Timeout } | null = null;
  private sequence = 0;
  private chain: Promise<unknown> = Promise.resolve();
  private stderrTail: string[] = [];
  /**
   * Set by a timeout and cleared by the next successful frame: while it is set, calls use the
   * short suspect timeout so a blocked WPS cannot pin the session for a full minute per call.
   */
  private suspect = false;
  /**
   * The host killed most recently, kept until its process is really gone. The next spawn waits for
   * that exit, because the host holds a single-instance lease over WPS and a successor that arrives
   * while the corpse still owns it would be refused as a second session.
   */
  private awaitingExit: ChildProcess | null = null;

  /** True while a host process is attached. */
  get isRunning(): boolean {
    return this.child !== null && !this.child.killed;
  }

  /** True while WPS is presumed blocked after a timeout (short-timeout mode). */
  get isSuspect(): boolean {
    return this.suspect;
  }

  /** Invoke one WPS action through the resident host. */
  invoke(action: string, params: Record<string, unknown> = {}): Promise<WpsActionOutcome> {
    const run = this.chain.then(() => this.dispatch(action, params));
    this.chain = run.catch(() => undefined);
    return run;
  }

  /** Stop the host; the next invoke starts a fresh one. */
  async stop(): Promise<void> {
    const child = this.child;
    this.child = null;
    this.failInflight(new Error('COM host stopped'));
    if (child && !child.killed) {
      this.awaitingExit = child;
      child.once('exit', () => { if (this.awaitingExit === child) this.awaitingExit = null; });
      try { child.kill(); } catch { this.awaitingExit = null; }
    }
  }

  private async dispatch(action: string, params: Record<string, unknown>): Promise<WpsActionOutcome> {
    await this.ensureStarted();

    const id = ++this.sequence;
    return new Promise<WpsActionOutcome>((resolve, reject) => {
      const child = this.child;
      if (!child || !child.stdin || child.killed) {
        reject(new Error('COM host is not available'));
        return;
      }

      const timeoutMs = timeoutFor(action, this.suspect);
      const timer = setTimeout(() => {
        // A timed-out action means the host is stuck inside COM; the process is not reusable.
        // WPS itself is left running on purpose - see SUSPECT_TIMEOUT_MS.
        this.suspect = true;
        log.warn('COM host timeout; WPS state unknown, host killed', { action, id, timeoutMs });
        this.inflight = null;
        this.killChild(new Error('COM host timeout calling ' + action));
        reject(new Error(
          '操作超时：' + action + ' 超过 ' + timeoutMs + 'ms 未返回。' +
          '状态未知：WPS 可能正被一个对话框阻塞（例如文档密码框），也可能已经无响应；' +
          '插件不会自动关闭 WPS，以免丢掉你未保存的内容。' +
          '下一步：切到 WPS 窗口处理弹出的对话框，然后重试' +
          '（后续调用会先用 ' + SUSPECT_TIMEOUT_MS + 'ms 的短超时快速失败，成功一次即恢复）。'
        ));
      }, timeoutMs);

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

  private async ensureStarted(): Promise<void> {
    if (this.isRunning) return;

    // A host killed a moment ago may still be tearing down; spawning into that window would look
    // like a second DSH session to the lease and fail for a reason the user cannot act on.
    const dying = this.awaitingExit;
    if (dying && dying.exitCode === null && dying.signalCode === null) {
      await new Promise<void>((resolve) => {
        const timer = setTimeout(() => resolve(), 3000);
        dying.once('exit', () => { clearTimeout(timer); resolve(); });
      });
    }
    if (this.awaitingExit === dying) this.awaitingExit = null;
    if (this.isRunning) return;

    return new Promise<void>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.readyWaiters = this.readyWaiters.filter((w) => w.timer !== timer);
        this.killChild(new Error('COM host startup timed out'));
        reject(new Error('COM host 启动超时（' + STARTUP_TIMEOUT_MS + 'ms）: ' + HOST_SCRIPT));
      }, STARTUP_TIMEOUT_MS);

      this.readyWaiters.push({ resolve, reject, timer });

      if (this.child) return; // a start is already in flight

      log.info('Starting resident COM host', { script: HOST_SCRIPT });
      this.stderrTail = [];
      this.stdoutBuffer = '';

      let child: ChildProcess;
      try {
        child = spawn(
          POWERSHELL_EXE,
          ['-NoProfile', '-NoLogo', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', HOST_SCRIPT],
          {
            windowsHide: true,
            stdio: ['pipe', 'pipe', 'pipe'],
            // The host records this pid in its single-instance lease, so a host left behind by a
            // dead MCP server can be recognised as stale and taken over instead of blocking WPS.
            env: { ...process.env, WPS_OFFICE_CLIENT_PID: String(process.pid) },
          }
        );
      } catch (error) {
        this.failReady(new Error('COM host spawn failed: ' + (error as Error).message));
        return;
      }

      this.child = child;

      child.stdout?.on('data', (chunk: Buffer) => this.onStdout(child, chunk.toString()));
      child.stderr?.on('data', (chunk: Buffer) => this.onStderr(child, chunk.toString()));
      child.on('error', (error) => {
        if (this.child !== child) return;
        this.killChild(new Error('COM host process error: ' + error.message));
      });
      child.on('exit', (code, signal) => {
        // A host that was already replaced must not be able to kill its successor: the timeout
        // path kills the old process and the caller's next call spawns a new one immediately, and
        // the old process's exit event then arrives after that new host is already attached.
        if (this.child !== child) return;
        const detail = 'COM host exited (code=' + String(code) + ', signal=' + String(signal) + ')';
        this.killChild(new Error(detail + (this.stderrTail.length ? ' :: ' + this.stderrTail.join(' | ') : '')));
      });
    });
  }

  private onStdout(child: ChildProcess, chunk: string): void {
    // Output from a host that was already replaced must not be mixed into the current buffer.
    if (this.child !== child) return;
    this.stdoutBuffer += chunk;
    let index = this.stdoutBuffer.indexOf('\n');
    while (index >= 0) {
      const line = this.stdoutBuffer.slice(0, index).trim();
      this.stdoutBuffer = this.stdoutBuffer.slice(index + 1);
      if (line.length > 0) this.onFrame(line);
      index = this.stdoutBuffer.indexOf('\n');
    }
  }

  private onFrame(line: string): void {
    let frame: HostFrame;
    try {
      frame = JSON.parse(line) as HostFrame;
    } catch {
      // The protocol owns stdout; anything else is a defect worth surfacing.
      log.warn('COM host wrote a non-JSON stdout line', { line: line.slice(0, 200) });
      return;
    }

    if (frame.ready === true) {
      const major = Number(String(frame.psVersion || '').split('.')[0]);
      if (frame.psVersion && major !== 5) {
        this.killChild(
          new Error('resident COM host reported PowerShell ' + frame.psVersion + '; this host requires Windows PowerShell 5.1 at ' + POWERSHELL_EXE)
        );
        return;
      }
      log.info('Resident COM host ready', { pid: frame.pid, clientPid: frame.clientPid, psVersion: frame.psVersion, exe: POWERSHELL_EXE });
      // Deliberately NOT clearing `suspect` here: a freshly spawned host saying "ready" says
      // nothing about WPS, and clearing it would hand a still-blocked WPS a full timeout again.
      this.failReady(null);
      return;
    }
    if (frame.ready === false) {
      this.killChild(new Error(String(frame.error || 'resident COM host refused to start')));
      return;
    }

    const inflight = this.inflight;
    if (!inflight || frame.id !== inflight.id) {
      log.warn('COM host frame without a matching request', { id: frame.id });
      return;
    }

    clearTimeout(inflight.timer);
    this.inflight = null;
    // WPS answered, so whatever blocked it before is gone; the normal timeout comes back.
    this.suspect = false;
    inflight.resolve(frame.result || { success: false, error: 'COM host returned no result' });
  }

  private onStderr(child: ChildProcess, chunk: string): void {
    if (this.child !== child) return;
    const text = chunk.trim();
    if (!text) return;
    this.stderrTail.push(text);
    if (this.stderrTail.length > 8) this.stderrTail.shift();
    log.debug('COM host stderr', { text: text.slice(0, 400) });
  }

  private failReady(error: Error | null): void {
    const waiters = this.readyWaiters;
    this.readyWaiters = [];
    for (const waiter of waiters) {
      clearTimeout(waiter.timer);
      if (error) waiter.reject(error);
      else waiter.resolve();
    }
  }

  private failInflight(error: Error): void {
    const inflight = this.inflight;
    if (!inflight) return;
    clearTimeout(inflight.timer);
    this.inflight = null;
    inflight.reject(error);
  }

  private killChild(cause: Error): void {
    const child = this.child;
    this.child = null;
    this.stdoutBuffer = '';
    if (child && !child.killed) {
      this.awaitingExit = child;
      child.once('exit', () => { if (this.awaitingExit === child) this.awaitingExit = null; });
      try { child.kill(); } catch { this.awaitingExit = null; }
    }
    this.failReady(cause);
    this.failInflight(cause);
  }
}

/** Shared host instance used by the WPS client. */
export const comHost = new ComHost();

export default comHost;

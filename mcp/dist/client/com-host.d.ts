/** Path to the resident host script; overridable for tests and relocated installs. */
export declare const HOST_SCRIPT: string;
/**
 * Actions that legitimately run for minutes (export, pivot, beautify, batch insert).
 *
 * Every name here must be a real bridge action: a typo silently downgrades the action to the default
 * 60s timeout, so a slow `calculateSheet` would be killed mid-flight and reported as「状态未知」.
 * FIXES 81 fixed three such names (`beautify` -> `beautifySlide`, `recalculate` -> `calculateSheet`,
 * `proofreadBasic` -> not a bridge action at all) and `scripts/verify.mjs` now asserts they all exist.
 */
export declare const LONG_ACTIONS: Set<string>;
/**
 * Absolute path to Windows PowerShell 5.1. The action layer depends on 5.1 COM adapter
 * semantics, so the host is pinned to the in-box interpreter instead of whatever 'powershell'
 * resolves to on PATH.
 */
export declare const POWERSHELL_EXE: string;
/** Action result as produced by the COM layer: success plus data or error. */
export interface WpsActionOutcome {
    success: boolean;
    data?: unknown;
    error?: string;
    [key: string]: unknown;
}
/**
 * 按行切帧的 UTF-8 读取器。
 *
 * 为什么不能写 `buf += chunk.toString()`：一个 chunk 的边界可以落在多字节字符中间（中文一个字 3 字节），
 * `toString()` 会把半个字符解成 U+FFFD，整帧 JSON 里的中文就坏了。`scripts/gen-skill-tools.mjs`、
 * `test/spec-reproduction.test.mjs` 的收帧都吃过这个坑；正式客户端此前没修（改进指南 §9 P1）。
 * StringDecoder 会把不完整的尾部字节留到下一次 `write()` 再拼，所以任何切法都能还原。
 */
export declare class Utf8LineReader {
    private decoder;
    private buffer;
    /** 喂一个 chunk，返回本次凑齐的完整行（已 trim，空行丢弃）。 */
    push(chunk: Buffer | string): string[];
    /** 换一个子进程时必须清空：上一个进程的半帧不能混进新进程的输入。 */
    reset(): void;
}
/**
 * Resident COM host.
 *
 * One PowerShell process is kept alive for the lifetime of the MCP server. Requests are
 * serialized because WPS COM automation is single threaded; a hung or crashed host is
 * killed and respawned on the next call.
 */
export declare class ComHost {
    private child;
    /** 按字节收帧的 stdout 读取器（见 Utf8LineReader：chunk 边界会切开中文）。 */
    private readonly stdoutReader;
    private stderrDecoder;
    private readyWaiters;
    private inflight;
    private sequence;
    private chain;
    private stderrTail;
    /**
     * Set by a timeout and cleared by the next successful frame: while it is set, calls use the
     * short suspect timeout so a blocked WPS cannot pin the session for a full minute per call.
     */
    private suspect;
    /**
     * The host killed most recently, kept until its process is really gone. The next spawn waits for
     * that exit, because the host holds a single-instance lease over WPS and a successor that arrives
     * while the corpse still owns it would be refused as a second session.
     */
    private awaitingExit;
    /** True while a host process is attached. */
    get isRunning(): boolean;
    /** True while WPS is presumed blocked after a timeout (short-timeout mode). */
    get isSuspect(): boolean;
    /** Invoke one WPS action through the resident host. */
    invoke(action: string, params?: Record<string, unknown>): Promise<WpsActionOutcome>;
    /** Stop the host; the next invoke starts a fresh one. */
    stop(): Promise<void>;
    private dispatch;
    private ensureStarted;
    private onStdout;
    private onFrame;
    private onStderr;
    private failReady;
    private failInflight;
    private killChild;
}
/** Shared host instance used by the WPS client. */
export declare const comHost: ComHost;
export default comHost;
//# sourceMappingURL=com-host.d.ts.map
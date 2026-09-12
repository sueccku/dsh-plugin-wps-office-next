/** Path to the resident host script; overridable for tests and relocated installs. */
export declare const HOST_SCRIPT: string;
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
 * Resident COM host.
 *
 * One PowerShell process is kept alive for the lifetime of the MCP server. Requests are
 * serialized because WPS COM automation is single threaded; a hung or crashed host is
 * killed and respawned on the next call.
 */
export declare class ComHost {
    private child;
    private stdoutBuffer;
    private readyWaiters;
    private inflight;
    private sequence;
    private chain;
    private stderrTail;
    /** True while a host process is attached. */
    get isRunning(): boolean;
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
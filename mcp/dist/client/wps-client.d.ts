/**
 * Input: 平台信息与WPS调用参数
 * Output: WPS API 调用结果
 * Pos: Windows COM 专用客户端。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * WPS 通信客户端 - 仅 Windows，通过 PowerShell 调用 WPS COM 接口
 */
import { WpsEndpointConfig, WpsApiRequest, WpsApiResponse, WpsAppType, WpsClientStatus, DocumentInfo, WorkbookInfo, PresentationInfo } from '../types/wps';
export declare function setPptTarget(name?: string): void;
export declare function getPptTarget(): string | undefined;
/**
 * WPS客户端类 - Windows COM 通信
 */
export declare class WpsClient {
    private status;
    constructor(_config?: Partial<WpsEndpointConfig>);
    /**
     * 调用WPS接口（跨平台）
     */
    invokeAction<T = unknown>(action: string, params?: Record<string, unknown>): Promise<WpsApiResponse<T>>;
    /**
     * 兼容旧API
     */
    callApi<T = unknown>(request: WpsApiRequest): Promise<WpsApiResponse<T>>;
    /**
     * 检查WPS连接状态
     */
    checkConnection(): Promise<boolean>;
    /**
     * 获取客户端状态
     */
    getStatus(): WpsClientStatus;
    getActiveWorkbook(): Promise<WorkbookInfo | null>;
    getCellValue(sheet: string | number, row: number, col: number): Promise<unknown>;
    setCellValue(sheet: string | number, row: number, col: number, value: unknown): Promise<boolean>;
    getRangeData(sheet: string | number | undefined, range: string): Promise<unknown[][]>;
    setRangeData(sheet: string | number | undefined, range: string, data: unknown[][]): Promise<boolean>;
    setFormula(sheet: string | number, row: number, col: number, formula: string): Promise<boolean>;
    getActiveDocument(): Promise<DocumentInfo | null>;
    createDocument(): Promise<boolean>;
    insertText(text: string, position?: number): Promise<boolean>;
    getDocumentText(): Promise<string>;
    getActivePresentation(): Promise<PresentationInfo | null>;
    createPresentation(): Promise<boolean>;
    addSlide(layout?: string): Promise<boolean>;
    executeMethod<T = unknown>(method: string, params?: Record<string, unknown>, appType?: WpsAppType): Promise<WpsApiResponse<T>>;
    openFile(filePath: string, _appType?: WpsAppType): Promise<boolean>;
    saveFile(_appType?: WpsAppType): Promise<boolean>;
    saveFileAs(filePath: string, _appType?: WpsAppType): Promise<boolean>;
}
export declare const wpsClient: WpsClient;
export default WpsClient;
//# sourceMappingURL=wps-client.d.ts.map
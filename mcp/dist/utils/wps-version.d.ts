/**
 * Input: WPS 应用名与"真实版本号"字符串（取自 exe 文件版本，不是 Application.Version）
 * Output: 版本解析结果与兼容性判断
 * Pos: 运行时 WPS 版本前置检查（wps_status 使用）。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */
/** 本插件支持的 WPS 下限，与 README / scripts/doctor.mjs 保持一致。 */
export declare const MIN_WPS_MAJOR = 12;
export declare const MIN_WPS_MINOR = 1;
/**
 * "12.1.0.28488" / "12,1,0,28488" -> [12, 1, 0, 28488]。
 * 逗号是 Windows 文件版本（FileVersionInfo）的写法，桥那边已经归一化成点号，这里两种都收。
 * 空串、非字符串、解析不出任何数字时返回空数组，调用方据此判断「读不到版本」，而不是当成 0。
 */
export declare function parseWpsVersion(raw: unknown): number[];
/** 只有至少读到 major.minor 才下结论；读不到就当作「未知」，不误报太旧。 */
export declare function wpsVersionTooOld(version: number[]): boolean;
/** 版本过旧时给出面向用户的一句中文提醒；版本读不到或够新时返回 undefined。 */
export declare function wpsCompatibilityWarning(appName: string | undefined, rawVersion: unknown): string | undefined;
//# sourceMappingURL=wps-version.d.ts.map
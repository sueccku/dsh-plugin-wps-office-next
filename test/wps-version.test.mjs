// P3: runtime WPS version pre-flight (wps_status). The decision is pure, so it is covered here without
// any WPS installation; the live status surface is exercised by test/deprecated.test.mjs.
// Run: node test/wps-version.test.mjs
import { createRequire } from "node:module";
import path from "node:path";
// P4：改用共用 harness（新测试的样板；老文件按需增量迁移）。
import { createReport } from "./lib/harness.mjs";

const require = createRequire(import.meta.url);
const { MIN_WPS_MAJOR, MIN_WPS_MINOR, parseWpsVersion, wpsVersionTooOld, wpsCompatibilityWarning } =
  require(path.resolve("mcp/dist/utils/wps-version.js"));

const { check, finish } = createReport("WPS VERSION TESTS");

check("supported floor is 12.1", MIN_WPS_MAJOR === 12 && MIN_WPS_MINOR === 1, MIN_WPS_MAJOR + "." + MIN_WPS_MINOR);
check("a real version parses into numbers", JSON.stringify(parseWpsVersion("12.1.0.28488")) === "[12,1,0,28488]", JSON.stringify(parseWpsVersion("12.1.0.28488")));
check("a Windows file version with commas parses too", JSON.stringify(parseWpsVersion("12,1,0,28488")) === "[12,1,0,28488]", JSON.stringify(parseWpsVersion("12,1,0,28488")));
check("a two-part version is enough", JSON.stringify(parseWpsVersion("11.8")) === "[11,8]", JSON.stringify(parseWpsVersion("11.8")));
check("unreadable input yields no numbers", parseWpsVersion("").length === 0 && parseWpsVersion(undefined).length === 0 && parseWpsVersion(12.1).length === 0, JSON.stringify([parseWpsVersion(""), parseWpsVersion(undefined), parseWpsVersion(12.1)]));
check("12.1 is accepted", wpsVersionTooOld(parseWpsVersion("12.1.0.1")) === false, "12.1.0.1");
check("12.2 is accepted", wpsVersionTooOld(parseWpsVersion("12.2.0")) === false, "12.2.0");
check("13.0 is accepted", wpsVersionTooOld(parseWpsVersion("13.0.1")) === false, "13.0.1");
check("12.0 is rejected", wpsVersionTooOld(parseWpsVersion("12.0.9")) === true, "12.0.9");
check("11.8 is rejected", wpsVersionTooOld(parseWpsVersion("11.8.2")) === true, "11.8.2");
check("an unknown version is not called too old", wpsVersionTooOld(parseWpsVersion("1")) === false && wpsVersionTooOld([]) === false, "1 / []");
const warn = String(wpsCompatibilityWarning("WPS 文字", "11.8.2"));
check("the warning names the version and the floor", warn.includes("11.8.2") && warn.includes("12.1") && warn.includes("WPS 文字"), warn.slice(0, 90));
check("a supported version has no warning", wpsCompatibilityWarning("WPS 文字", "12.1.0.28488") === undefined, String(wpsCompatibilityWarning("WPS 文字", "12.1.0.28488")));
check("an unreadable version has no warning", wpsCompatibilityWarning("WPS 表格", "") === undefined, "");

finish();

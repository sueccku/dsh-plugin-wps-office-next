// 按 scripts/lib/static-tests.mjs 的清单跑所有**不需要 WPS** 的测试文件；CI 与本地共用这份清单。
//
// 用法：node scripts/run-static-tests.mjs [--only=plugin,style-names]
// 关键（与 run-tests.ps1 一致）：把插件自己发布的两个环境变量钉到仓库内路径，否则从「已装本插件」的
// DSH 会话里起终端时测的是 profile 里那份副本，plugin.test.mjs 会报一处假失败。
import { spawnSync } from "node:child_process";
import { resolve } from "node:path";
import { STATIC_TESTS } from "./lib/static-tests.mjs";

const env = {
  ...process.env,
  WPS_OFFICE_MCP_ENTRY: resolve("mcp/dist/index.js"),
  WPS_OFFICE_HOST_SCRIPT: resolve("host/wps-com-host.ps1"),
};
const onlyArg = process.argv.find((a) => a.startsWith("--only="));
const only = onlyArg ? onlyArg.slice("--only=".length).split(",").filter(Boolean) : null;
const files = only ? STATIC_TESTS.filter((f) => only.some((o) => f.includes(o))) : STATIC_TESTS;

let failed = 0;
for (const file of files) {
  const started = Date.now();
  const r = spawnSync(process.execPath, [file], { encoding: "utf8", env, windowsHide: true });
  const output = (r.stdout || "") + (r.stderr || "");
  const pass = (output.match(/^PASS /gm) || []).length;
  const fail = (output.match(/^FAIL /gm) || []).length;
  const bad = r.status !== 0 || fail > 0;
  if (bad) failed++;
  console.log(
    (bad ? "FAIL " : "ok   ") + file.padEnd(38) +
    " PASS=" + String(pass).padEnd(4) + " FAIL=" + String(fail).padEnd(3) +
    " exit=" + String(r.status).padEnd(4) + " " + ((Date.now() - started) / 1000).toFixed(1) + "s",
  );
  if (bad) {
    const lines = output.split(/\r?\n/).filter((l) => /^FAIL |Error|error/i.test(l)).slice(0, 6);
    if (lines.length) console.log("      " + lines.join("\n      "));
  }
}
console.log("static tests: " + (files.length - failed) + "/" + files.length + " green" + (only ? " (filtered)" : ""));
process.exit(failed ? 1 : 0);

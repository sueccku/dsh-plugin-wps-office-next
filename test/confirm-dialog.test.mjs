// S3 residual: do any destructive actions pop a modal confirmation dialog?
//
// A modal dialog blocks the STA host, so an unsuppressed one would hang the action until the client
// timeout. This runs the scenarios that exercise every destructive action from docs/destructive-
// operations.md while a background watcher enumerates visible windows and records any whose class
// starts with "Qt" (WPS renders its modal dialogs as Qt windows; document frames are XLMAIN/OpusApp/
// PP12FrameClass). Passing children + zero recorded dialogs is the acceptance.
//
// FIXES 92（P1）：看门狗是**全机**枚举的，只按「类名以 Qt 开头」判定 —— 本机**微信**（Weixin.exe）的窗口
// 类名正是 Qt51514QWindowIcon，与 WPS 对话框同类名前缀，于是这条断言在任何装了 Qt 应用的机器上**永远红**
// （实测：窗口在看门狗启动后 25 ms 就被记账，早于任何场景开跑）。这里改用**基线**：只对"场景开跑之后
// 新出现"的窗口做判定。注意不能反过来"看到 DIALOG 就忽略" —— 那会让看门狗失效伪装成没有弹框。
// Run: node test/confirm-dialog.test.mjs   (needs WPS)
import { spawn, spawnSync } from "node:child_process";
import { writeFileSync, readFileSync, existsSync, rmSync } from "node:fs";

const WATCHER = "test/.artifacts/dialog-watch-wps.ps1";
const LOG = "test/.artifacts/dialog-watch-wps.log";
const STOP = LOG + ".stop";
writeFileSync(WATCHER, "param([string]$Log = 'test/.artifacts/dialog-watch.log')\nAdd-Type -TypeDefinition @'\nusing System;\nusing System.Text;\nusing System.Runtime.InteropServices;\npublic class WpsWatch {\n  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);\n  [DllImport(\"user32.dll\")] public static extern bool EnumWindows(EnumProc cb, IntPtr lParam);\n  [DllImport(\"user32.dll\")] public static extern bool IsWindowVisible(IntPtr hWnd);\n  [DllImport(\"user32.dll\", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int max);\n}\n'@\n$stopFile = $Log + '.stop'\nif (Test-Path $stopFile) { Remove-Item -LiteralPath $stopFile -Force }\nSet-Content -LiteralPath $Log -Value ('watcher started ' + (Get-Date).ToString('o'))\n$seen = @{}\nwhile (-not (Test-Path $stopFile)) {\n  $found = New-Object System.Collections.ArrayList\n  $cb = [WpsWatch+EnumProc]{\n    param($h, $l)\n    if ([WpsWatch]::IsWindowVisible($h)) {\n      $sb = New-Object System.Text.StringBuilder 256\n      [void][WpsWatch]::GetClassName($h, $sb, 256)\n      $c = $sb.ToString()\n      if ($c -like 'Qt*') { [void]$found.Add($c) }\n    }\n    return $true\n  }\n  [void][WpsWatch]::EnumWindows($cb, [IntPtr]::Zero)\n  foreach ($c in $found) {\n    if (-not $seen.ContainsKey($c)) {\n      $seen[$c] = $true\n      Add-Content -LiteralPath $Log -Value ('DIALOG ' + (Get-Date).ToString('o') + ' ' + $c)\n    }\n  }\n  Start-Sleep -Milliseconds 200\n}\nAdd-Content -LiteralPath $Log -Value ('watcher stopped ' + (Get-Date).ToString('o'))\n");
if (existsSync(LOG)) rmSync(LOG, { force: true });
if (existsSync(STOP)) rmSync(STOP, { force: true });

const results = [];
function check(name, ok, detail) { results.push({ name, ok }); console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : "")); }

const watcher = spawn("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", WATCHER, "-Log", LOG], { windowsHide: true, stdio: "ignore" });
check("dialog watcher started", !!watcher.pid, "pid=" + watcher.pid);
await new Promise((r) => setTimeout(r, 2500)); // let Add-Type compile before the first scenario
// FIXES 92（P1）：场景开跑的时刻就是基线分界线 —— 早于它的 DIALOG 都是"开跑前就存在的"（如微信的窗口）。
const scenariosStart = Date.now();

// These three scenarios exercise all 25 destructive actions from the inventory.
const SCENARIOS = [
  ["test/destructive-guard.test.mjs", "destructive-guard"],
  ["test/excel-advanced.test.mjs", "excel-advanced"],
  ["test/excel-page-setup.test.mjs", "excel-page-setup"],
];
for (const [file, label] of SCENARIOS) {
  const run = spawnSync(process.execPath, [file], { encoding: "utf8", windowsHide: true });
  const tail = String(run.stdout || "").trim().split(/\r?\n/).slice(-1)[0] || "";
  check(label + " scenario passed under the watcher", run.status === 0, "exit=" + run.status + " " + tail.slice(0, 80));
}

// Stop the watcher and read what it saw.
writeFileSync(STOP, "");
await new Promise((r) => setTimeout(r, 1500));
if (!watcher.killed) { try { watcher.kill(); } catch { /* already gone */ } }

const log = existsSync(LOG) ? readFileSync(LOG, "utf8") : "";
// 日志行形如 `DIALOG <ISO 时间> <类名>`。
const dialogLines = log.split(/\r?\n/).filter((l) => l.startsWith("DIALOG "));
const parsed = dialogLines.map((l) => ({ line: l, at: Date.parse(l.split(" ")[1] || "") }));
// 时间戳解析不出来的行**不能**被当成"基线里的"而悄悄放过 —— 那会让"看门狗/日志格式坏了"伪装成"没有弹框"，
// 所以单独断言（FIXES 92 / P1）。
const unparsed = parsed.filter((d) => Number.isNaN(d.at));
const observed = parsed.filter((d) => !Number.isNaN(d.at));
const preexisting = observed.filter((d) => d.at < scenariosStart);
const dialogs = observed.filter((d) => d.at >= scenariosStart);
check("watcher reported a clean start", log.includes("watcher started"), "");
check(
  "every DIALOG line carries a parseable timestamp",
  unparsed.length === 0,
  unparsed.length ? unparsed.map((d) => d.line).join(" | ") : dialogLines.length + " dialog line(s), all parseable"
);
check("no modal confirmation dialog appeared", dialogs.length === 0, (dialogs.length ? dialogs.map((d) => d.line).join(" | ") : "0 new Qt windows while " + SCENARIOS.length + " scenarios ran") + "（开跑前已存在、按基线忽略 " + preexisting.length + " 个）");

rmSync(WATCHER, { force: true });
rmSync(STOP, { force: true });
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "CONFIRM DIALOG TESTS OK (" + results.length + ")" : "CONFIRM DIALOG TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);

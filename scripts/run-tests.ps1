# Run every test/*.test.mjs and reap the WPS instances a test file leaves behind.
#
# Why the reaping exists (FIXES 65): on Windows Node puts every spawned child into a job object and
# tears the whole tree down when that child is killed. The test files end by hard-killing the MCP
# server, so the resident COM host dies with it and never reaches its own cleanup; whatever WPS
# application it started is orphaned (measured: one Word document = 3-5 wps.exe processes, ~10 per
# full run).
#
# What counts as ours (FIXES 106): a PID baseline, not the window title.
#   - Everything alive when the run starts belongs to someone else (you, or a DSH session that already
#     holds a WPS instance) and is NEVER killed - the suite may well be reusing it.
#   - Anything that appears later was started by a resident host inside this run, so it is ours, with
#     or without a window. The old "no window title" rule only approximated that: it missed every
#     instance that shows a window (every Excel instance, now that FIXES 104 makes the window visible)
#     and it would kill a window-less WPS process that was never ours.
#   - -AllWps also clears the baseline before and after the run. That is what a dedicated test machine
#     wants; on a machine where you keep your own WPS open, leave it off. There is deliberately NO
#     environment-variable switch for it: a mode that force-kills every WPS process must be visible in
#     the command line you actually ran, not inherited from ambient state.
#   - The set is et/wps/wpp PLUS wpscloudsvr, WPS's own cloud service. Measured here: one instance
#     started through COM is a family - wps.exe (frame, started by svchost/DCOM, which is exactly why
#     the job object never reaches it), et.exe/wpp.exe, wpscloudsvr, and wpscloudsvr's wps.exe workers -
#     and wpscloudsvr restarts a killed worker within about a second. Reaping only et/wps/wpp therefore
#     always leaves one process behind (the ORPHANS_LEFT=1 in every historical run): WPS's service, not
#     a test orphan. Killing wpscloudsvr with the rest reaches a real zero.
#   - Nothing here touches COM, so a wedged WPS (modal dialog, RPC_E_CALL_REJECTED) can never hang the
#     suite - contrast Invoke-WpsOrphanReclaim, whose COM calls have no timeout (FIXES 102).
#
# Sibling tool: scripts\reap-wps.ps1 cleans a machine by hand (single-file runs, wedged instances).
#
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\run-tests.ps1 [-Filter <wildcard>] [-KeepOrphans] [-AllWps] [-Static] [-PerFileTimeoutSec <sec>]
param([string]$Filter = '', [switch]$KeepOrphans, [switch]$AllWps, [switch]$Static, [int]$PerFileTimeoutSec = 600)
$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# 每个测试文件都必须驱动**仓库里的**宿主与 MCP 入口。从「已装本插件」的 DSH 会话里起终端时，plugin.js 会把
# WPS_OFFICE_MCP_ENTRY / WPS_OFFICE_HOST_SCRIPT 指到 profile 里那份副本，测试于是测了**上一个发布版**——
# 实测表现是宿主沿用旧参数白名单（unknown parameter confirm），红得让人以为是自己刚改的代码坏了。
# 光删掉这两个变量不够（plugin.test.mjs 断言的正是「plugin.js 发布的入口 == 包内入口」，副本路径会让它假红），
# 所以在**每个测试文件 spawn 之前**显式把它们指向仓库内路径（见下面的循环）。这样 47 个文件都确定性地跑仓库代码（FIXES 87）。
$env:WPS_OFFICE_MCP_ENTRY = [System.IO.Path]::Combine($root, 'mcp', 'dist', 'index.js')
$env:WPS_OFFICE_HOST_SCRIPT = [System.IO.Path]::Combine($root, 'host', 'wps-com-host.ps1')

function Get-WpsProcesses { @(Get-Process et,wps,wpp,wpscloudsvr -ErrorAction SilentlyContinue) }
# PIDs alive before the run. Anything outside this set was started by the suite, so it is ours.
$script:BaselinePids = @{}
function Set-WpsBaseline {
    $script:BaselinePids = @{}
    foreach ($p in (Get-WpsProcesses)) { $script:BaselinePids[[int]$p.Id] = $true }
}
function Get-NewWps { @((Get-WpsProcesses) | Where-Object { -not $script:BaselinePids.ContainsKey([int]$_.Id) }) }
function Clear-Orphans {
    if ($KeepOrphans) { return 0 }
    $new = @(Get-NewWps)
    foreach ($p in $new) { try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { } }
    return $new.Count
}
function Clear-AllWps {
    $all = @(Get-WpsProcesses)
    foreach ($p in $all) { try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { } }
    return $all.Count
}

if ($AllWps) {
    $cleared = Clear-AllWps
    if ($cleared -gt 0) { "AllWps: cleared $cleared pre-existing WPS process(es) before the run" }
    Start-Sleep -Milliseconds 600
}
Set-WpsBaseline
$files = Get-ChildItem (Join-Path $root 'test') -Filter '*.test.mjs' | Sort-Object Name
if ($Filter) { $files = $files | Where-Object { $_.Name -like $Filter } }
if ($Static) {
    # 不需要 WPS 的那批：清单的**唯一来源**是 scripts/lib/static-tests.mjs（CI 的 static job 读同一份）。
    $names = (& node -e "import('./scripts/lib/static-tests.mjs').then(m => console.log(m.STATIC_TESTS.join(',')))") -split ',' | ForEach-Object { Split-Path $_ -Leaf }
    $files = $files | Where-Object { $names -contains $_.Name }
}
# 「整轮」才会写权威快照、才检查断言底线：-Filter / -Static 是子集运行，拿它覆盖 test/summary.json
# （文档数字的唯一来源）或用底线判失败都没有意义。
$isFullRun = (-not $Filter) -and (-not $Static)
$totalPass = 0; $totalFail = 0; $reaped = 0; $badFiles = @()
# FIXES 91（B4）：把这次整轮的权威计数落成一份机器可读的快照，供 docs/current-numbers.md 与
# spec/current-numbers.json 引用 —— 文档里的数字从此只有一个来源（这个文件由跑测试产生，不手写）。
$summaryRows = @()
foreach ($f in $files) {
    $env:WPS_OFFICE_MCP_ENTRY = [System.IO.Path]::Combine($root, 'mcp', 'dist', 'index.js')
    $env:WPS_OFFICE_HOST_SCRIPT = [System.IO.Path]::Combine($root, 'host', 'wps-com-host.ps1')
    # P4：每个文件都有**墙钟上限**，并且用 .NET 进程对象同时拿输出与退出码。
    #   为什么不用 Start-Process -PassThru：Windows PowerShell 5.1 下 WaitForExit(ms) 之后
    #   $proc.ExitCode 是 $null（[int]$null → 0），于是"退出码非零但没打印 FAIL"的测试会被当成绿
    #   （2026-10-10 实测）。.NET 的 Process 在 5.1 上行为正确。
    #   输出编码必须显式设成 UTF-8：node 写 UTF-8，按 ANSI 读会把中文错配、偶发吞掉换行、少计 PASS。
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'node'
    $psi.Arguments = '"' + $f.FullName + '"'
    $psi.WorkingDirectory = $root
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $null = $proc.Start()
    $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
    $stderrTask = $proc.StandardError.ReadToEndAsync()
    $finished = $proc.WaitForExit($PerFileTimeoutSec * 1000)
    $timedOut = -not $finished
    if ($timedOut) { & taskkill /T /F /PID $proc.Id 2>&1 | Out-Null; Start-Sleep -Milliseconds 800 }
    $sw.Stop()
    $exitCode = if ($finished) { [int]$proc.ExitCode } else { -1 }
    $out = [string]$stdoutTask.Result + "`n" + [string]$stderrTask.Result
    $p = ([regex]::Matches($out, '(?m)^PASS ')).Count
    $x = ([regex]::Matches($out, '(?m)^FAIL ')).Count
    $totalPass += $p; $totalFail += $x
    $reaped += Clear-Orphans
    $tail = ($out -split "`r?`n" | Where-Object { $_ -match 'FAIL|ERROR|OK \(|FAILED' } | Select-Object -Last 4) -join ' | '
    if ($timedOut) { $tail = "TIMEOUT after $($PerFileTimeoutSec)s | " + $tail }
    "{0,-42} PASS={1,-4} FAIL={2,-3} exit={3,-3} {4,6:N0}s  {5}" -f $f.Name, $p, $x, $exitCode, $sw.Elapsed.TotalSeconds, $tail
    $summaryRows += [ordered]@{ file = $f.Name; pass = $p; fail = $x; exit = $exitCode; timeout = $timedOut; seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1) }
    if ($exitCode -ne 0 -or $x -gt 0) { $badFiles += $f.Name }
}
# Authoritative final reading: kill what is new, give lazily starting/exiting helpers a moment to show
# up, kill again - so ORPHANS_LEFT=0 really means "nothing this run started is still alive".
for ($i = 0; $i -lt 3; $i++) {
    $null = Clear-Orphans
    Start-Sleep -Milliseconds 1200
    if (@(Get-NewWps).Count -eq 0) { break }
}
$left = @(Get-NewWps).Count
if ($AllWps) {
    $null = Clear-AllWps
    Start-Sleep -Milliseconds 800
    $left = @(Get-WpsProcesses).Count
}
""
"TOTAL PASS=$totalPass FAIL=$totalFail FILES=$($files.Count) REAPED=$reaped ORPHANS_LEFT=$left"
if ($badFiles.Count -gt 0) { 'BAD FILES: ' + ($badFiles -join ', ') } else { 'ALL TEST FILES GREEN' }

$summary = [ordered]@{
    generatedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ssK')
    files = $files.Count
    assertions = $totalPass + $totalFail
    pass = $totalPass
    fail = $totalFail
    badFiles = @($badFiles)
    perFile = $summaryRows
}
if ($isFullRun) {
    $summaryPath = Join-Path $root 'test/summary.json'
    [System.IO.File]::WriteAllText($summaryPath, ($summary | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding($false)))
    "SUMMARY written: test/summary.json (assertions=$($summary.assertions) files=$($summary.files))"
} else {
    "SUMMARY skipped (subset run: -Filter / -Static 不覆盖权威快照)"
}

# ---- P6：断言数 / 文件数的底线 ---------------------------------------------------------------
# 低于底线说明有测试被顺手删掉或漏跑（这在以前完全看不出来）。要降低必须显式改 test/baseline.json。
$baselineViolations = @()
$baselinePath = Join-Path $root 'test/baseline.json'
if (-not $isFullRun) {
    'BASELINE skipped (subset run)'
} elseif (Test-Path $baselinePath) {
    # test/baseline.json 是**无 BOM** 的 UTF-8（还带中文 note），而 Windows PowerShell 5.1 的
    # Get-Content 默认按 ANSI 解码：中文 note 解坏后 ConvertFrom-Json 抛错，底线比较**静默失效**
    # （打印 `>= ` 后面是空的，等于门禁本身成为摆设）。显式 -Encoding UTF8 读；真解析失败就判违规。
    $baseline = $null
    try { $baseline = Get-Content -LiteralPath $baselinePath -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { $baselineViolations += 'cannot parse the baseline file: ' + $_.Exception.Message }
    if ($null -eq $baseline) {
        'BASELINE VIOLATION: ' + ($baselineViolations -join '; ')
    } else {
        if ($summary.assertions -lt [int]$baseline.minAssertions) { $baselineViolations += "assertions $($summary.assertions) < $($baseline.minAssertions)" }
        if ($summary.files -lt [int]$baseline.minFiles) { $baselineViolations += "files $($summary.files) < $($baseline.minFiles)" }
        if ($baselineViolations.Count -gt 0) { 'BASELINE VIOLATION: ' + ($baselineViolations -join '; ') }
        else { "BASELINE OK (assertions=$($summary.assertions) >= $($baseline.minAssertions), files=$($summary.files) >= $($baseline.minFiles))" }
    }
}
if ($totalFail -gt 0 -or $badFiles.Count -gt 0 -or $baselineViolations.Count -gt 0) { exit 1 }

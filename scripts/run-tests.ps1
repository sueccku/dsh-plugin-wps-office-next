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
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\run-tests.ps1 [-Filter <wildcard>] [-KeepOrphans] [-AllWps]
param([string]$Filter = '', [switch]$KeepOrphans, [switch]$AllWps)
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
$totalPass = 0; $totalFail = 0; $reaped = 0; $badFiles = @()
# FIXES 91（B4）：把这次整轮的权威计数落成一份机器可读的快照，供 docs/current-numbers.md 与
# spec/current-numbers.json 引用 —— 文档里的数字从此只有一个来源（这个文件由跑测试产生，不手写）。
$summaryRows = @()
foreach ($f in $files) {
    $env:WPS_OFFICE_MCP_ENTRY = [System.IO.Path]::Combine($root, 'mcp', 'dist', 'index.js')
    $env:WPS_OFFICE_HOST_SCRIPT = [System.IO.Path]::Combine($root, 'host', 'wps-com-host.ps1')
    $out = & node $f.FullName 2>&1 | Out-String
    $p = ([regex]::Matches($out, '(?m)^PASS ')).Count
    $x = ([regex]::Matches($out, '(?m)^FAIL ')).Count
    $totalPass += $p; $totalFail += $x
    $reaped += Clear-Orphans
    $tail = ($out -split "`r?`n" | Where-Object { $_ -match 'FAIL|ERROR|OK \(|FAILED' } | Select-Object -Last 4) -join ' | '
    "{0,-42} PASS={1,-4} FAIL={2,-3} exit={3}  {4}" -f $f.Name, $p, $x, $LASTEXITCODE, $tail
    $summaryRows += [ordered]@{ file = $f.Name; pass = $p; fail = $x; exit = $LASTEXITCODE }
    if ($LASTEXITCODE -ne 0 -or $x -gt 0) { $badFiles += $f.Name }
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
$summaryPath = Join-Path $root 'test/summary.json'
[System.IO.File]::WriteAllText($summaryPath, ($summary | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding($false)))
"SUMMARY written: test/summary.json (assertions=$($summary.assertions) files=$($summary.files))"

if ($totalFail -gt 0 -or $badFiles.Count -gt 0) { exit 1 }

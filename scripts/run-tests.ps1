# Run every test/*.test.mjs and reap the headless WPS instances a test file leaves behind.
#
# Why the reaping exists (FIXES 65): on Windows Node puts every spawned child into a job object and
# tears the whole tree down when that child is killed. The test files end by hard-killing the MCP
# server, so the resident COM host dies with it and never reaches its own cleanup; whatever WPS
# application it started is orphaned (measured: one Word document = 3-5 wps.exe processes, ~10 per
# full run). Those orphans are headless - no visible document window - and nothing in the suite reuses
# them, so after each file we close the ones whose window title is empty. A real WPS window you have
# open has a title and is never touched.
#
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\run-tests.ps1 [-Filter <wildcard>] [-KeepOrphans]
param([string]$Filter = '', [switch]$KeepOrphans)
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

function Get-Orphans {
    Get-Process wps,et,wpp -ErrorAction SilentlyContinue |
        Where-Object { [string]$_.MainWindowTitle -eq '' }
}
function Clear-Orphans {
    if ($KeepOrphans) { return 0 }
    $o = @(Get-Orphans)
    foreach ($p in $o) { try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { } }
    return $o.Count
}
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
$left = @(Get-Orphans).Count
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

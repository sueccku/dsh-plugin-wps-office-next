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
foreach ($f in $files) {
    $out = & node $f.FullName 2>&1 | Out-String
    $p = ([regex]::Matches($out, '(?m)^PASS ')).Count
    $x = ([regex]::Matches($out, '(?m)^FAIL ')).Count
    $totalPass += $p; $totalFail += $x
    $reaped += Clear-Orphans
    $tail = ($out -split "`r?`n" | Where-Object { $_ -match 'FAIL|ERROR|OK \(|FAILED' } | Select-Object -Last 4) -join ' | '
    "{0,-42} PASS={1,-4} FAIL={2,-3} exit={3}  {4}" -f $f.Name, $p, $x, $LASTEXITCODE, $tail
    if ($LASTEXITCODE -ne 0 -or $x -gt 0) { $badFiles += $f.Name }
}
$left = @(Get-Orphans).Count
""
"TOTAL PASS=$totalPass FAIL=$totalFail FILES=$($files.Count) REAPED=$reaped ORPHANS_LEFT=$left"
if ($badFiles.Count -gt 0) { 'BAD FILES: ' + ($badFiles -join ', ') } else { 'ALL TEST FILES GREEN' }
if ($totalFail -gt 0 -or $badFiles.Count -gt 0) { exit 1 }

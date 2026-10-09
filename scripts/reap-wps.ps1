# Reap the WPS processes this repo's tests, probes and crashed sessions leave behind.
#
# Why this exists: the test runner cleans up between files, but nothing does when you drive a single
# test file, a script, or the resident host by hand. This is the explicit "clean machine" button for
# development. It is deliberately COM-free - Get-Process / Stop-Process only - so a wedged WPS
# (modal dialog, RPC_E_CALL_REJECTED) can never hang it. That is the opposite trade-off from the
# product path (Invoke-WpsOrphanReclaim), whose COM calls have no timeout and therefore must stay
# gentle (FIXES 102): here we are allowed to be blunt because tests only ever create throwaway
# documents.
#
# Modes (exactly one; with no mode it only LISTS, and kills nothing):
#   (none)              list every WPS process; kill nothing.
#   -All                every et/wps/wpp process PLUS wpscloudsvr. Meant for a dedicated test machine,
#                       where nothing running WPS is worth preserving. wpscloudsvr is WPS's cloud
#                       service: a COM-started instance is a family (wps.exe frame from svchost/DCOM,
#                       et.exe/wpp.exe, wpscloudsvr, and its wps.exe workers) and wpscloudsvr restarts
#                       a killed worker within ~1s - reaping only et/wps/wpp never reaches zero.
#   -Since <DateTime>   only processes that started at or after that moment, i.e. what the last few
#                       minutes of work left behind. Safe on a machine where you keep WPS open.
#   -DryRun             with -All or -Since: print the decision and every target, kill nothing.
#
# Byte convention: ASCII only, LF, no BOM - same as build-host-actions.ps1. If you ever add
# non-ASCII text here, give the file a UTF-8 BOM (Windows PowerShell 5.1 reads BOM-less files as
# ANSI/GBK) and update the convention list in scripts/lint.mjs.
#
# Usage:
#   powershell -NoProfile -File scripts\reap-wps.ps1                 # list only
#   powershell -NoProfile -File scripts\reap-wps.ps1 -All
#   powershell -NoProfile -File scripts\reap-wps.ps1 -Since (Get-Date).AddMinutes(-15)
param(
    [switch]$All,
    [datetime]$Since,
    [switch]$DryRun
)
$ErrorActionPreference = 'Continue'

# "Was -Since actually passed?" must be asked of the binder, not of the value: an unbound [datetime]
# parameter is not reliably comparable to [datetime]::MinValue, and getting that wrong here would
# silently turn the listing mode into the killing mode.
$mode = if ($All) { 'All' } elseif ($PSBoundParameters.ContainsKey('Since')) { 'Since' } else { 'List' }

# NOTE: never name a local $all - the [switch]$All parameter makes that the same variable (PowerShell
# variable names are case-insensitive) and assigning an array to a [switch] throws at runtime.
$procs = @(Get-Process et,wps,wpp,wpscloudsvr -ErrorAction SilentlyContinue)
$targets = if ($mode -eq 'Since') { @($procs | Where-Object { $_.StartTime -ge $Since }) } else { $procs }
$doKill = ($mode -ne 'List') -and (-not $DryRun)

if ($mode -eq 'List') {
    'no mode given: listing every WPS process; nothing is killed (use -All, or -Since <DateTime>)'
}
"mode: $mode  targets: $($targets.Count)"
foreach ($p in $targets) {
    $started = ''
    try { $started = $p.StartTime.ToString('yyyy-MM-dd HH:mm:ss') } catch { $started = '(unknown)' }
    '  {0}:{1} start={2} title=[{3}]' -f $p.ProcessName, $p.Id, $started, $p.MainWindowTitle
}

if (-not $doKill) {
    'nothing killed'
    exit 0
}

$killed = 0
foreach ($p in $targets) {
    try { Stop-Process -Id $p.Id -Force -ErrorAction Stop; $killed++ }
    catch { '  could not kill {0}:{1}: {2}' -f $p.ProcessName, $p.Id, $_.Exception.Message }
}
Start-Sleep -Milliseconds 800
$left = @(Get-Process et,wps,wpp,wpscloudsvr -ErrorAction SilentlyContinue).Count
"killed: $killed  wps processes left: $left"

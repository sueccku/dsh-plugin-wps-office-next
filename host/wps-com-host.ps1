# Input: newline-delimited JSON requests on stdin
# Output: newline-delimited JSON responses on stdout
# Pos: Resident WPS COM host. Replaces one PowerShell process per tool call.
#
# Protocol (one JSON object per line, UTF-8):
#   request  {"id": <any>, "action": "<camelCaseAction>", "params": <object>}
#   response {"id": <any>, "ok": <bool>, "result": {success, data|error}, "ms": <int>}
#   control  action "__shutdown" stops the loop.
#
# Only stdout carries protocol frames; the success stream of the action call is
# discarded and the error stream is left for the parent process to capture.

param(
    [string]$Root = $PSScriptRoot
)

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

$script:WpsResult = $null

. (Join-Path $Root 'wps-actions.ps1')

function Write-Frame([string]$text) {
    [Console]::Out.WriteLine($text)
    [Console]::Out.Flush()
}

function Send-Response($id, $ok, $result, $ms) {
    $payload = @{ id = $id; ok = [bool]$ok; result = $result; ms = $ms }
    $json = $null
    try { $json = $payload | ConvertTo-Json -Depth 12 -Compress } catch { $json = $null }
    if (-not $json) {
        $json = '{"id":0,"ok":false,"result":{"success":false,"error":"response serialization failed"},"ms":0}'
    }
    Write-Frame $json
}

# Warm the 248-case dispatch once so the first real request is not charged for JIT parsing.
try { $null = Invoke-WpsAction -Action '__warmup' -Params '{}' } catch { }
$script:WpsResult = $null

# The action layer is written against Windows PowerShell 5.1 semantics. Spawning the absolute
# path already pins the version; this assertion is the belt to that suspenders: if a different
# PowerShell is somehow used, fail loudly instead of misbehaving subtly.
$psVersion = $PSVersionTable.PSVersion.ToString()
$psMajor = [int]$PSVersionTable.PSVersion.Major
if ($psMajor -ne 5) {
    Write-Frame ('{"ready":false,"protocol":1,"error":"unsupported PowerShell ' + $psVersion + '; this host requires Windows PowerShell 5.1"}')
    exit 1
}

Write-Frame ('{"ready":true,"pid":' + $PID + ',"protocol":1,"psVersion":"' + $psVersion + '","actions":"wps-actions.ps1"}')

while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }
    $line = $line.Trim()
    if ($line.Length -eq 0) { continue }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $req = $null
    try { $req = $line | ConvertFrom-Json } catch { Send-Response 0 $false @{ success = $false; error = 'invalid request json' } 0; continue }

    $id = 0
    if ($null -ne $req.id) { $id = $req.id }
    $action = [string]$req.action
    if (-not $action) { Send-Response $id $false @{ success = $false; error = 'missing action' } 0; continue }
    if ($action -eq '__env') {
        Send-Response $id $true @{ success = $true; data = @{
            psVersion = $PSVersionTable.PSVersion.ToString()
            psMajor = [int]$PSVersionTable.PSVersion.Major
            pid = $PID
            exe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        } } 0
        continue
    }

    # Contract probe: answer what an action accepts without executing anything. Used by
    # scripts/param-contract.mjs to compare every tool schema against the bridge in one pass.
    if ($action -eq '__validateParams') {
        $target = [string]$req.params.action
        $accepted = $script:ActionParamKeys[$target]
        $given = @()
        if ($null -ne $req.params.keys) { $given = @($req.params.keys) }
        $unknown = @()
        if ($null -ne $accepted) { $unknown = @($given | Where-Object { $accepted -notcontains $_ }) }
        Send-Response $id $true @{ success = $true; data = @{
            action = $target
            validated = ($null -ne $accepted)
            accepted = @($accepted)
            given = $given
            unknown = $unknown
        } } 0
        continue
    }

    if ($action -eq '__shutdown') { Send-Response $id $true @{ success = $true; data = @{ message = 'shutdown' } } 0; break }

    $paramsJson = '{}'
    if ($null -ne $req.params) {
        try { $paramsJson = $req.params | ConvertTo-Json -Depth 12 -Compress } catch { $paramsJson = '{}' }
    }

    $script:WpsResult = $null
    try {
        $null = Invoke-WpsAction -Action $action -Params $paramsJson
    } catch {
        $script:WpsResult = @{ success = $false; error = $_.Exception.Message; errorType = $_.Exception.GetType().Name }
    }

    $sw.Stop()
    $result = $script:WpsResult
    if ($null -eq $result) { $result = @{ success = $false; error = ('action produced no result: ' + $action) } }
    $ok = $true
    if ($result.success -eq $false) { $ok = $false }
    Send-Response $id $ok $result $sw.ElapsedMilliseconds
}

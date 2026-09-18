# Input: newline-delimited JSON requests on stdin
# Output: newline-delimited JSON responses on stdout
# Pos: Resident WPS COM host. Replaces one PowerShell process per tool call.
#      Owns the single-instance lease (S2): WPS is one shared instance, so two hosts must never
#      drive it at the same time; the second one says so in Chinese instead of interleaving.
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

# ==================== Single-instance lease (S2) ====================
# Why: WPS exposes one shared instance to every automation client. Two DSH sessions each starting
# their own host used to interleave calls into that instance, which looks like "randomly stuck".
# A named mutex makes the second host refuse with a readable reason instead of racing.
#
# The mutex alone is not enough: a host killed while blocked inside COM lingers just long enough
# for the next one to arrive, and a host whose MCP server is gone would hold the lease forever.
# So the lease carries a state file (owner pid, owning client pid, heartbeat) and the newcomer
# takes over only when the recorded owner is provably dead: its process is gone, its client is
# gone, or its heartbeat stopped moving.
$script:MutexName = 'Local\dsh-plugin-wps-office-next-com-host'
$script:StateDir = Join-Path $env:USERPROFILE '.wps-office-mcp'
$script:StateFile = Join-Path $script:StateDir 'com-host.json'
$script:StaleHeartbeatSeconds = 120
$script:StartedUtc = (Get-Date).ToUniversalTime().ToString('o')
$script:BusySinceUtc = ''
$script:ClientPid = 0
if ($env:WPS_OFFICE_CLIENT_PID) { try { $script:ClientPid = [int]$env:WPS_OFFICE_CLIENT_PID } catch { $script:ClientPid = 0 } }
$script:HostMutex = $null

function Update-HostState([string]$phase, [string]$action = '', [int]$ms = 0) {
    # Best effort by design: bookkeeping must never take an action down.
    $snapshot = [ordered]@{
        phase = $phase
        hostPid = $PID
        clientPid = $script:ClientPid
        startedUtc = $script:StartedUtc
        updatedUtc = (Get-Date).ToUniversalTime().ToString('o')
        busySinceUtc = $script:BusySinceUtc
        lastAction = $action
        lastMs = $ms
    }
    try {
        if (-not (Test-Path $script:StateDir)) { New-Item -ItemType Directory -Force -Path $script:StateDir | Out-Null }
        $tmp = $script:StateFile + '.' + $PID + '.tmp'
        [System.IO.File]::WriteAllText($tmp, ($snapshot | ConvertTo-Json -Compress), (New-Object System.Text.UTF8Encoding($false)))
        Move-Item -LiteralPath $tmp -Destination $script:StateFile -Force
    } catch { }
}

function Read-HostState {
    try {
        if (-not (Test-Path $script:StateFile)) { return $null }
        return (Get-Content -LiteralPath $script:StateFile -Raw | ConvertFrom-Json)
    } catch { return $null }
}

function Test-ProcessAlive($processId) {
    if ($null -eq $processId) { return $false }
    try { return $null -ne (Get-Process -Id ([int]$processId) -ErrorAction SilentlyContinue) } catch { return $false }
}

function Get-StaleOwnerReason {
    # $null means "the owner looks alive" - the newcomer must refuse.
    $state = Read-HostState
    if ($null -eq $state) { return 'the lease file is missing or unreadable' }
    if (-not (Test-ProcessAlive $state.hostPid)) { return 'the owning host process is gone' }
    if ([int]$state.clientPid -gt 0 -and -not (Test-ProcessAlive $state.clientPid)) { return 'the DSH/MCP client that owned it is gone' }
    $age = 99999
    try { $age = ((Get-Date).ToUniversalTime() - ([datetime]::Parse([string]$state.updatedUtc)).ToUniversalTime()).TotalSeconds } catch { $age = 99999 }
    if ($age -gt $script:StaleHeartbeatSeconds) { return ('its heartbeat stopped ' + [int]$age + 's ago') }
    return $null
}

function Acquire-HostMutex {
    $created = $false
    $mutex = $null
    try { $mutex = New-Object System.Threading.Mutex($true, $script:MutexName, [ref]$created) } catch { return $null }
    if ($created) { return $mutex }
    $taken = $false
    try { $taken = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $taken = $true }
    if ($taken) { return $mutex }
    $reason = Get-StaleOwnerReason
    if ($null -eq $reason) { return $null }
    $state = Read-HostState
    if ($null -ne $state -and (Test-ProcessAlive $state.hostPid)) {
        # Only the stale host is stopped. WPS itself is never killed: its documents may hold
        # unsaved work that nobody asked this plugin to throw away.
        try { Stop-Process -Id ([int]$state.hostPid) -Force -ErrorAction SilentlyContinue } catch { }
    }
    try { if ($mutex.WaitOne(5000)) { return $mutex } } catch [System.Threading.AbandonedMutexException] { return $mutex } catch { }
    return $null
}

$script:HostMutex = Acquire-HostMutex
if ($null -eq $script:HostMutex) {
    $state = Read-HostState
    $holder = ''
    if ($null -ne $state -and $state.hostPid) { $holder = '（占用者 hostPid=' + [string]$state.hostPid + '，clientPid=' + [string]$state.clientPid + '，最近动作：' + [string]$state.lastAction + '）' }
    $message = '检测到另一个 DSH 会话或自动化程序正在控制 WPS：本插件同一时间只允许一个宿主，' +
        '以免两个会话交叉操作同一个 WPS 实例（表现为「莫名其妙卡住」）。' +
        '请先关闭那个会话或结束它的 WPS 操作，然后重试。' + $holder
    try { Write-Frame ((@{ ready = $false; protocol = 1; error = $message } | ConvertTo-Json -Compress)) } catch { Write-Frame '{"ready":false,"protocol":1,"error":"another COM host is already running"}' }
    exit 1
}
Update-HostState 'starting'

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

Write-Frame ('{"ready":true,"pid":' + $PID + ',"clientPid":' + $script:ClientPid + ',"protocol":1,"psVersion":"' + $psVersion + '","actions":"wps-actions.ps1"}')
Update-HostState 'idle'

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
            clientPid = $script:ClientPid
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
        $containers = $script:ActionNestedParams[$target]
        Send-Response $id $true @{ success = $true; data = @{
            action = $target
            validated = ($null -ne $accepted)
            accepted = @($accepted)
            given = $given
            unknown = $unknown
            containers = @($containers)
        } } 0
        continue
    }

    if ($action -eq '__shutdown') {
        $script:BusySinceUtc = ''
        # Release the WPS instances this host started; instances that were already running, or that
        # hold unsaved work, are deliberately left alone (FIXES 57).
        $closed = @()
        try { $closed = @(Close-WpsAppsStartedByUs) } catch { $closed = @() }
        Update-HostState 'stopped'
        Send-Response $id $true @{ success = $true; data = @{ message = 'shutdown'; closed = $closed } } 0
        break
    }

    $paramsJson = '{}'
    if ($null -ne $req.params) {
        try { $paramsJson = $req.params | ConvertTo-Json -Depth 12 -Compress } catch { $paramsJson = '{}' }
    }

    # Published before the call so a watcher can tell "busy since X" from "idle"; a request that
    # never returns therefore leaves a visible timestamp behind instead of a silent gap.
    $script:BusySinceUtc = (Get-Date).ToUniversalTime().ToString('o')
    Update-HostState 'busy' $action 0

    $script:WpsResult = $null
    try {
        $null = Invoke-WpsAction -Action $action -Params $paramsJson
    } catch {
        $script:WpsResult = @{ success = $false; error = (Format-WpsErrorText $_.Exception.Message $action); errorType = $_.Exception.GetType().Name }
    }

    $sw.Stop()
    $script:BusySinceUtc = ''
    Update-HostState 'idle' $action ([int]$sw.ElapsedMilliseconds)
    $result = $script:WpsResult
    if ($null -eq $result) { $result = @{ success = $false; error = ('action produced no result: ' + $action) } }
    $ok = $true
    if ($result.success -eq $false) { $ok = $false }
    Send-Response $id $ok $result $sw.ElapsedMilliseconds
}

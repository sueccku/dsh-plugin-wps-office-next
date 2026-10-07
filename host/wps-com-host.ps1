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

# FIXES 96（W5-1）：**PID 会被复用** —— 只看"这个进程号在不在进程表里"，会把早就死掉的宿主判成"还活着"，
# 于是新宿主永远拒绝接管，报「检测到另一个 DSH 会话正在控制 WPS」，而其实一个宿主都没有。
# 实测：租约里 hostPid=41592，系统里 **0 个** wps-com-host.ps1 进程 —— 那个号已被无关进程占用；
# 用户除了手工删 %USERPROFILE%\.wps-office-mcp\com-host.json 之外没有别的办法。
# 租约里本来就记着宿主的启动时刻（startedUtc），拿它核对身份即可分辨复用。
function Test-OwnerAlive($processId, $startedUtc) {
    if (-not (Test-ProcessAlive $processId)) { return $false }
    if ([string]::IsNullOrEmpty($startedUtc)) { return $true }   # 旧租约没这一项：退回只看 PID
    try {
        $proc = Get-Process -Id ([int]$processId) -ErrorAction Stop
        $actual = $proc.StartTime.ToUniversalTime()
        $recorded = ([datetime]::Parse([string]$startedUtc)).ToUniversalTime()
        # 允许 2 秒误差：记录的时刻是宿主自己取的，与进程表里的 StartTime 可能有亚秒级差异。
        return ([math]::Abs(($actual - $recorded).TotalSeconds) -lt 2)
    } catch { return $false }
}

function Get-StaleOwnerReason {
    # $null means "the owner looks alive" - the newcomer must refuse.
    $state = Read-HostState
    if ($null -eq $state) { return 'the lease file is missing or unreadable' }
    if (-not (Test-OwnerAlive $state.hostPid $state.startedUtc)) { return 'the owning host process is gone (or its pid was reused by an unrelated process)' }
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
    # FIXES 96（W5-1）：必须用**身份核对**而不是"PID 还在"来决定要不要杀 —— 否则一个被复用的 PID
    # 会让新宿主去 Stop-Process 一个毫不相干的进程。
    if ($null -ne $state -and (Test-OwnerAlive $state.hostPid $state.startedUtc)) {
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
        '请先关闭那个会话或结束它的 WPS 操作，然后重试。' + $holder +
        '（FIXES 96：如果确认没有别的会话在用 WPS，删掉 ' + $script:StateFile + ' 再重试即可。）'
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

# FIXES 102（A4 定性）：**ready 必须等回收做完再发**。
# 以前这里是「先发 ready，再回收」——而回收里全是针对遗留 WPS 的 COM 调用
# （GetActiveObject / Test-WpsAppUsable / Test-WpsAppHasUnsavedWork / Quit），这条路上**没有任何超时**。
# 后果（实测复现）：
#   ① 客户端看到 ready 就以为宿主可用（它的 30 秒启动兜底**在这一刻就满足了**）；
#   ② 客户端紧接着发的第一个请求排在管道里 —— 宿主此时还没进主循环，根本不读 stdin；
#   ③ 如果那个遗留实例正卡在模态对话框/繁忙状态（本项目一直在防的失效模式），COM 会一直等，
#      于是**第一个请求永远得不到响应**。param-contract 的 2 小时挂起就是这么来的（它有 invoke 无超时）。
# 顺序证据：在"响应到达的那一刻"检查回收记录文件 —— 已经被 Clear-WpsOwnedApps 删掉了，
# 说明响应确实排在回收之后（详见 docs/FIXES.md 102）。
# 把 ready 挪到回收之后：客户端等待的是"宿主真正可用"，卡住时会由它的启动超时报出来，而不是无限等。
#
# FIXES 66: a predecessor that was force-killed (FIXES 65) left WPS instances behind. Now that this
# host holds the lease, reclaim them - unless the recorded owner is still alive, or a document holds
# unsaved work. One file probe when there is nothing to reclaim.
try { $null = Invoke-WpsOrphanReclaim } catch { }

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
        # Ownership is settled; the exit net at the end of the file must not quit them again.
        $script:WpsAppOwned = @{}
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

# Exit net (FIXES 65): the loop above can also end on stdin EOF - a client that closes the pipe or
# exits without sending __shutdown. Cleanup used to live only inside the __shutdown branch, so those
# exits abandoned every WPS instance this host had started. Safe on any path: Close-WpsAppsStartedByUs
# skips instances it did not start and instances that hold unsaved work.
# Note: a force-killed client takes this whole process tree down on Windows before any of this runs,
# so that path is handled by scripts/run-tests.ps1 reaping the headless orphans instead.
$script:BusySinceUtc = ''
try { $null = Close-WpsAppsStartedByUs } catch { }
Update-HostState 'stopped'

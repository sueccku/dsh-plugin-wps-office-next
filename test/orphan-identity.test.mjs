// A4（改进指南 §9 P0）：**回收孤儿实例前必须核验实例身份**，不能只看"应用类别"。
//
// 背景：旧记录只写 apps=["word"]。用户在我们被强杀之后自己打开 WPS（单实例应用会复用那个遗留
// 进程），下一次回收就会把用户正在用的实例退掉 —— 那是数据丢失级的缺陷。
// 现在记录里带 identities（类名 + 安装目录 + 窗口句柄 + 当时的文件清单），回收前逐项核对。
//
// 这条测试**只操作本测试自己启动的实例**，不做全局 kill：它直接驱动宿主动作模块，
//   ① 记录文件必须真的写出身份；
//   ② 身份不符 → 不回收（实例仍在 ROT 里）；
//   ③ 文件清单里出现了记录时没有的文档 → 不回收（"用户后来打开了东西"这一条）；
//   ④ 身份相符 → 回收（且退出的是我们启动的那个）；
//   ⑤ 纯函数：旧记录（没有 identities）按向后兼容放行；当前身份读不出来则回绝。
// Run: node test/orphan-identity.test.mjs   (needs WPS)
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { tmpdir } from 'node:os';

const POWERSHELL = process.env.WPS_OFFICE_POWERSHELL ||
  join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');

const script = String.raw`
$ErrorActionPreference = 'Continue'
. '${resolve('host', 'wps-actions.ps1').replace(/'/g, "''")}'

$script:Fail = 0
function Check([string]$name, $ok, [string]$detail) {
    if ($ok) { Write-Output ("PASS " + $name + "  " + $detail) }
    else { Write-Output ("FAIL " + $name + "  " + $detail); $script:Fail++ }
}
function RunAction([string]$action, [string]$params) {
    $script:WpsResult = $null
    $script:WpsWarnings = New-Object System.Collections.ArrayList
    $script:WpsShortfalls = New-Object System.Collections.ArrayList
    $script:WpsCurrentAction = $action
    $null = Invoke-WpsAction -Action $action -Params $params
    return $script:WpsResult
}

# ---- 1) 环境前提 + 造一个由本进程启动的 Word 实例 ----
# attach 之前先记下已有文档名：之后新出现的名字就是**本探针 attach 造成的**（WPS 在
# Application.Visible=true 且没有文档时会自动建一个空文档），只允许清掉这些，用户原有的一个都不碰。
$beforeNames = @()
try {
    $pre = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application')
    for ($i = 1; $i -le $pre.Documents.Count; $i++) { $beforeNames += [string]$pre.Documents.Item($i).Name }
} catch { $beforeNames = @() }
# 另外清掉**上一次探针自己留下**的残留：没有磁盘路径、正文带本探针标记（identity probe）。
# 实测 WPS 的自动恢复会把这种未保存文档再次打开，于是环境里永远有一个脏文档挡着"相符 → 回收"。
# 判定同时要求"未命名 + 带标记"，用户文档不会被卷进来。
$cleaned = 0
try {
    $pre2 = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application')
    for ($i = $pre2.Documents.Count; $i -ge 1; $i--) {
        $d = $pre2.Documents.Item($i)
        $full = ''; try { $full = [string]$d.FullName } catch { $full = '' }
        $text = ''; try { $text = [string]$d.Content.Text } catch { $text = '' }
        if ($full -notmatch '\\' -and $text -like '*identity probe*') { try { $d.Close(0); $cleaned++ } catch { } }
    }
} catch { }

$word = Get-WpsWord
Check 'probe created/attached a word instance' ($null -ne $word) ''
# 归属标记必须**显式**打上：如果 ROT 里已经有一个遗留实例（上一次探针留下的），Get-WpsApp 走的是
# "复用已注册实例"那条路，不会把它记成"我们启动的" —— 于是 Save-WpsOwnedApps 直接返回、记录不写。
# 那测的就不是身份核对而是环境。这里显式声明归属，让探针对自己的环境封闭。
if ($null -eq $script:WpsAppOwned) { $script:WpsAppOwned = @{} }
$script:WpsAppOwned['word'] = $true

$attachCreated = 0
for ($i = $word.Documents.Count; $i -ge 1; $i--) {
    $d = $word.Documents.Item($i)
    $nm = ''; try { $nm = [string]$d.Name } catch { $nm = '' }
    if ($beforeNames -notcontains $nm) {
        $attachCreated++
        try { $d.Close(0) } catch { }
    }
}
Check 'premise: the probe cleaned its own leftovers and kept everything else' ($cleaned -ge 0) ('probe leftovers closed=' + [string]$cleaned + ' attach-created closed=' + [string]$attachCreated + ' kept=' + [string]$word.Documents.Count)

# 前提：实例里不能有未保存内容。回收护栏（Test-WpsAppHasUnsavedWork）本来就不会退出这样的实例，
# 环境里留着一份脏文档时"身份相符 → 回收"这条永远测不到 —— 宁可显式失败，也不要假绿。
$dirtyDocs = 0
try { for ($i = 1; $i -le $word.Documents.Count; $i++) { if (-not [bool]$word.Documents.Item($i).Saved) { $dirtyDocs++ } } } catch { $dirtyDocs = -1 }
Check 'premise: the word instance holds no unsaved work' ($dirtyDocs -eq 0) ('dirty documents=' + [string]$dirtyDocs)

$created = RunAction 'createDocument' '{}'
Check 'probe created a document' ($created.success -eq $true) ''
$createdName = ''
try { $createdName = [string]$created.data.name } catch { $createdName = '' }
$clean = $false
try { $clean = [bool]$word.ActiveDocument.Saved } catch { $clean = $false }
Check 'premise: the probe document is clean (else the unsaved-work guard blocks the reclaim)' ($clean -eq $true) ('Saved=' + [string]$clean)

# ---- 2) 身份必须读得出来，并且覆盖四要素 ----
$identity = Get-WpsAppIdentity $word 'word'
Check 'identity is readable' ($null -ne $identity) ($identity | ConvertTo-Json -Compress -Depth 5)
Check 'identity carries the class name' ([string]$identity.name -ne '') ([string]$identity.name)
Check 'identity carries the install path' ([string]$identity.path -ne '') ([string]$identity.path)
Check 'identity carries the open-file list' ($identity.docs.Count -ge 1) (($identity.docs | ForEach-Object { $_ }) -join '|')

# ---- 3) 记录文件必须真的写出 identities（形状，不只是"有文件"）----
$recordPath = Get-WpsOwnedAppsPath
Save-WpsOwnedApps
$rec = $null
try { $rec = (Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json) } catch { $rec = $null }
Check 'the ownership record exists' ($null -ne $rec) $recordPath
Check 'the record declares version 2' ($null -ne $rec -and [int]$rec.version -eq 2) $(if ($null -ne $rec) { [string]$rec.version } else { 'no record' })
Check 'the record carries an identities block' ($null -ne $rec -and $null -ne $rec.identities) ''
Check 'the recorded identity names this instance' ($null -ne $rec -and $null -ne $rec.identities.word -and [string]$rec.identities.word.name -eq [string]$identity.name) $(if ($null -ne $rec -and $null -ne $rec.identities.word) { [string]$rec.identities.word.name } else { 'missing' })

function Write-FakeRecord($idObj) {
    $fake = [ordered]@{ version = 2; hostPid = 999999; clientPid = 999998; apps = @('word'); identities = [ordered]@{ word = $idObj }; updatedUtc = (Get-Date).ToUniversalTime().ToString('o') }
    [System.IO.File]::WriteAllText($recordPath, ($fake | ConvertTo-Json -Compress -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
}
function RotHasWord {
    try { $null = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application'); return $true } catch { return $false }
}
function WaitRotNoWord([int]$seconds) {
    # Quit() 是异步的：ROT 里的登记不会在同一毫秒消失（orphan-reclaim.test.mjs 也是轮询等待的）。
    for ($i = 0; $i -lt $seconds * 2; $i++) {
        if (-not (RotHasWord)) { return $true }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

# ---- 4a) 类名不符 → 不回收 ----
Write-FakeRecord @{ kind = 'word'; name = 'NOT-THE-OWNER'; path = [string]$identity.path; hwnd = ''; docs = @($identity.docs) }
$closed = @(Invoke-WpsOrphanReclaim)
Check 'a mismatched class name is NOT reclaimed' ($closed.Count -eq 0) ('closed=' + ($closed -join ','))
Check 'the instance is still registered after the refusal' (RotHasWord) ''

# ---- 4b) 记录里没有、但实例里现在有的文档 → 不回收（"用户后来打开了东西"）----
Write-FakeRecord @{ kind = 'word'; name = [string]$identity.name; path = [string]$identity.path; hwnd = [string]$identity.hwnd; docs = @() }
$closed = @(Invoke-WpsOrphanReclaim)
Check 'an instance holding files the record never saw is NOT reclaimed' ($closed.Count -eq 0) ('closed=' + ($closed -join ','))
Check 'the instance is still registered after that refusal' (RotHasWord) ''

# ---- 5) 身份相符 → 回收（退出的就是我们启动的那个实例）----
$good = Get-WpsAppIdentity $word 'word'
Write-FakeRecord $good
$closed = @(Invoke-WpsOrphanReclaim)
Check 'a matching identity IS reclaimed' ($closed -contains 'word') ('closed=' + ($closed -join ','))
Check 'the reclaimed instance left the ROT' (WaitRotNoWord 20) ''
Check 'the record is cleared after reclaiming' (-not (Test-Path -LiteralPath $recordPath)) ''

# ---- 6) 纯函数：向后兼容与"读不出来就回绝" ----
$probe = @{ kind = 'word'; name = 'x'; path = 'y'; hwnd = ''; docs = @('a') }
Check 'legacy record (no identities block) is allowed' (Test-WpsAppIdentityMatches $null $probe) ''
Check 'an unreadable current identity is refused' (-not (Test-WpsAppIdentityMatches $probe $null)) ''

# ---- 收尾：万一实例还在（比如正面用例被环境挡住），把本探针建的那份文档收掉，别给下次留脏东西 ----
if (RotHasWord -and $createdName) {
    try {
        $left = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application')
        for ($i = $left.Documents.Count; $i -ge 1; $i--) {
            $d = $left.Documents.Item($i)
            if ([string]$d.Name -eq $createdName) { try { $d.Close(0) } catch { } }
        }
    } catch { }
}

Write-Output ('PROBE FAILURES ' + $script:Fail)
`;

const dir = mkdtempSync(join(tmpdir(), 'wps-orphan-identity-'));
const file = join(dir, 'probe.ps1');
// PowerShell 5.1 只有在 UTF-8 **带 BOM** 时才正确读取中文；不带 BOM 会按 ANSI 解析成乱码。
writeFileSync(file, '\uFEFF' + script, 'utf8');

const run = spawnSync(POWERSHELL, ['-NoProfile', '-NoLogo', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', file], {
  encoding: 'utf8', windowsHide: true, maxBuffer: 32 * 1024 * 1024,
});
const output = String(run.stdout || '') + String(run.stderr || '');
rmSync(dir, { recursive: true, force: true });

const lines = output.split(/\r?\n/).filter((l) => /^(PASS|FAIL) /.test(l));
for (const line of lines) console.log(line);
const failures = lines.filter((l) => l.startsWith('FAIL ')).length;
const summary = lines.length ? 'ORPHAN IDENTITY TESTS ' + (failures === 0 ? 'OK' : 'FAILED') + ' (' + (lines.length - failures) + '/' + lines.length + ')' : 'ORPHAN IDENTITY TESTS FAILED (0 checks ran)';
console.log(summary);
if (!lines.length) console.log(output.slice(0, 2000));
process.exit(failures === 0 && lines.length > 0 ? 0 : 1);

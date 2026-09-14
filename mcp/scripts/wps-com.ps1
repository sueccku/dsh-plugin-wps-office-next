# Input: Action 名称与 JSON 参数
# Output: WPS COM 调用结果 JSON
# Pos: Windows COM 桥接脚本。一旦我被修改，请更新我的头部注释（Updated: 2026-05-26 15:30:00 CST），以及所属文件夹的md。
# WPS COM Bridge - PowerShell script for WPS COM operations
# Full implementation for Excel, Word, PPT, and common conversions
# Usage: powershell -File wps-com.ps1 -Action <action> -Params <json>

param(
    [string]$Action,
    [string]$Params = "{}"
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ==================== COM Object Getters ====================

function ConvertTo-WordMarginPoints($value) {
    # WPS Word's PageSetup margins accept a whole number of points only: a fractional value raises
    # 'Specified cast is not valid' and an out-of-range one a bare E_FAIL. Returns $null when the
    # value cannot be used, so the caller can say why.
    $points = [double]$value
    if ($points -lt 0 -or $points -gt 1584) { return $null }
    return [int][Math]::Round($points)
}
function Test-WpsSlideIndex($pres, $value) {
    # WPS reports Slides.Count as 0 for decks created through COM, so the upper bound cannot be
    # trusted: a strict check rejected every valid slide in such a deck. The lower bound is always
    # enforced, and the upper bound only when the count is actually known.
    if ($null -eq $value) { return $true }
    $index = [int]$value
    if ($index -lt 1) { return $false }
    $total = 0
    try { $total = [int]$pres.Slides.Count } catch { $total = 0 }
    if ($total -le 0) { return $true }
    return ($index -le $total)
}
function Invoke-ComMethodReflect($target, [string]$method, [object[]]$arguments) {
    # PowerShell's COM member resolution can miss a member the process has not materialised yet (the
    # same lazy-member problem seen with Range.Replace), which surfaces as "does not contain a method
    # named 'Align'" even though the member exists. Dispatch through IDispatch via reflection.
    if ($null -eq $target) { throw ('Invoke-ComMethodReflect: null target for ' + $method) }
    if ($null -eq $arguments) { $arguments = @() }
    return $target.GetType().InvokeMember($method, [System.Reflection.BindingFlags]::InvokeMethod, $null, $target, $arguments)
}

function Test-WpsAppUsable($app, [string]$kind) {
    # The acquisition chain can hand back an object that is not usable: [Activator]::CreateInstance
    # on a Coclass produces an instance whose collection is null, and a stale ROT entry behaves the
    # same way. Every later call then dies with 'cannot call a method on a null-valued expression',
    # which says nothing about the real cause, so every candidate is checked before it is accepted.
    if ($null -eq $app) { return $false }
    try {
        if ($kind -eq 'excel') { return ($null -ne $app.Workbooks) }
        if ($kind -eq 'ppt') { return ($null -ne $app.Presentations) }
        return ($null -ne $app.Documents)
    } catch { return $false }
}

function Reset-WpsApp([string]$kind) {
    # Drop the cached instance so the next Get-WpsApp acquires a fresh one.
    $cacheName = 'WpsAppCache_' + $kind
    Remove-Variable -Name $cacheName -Scope Script -ErrorAction SilentlyContinue
}

function Get-WpsApp([string]$kind) {
    # One validated instance per process: re-acquiring can return a broken object even after a good
    # acquisition, and a resident host would pay for that on every action.
    $cacheName = 'WpsAppCache_' + $kind
    $cached = Get-Variable -Name $cacheName -Scope Script -ErrorAction SilentlyContinue
    if ($null -ne $cached -and (Test-WpsAppUsable $cached.Value $kind)) { return $cached.Value }
    $progId = @{ excel = 'Ket.Application'; ppt = 'Kwpp.Application'; word = 'Kwps.Application' }[$kind]
    if ($null -eq $progId) { return $null }
    $candidates = @()
    try { $candidates += [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId) } catch { }
    try { $candidates += (New-Object -ComObject $progId) } catch { }
    # NOTE: [Activator]::CreateInstance on the Coclass is deliberately not used. It returns an
    # instance whose Presentations/Workbooks collection looks fine but whose Add() dereferences an
    # uninitialised document manager, so it fails with 'cannot call a method on a null-valued
    # expression' only later, far from the cause.
    foreach ($candidate in $candidates) {
        if (-not (Test-WpsAppUsable $candidate $kind)) { continue }
        if ($kind -ne 'excel') { try { $candidate.Visible = $true } catch { } }
        Set-Variable -Name $cacheName -Scope Script -Value $candidate
        return $candidate
    }
    return $null
}
$script:WpsWarnings = New-Object System.Collections.ArrayList

function Add-WpsWarning([string]$message) {
    # Collect a best-effort failure instead of swallowing it. Anything collected is attached to the
    # action's result, so a partially applied change is visible to the caller instead of looking clean.
    if ($null -eq $script:WpsWarnings) { $script:WpsWarnings = New-Object System.Collections.ArrayList }
    $null = $script:WpsWarnings.Add($message)
}

function Clear-WpsWarnings() {
    $script:WpsWarnings = New-Object System.Collections.ArrayList
}

function Get-WpsExcel { return Get-WpsApp 'excel' }

function Get-WpsWord { return Get-WpsApp 'word' }

function Get-WpsPpt { return Get-WpsApp 'ppt' }

# ==================== 模态弹窗围堵 (S1) ====================
# WPS 打开加密文档时会弹「文档已加密」模态框；模态框在 STA 里是阻塞的，整个宿主线程就此停住，
# 请求永不返回、客户端超时、后续每次调用一起挂。本机实测（WPS 12.1 x64，见 docs/FIXES.md 52）：
#   Documents.Open(path)                                    -> 卡死（弹框）
#   Documents.Open(path, ..., PasswordDocument:="")         -> 仍卡死（空串=「没给密码」）
#   Documents.Open(path, ..., PasswordDocument:="<哨兵值>")  -> 4 秒内返回错误，不弹框
#   Workbooks.Open(path)                                    -> 卡死（弹框）
#   Workbooks.Open(path, 0, ..., Password:="<哨兵值>")       -> 1 秒内返回错误，不弹框
# 所以真正把弹框挡在门外的是**非空哨兵密码**，`DisplayAlerts` 管不到它（实测 Word 的
# DisplayAlerts 为 0 时密码框照样弹）。哨兵值不可能等于任何真实密码，这正是重点：加密文件
# 改为**快速报错**，而不是无限等待一个没人知道的密码。
$script:WpsNoPassword = 'dsh-wps-office-next/no-password-supplied'

function Set-WpsAlertsSuppressed($app, [string]$kind) {
    # 关掉弹窗并返回原值供调用方还原；读不到或写不了就返回 $null，调用方据此跳过还原，
    # 绝不因为「还原失败」把动作本身搞错。
    if ($null -eq $app) { return $null }
    $off = if ($kind -eq 'ppt') { 1 } else { $false }
    $prev = $null
    try { $prev = $app.DisplayAlerts } catch { return $null }
    try { $app.DisplayAlerts = $off } catch { return $null }
    return $prev
}

function Restore-WpsAlerts($app, $prev) {
    # try/finally 里调用；还原失败只降级为静默，因为此时动作结果已经拿到。
    if ($null -eq $app -or $null -eq $prev) { return }
    try { $app.DisplayAlerts = $prev } catch { }
}

function Open-WordDocument($word, [string]$path) {
    # WPS 的 IDispatch 不支持命名参数（实测 E_INVALIDARG），只能按位置传，顺序不能改：
    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
    # Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair,
    # DocumentDirection, NoEncodingDialog
    return $word.Documents.Open($path, $false, $false, $true, $script:WpsNoPassword, "", $false, "", "", 0, 0, $false, $false, 0, $true)
}

function Open-ExcelWorkbook($excel, [string]$path, $updateLinks, [bool]$readOnly) {
    # UpdateLinks 默认 0（不更新外部链接）而不是留空：留空等同于「按设置来」，而设置里
    # AskToUpdateLinks 为真时照样弹框。调用方显式给了值就尊重它。
    # FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended
    $links = 0
    if ($null -ne $updateLinks) { try { $links = [int]$updateLinks } catch { $links = 0 } }
    return $excel.Workbooks.Open($path, $links, $readOnly, $null, $script:WpsNoPassword, "", $true)
}

function Open-PptPresentation($ppt, [string]$path) {
    # Presentations.Open 没有密码参数，加密演示文稿仍会弹框（已知残余限制，见 README）；
    # 其余弹窗由调用方的 DisplayAlerts 抑制。FileName, ReadOnly, Untitled, WithWindow
    return $ppt.Presentations.Open($path, $false, $false, $true)
}

function Format-WpsOpenError([string]$path, [string]$kind, [string]$message) {
    # 加密文件的原始错误在 WPS 里并不统一：Excel 给裸 HRESULT 0xFFF40006，Word 只给一句
    # 「文档打开失败。」这样的 stub。前者能精确判定，后者只能把「加密」当成最可能的原因说明白。
    $text = [string]$message
    if ($text -match 'FFF40006' -or $text -match '密码' -or $text -match '加密' -or $text -match 'password') {
        return ("无法打开「" + $path + "」：文件已加密，需要密码。插件不会弹出密码对话框（那会把整个会话卡死）；" +
            "请先用 WPS 手工打开它，另存为一份不加密的副本，再对副本操作。")
    }
    if ($text -match 'not exist' -or $text -match '找不到' -or $text -match '不存在' -or $text -match '0x80070002') {
        return ("无法打开「" + $path + "」：文件不存在或路径不可读。原始错误：" + $text)
    }
    return ("无法打开「" + $path + "」（" + $kind + "）：" + $text +
        " 若该文件设置了打开密码，插件不会弹出密码框（弹框会把整个会话卡死）；" +
        "请先用 WPS 手工打开它、另存为一份不加密的副本，再对副本操作。")
}

# 解析目标演示文稿：优先按 presentationName 精确定位（避免多文稿打开时 ActivePresentation 漂移），
# 未提供则回退到当前活动文稿；提供了但找不到则返回 $null（让调用方报明确错误，而不是误改其它文稿）。
function Get-TargetPres($ppt, $p) {
    if ($null -eq $ppt) { return $null }
    if ($p.presentationName) {
        try { return $ppt.Presentations.Item([string]$p.presentationName) } catch { return $null }
    }
    return $ppt.ActivePresentation
}

# 在一个 TextFrame 内查找替换：先精确匹配(保留原换行)；多行 find 未命中时，把 \r\n/\r/\v 与 \n 规范化后再匹配并写回。返回是否发生替换。
function Replace-FrameText($frame, [string]$findText, [string]$replaceText) {
    try {
        if ($null -eq $frame -or -not $frame.HasText) { return $false }
        $text = $frame.TextRange.Text
        if ($null -eq $text) { return $false }
        if ($text.Contains($findText)) {
            $frame.TextRange.Text = $text.Replace($findText, $replaceText)
            return $true
        }
        if ($findText.Contains("`n")) {
            $nf = ((($findText -replace "`r`n", "`n") -replace "`r", "`n") -replace "`v", "`n")
            $nt = ((($text -replace "`r`n", "`n") -replace "`r", "`n") -replace "`v", "`n")
            if ($nt.Contains($nf)) {
                $frame.TextRange.Text = $nt.Replace($nf, $replaceText)
                return $true
            }
            # 兜底1：去掉所有换行后，若 find 恰为整块文本(PDF抽取的"幻影换行")，整体替换该形状
            $sf = ($findText -replace "[`r`n`v]", "")
            $st = ($text -replace "[`r`n`v]", "")
            if ($sf.Length -gt 0 -and $st -eq $sf) {
                $frame.TextRange.Text = $replaceText
                return $true
            }
            # 兜底2：去换行做子串匹配(find 是大形状里的一段且含幻影换行)，命中后用 stripped→原文 索引映射在原位替换
            if ($sf.Length -gt 0) {
                $origChars = $text.ToCharArray()
                $sb = New-Object System.Text.StringBuilder
                $map = New-Object System.Collections.Generic.List[int]
                for ($k = 0; $k -lt $origChars.Length; $k++) {
                    $ch = $origChars[$k]
                    if ($ch -ne [char]13 -and $ch -ne [char]10 -and $ch -ne [char]11) {
                        [void]$sb.Append($ch); $map.Add($k)
                    }
                }
                $idx = $sb.ToString().IndexOf($sf)
                if ($idx -ge 0) {
                    $startOrig = $map[$idx]
                    $endOrig = $map[$idx + $sf.Length - 1]
                    $before = $text.Substring(0, $startOrig)
                    $after = $text.Substring($endOrig + 1)
                    $frame.TextRange.Text = $before + $replaceText + $after
                    return $true
                }
            }
        }
        return $false
    } catch { return $false }
}

# 递归处理一个形状(及其组合/表格内的所有子形状)的文本替换，返回命中次数。
# 修复：replacePptText 原来只扫顶层 Shapes，组合(Group)内文字(如封面/立项页的旧项目名)替换不到。
function Replace-InShapeTree($shape, [string]$findText, [string]$replaceText) {
    $cnt = 0
    try {
        $isGroup = $false
        try { if ($shape.Type -eq 6) { $isGroup = $true } } catch { Add-WpsWarning $_.Exception.Message }
        if ($isGroup) {
            for ($gi = 1; $gi -le $shape.GroupItems.Count; $gi++) {
                $cnt += Replace-InShapeTree $shape.GroupItems.Item($gi) $findText $replaceText
            }
        } else {
            try { if ($shape.HasTextFrame) { if (Replace-FrameText $shape.TextFrame $findText $replaceText) { $cnt++ } } } catch { Add-WpsWarning $_.Exception.Message }
            try {
                if ($shape.HasTable) {
                    $tbl = $shape.Table
                    for ($rr = 1; $rr -le $tbl.Rows.Count; $rr++) {
                        for ($cc = 1; $cc -le $tbl.Columns.Count; $cc++) {
                            try { if (Replace-FrameText $tbl.Cell($rr, $cc).Shape.TextFrame $findText $replaceText) { $cnt++ } } catch { Add-WpsWarning $_.Exception.Message }
                        }
                    }
                }
            } catch { Add-WpsWarning $_.Exception.Message }
        }
    } catch { Add-WpsWarning $_.Exception.Message }
    return $cnt
}

# ==================== COM boundary helpers ====================
# PowerShell caches the COM binder for a member after its first use, so one call site fails as
# soon as the value type changes (Int32 vs String). Every COM member written by the action
# layer must go through these helpers: they re-bind on each call and normalise the value.
function ConvertTo-ComValue($value) {
    if ($null -eq $value) { return $null }
    if ($value -is [bool]) { return [bool]$value }
    if ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal]) { return [double]$value }
    return [string]$value
}
function Find-ComProperty($target, [string]$member) {
    if ($null -eq $target) { throw ('Find-ComProperty: null target for member ' + $member) }
    $prop = $target.PSObject.Properties[$member]
    if ($null -ne $prop) { return $prop }
    # A COM adapter's member collection is populated lazily, so fall back to enumeration.
    return ($target.PSObject.Properties | Where-Object { $_.Name -eq $member } | Select-Object -First 1)
}
function Get-PropOrNull($p, [string]$name) {
    $prop = $p.PSObject.Properties | Where-Object { $_.Name -eq $name } | Select-Object -First 1
    if ($null -eq $prop) { return $null }
    return $prop.Value
}
function Resolve-Worksheet($excel, $wb, $p, [switch]$RequireName) {
    # Upstream tools spell the sheet parameter differently per tool (sheet/name/oldName).
    foreach ($key in @('sheet','name','oldName')) {
        $value = Get-PropOrNull $p $key
        if ($null -ne $value -and "$value" -ne "") { return $wb.Sheets.Item($value) }
    }
    if ($RequireName) { return $null }
    return $excel.ActiveSheet
}

function Open-ExportedFile($path, $flag) {
    # Returns "opened", "failed: <reason>" or $null when nothing was asked for.
    if (-not $flag) { return $null }
    if ($null -eq $path -or "$path" -eq "") { return "failed: no output path" }
    try {
        Start-Process -FilePath ([string]$path)
        return "opened"
    } catch {
        return ("failed: " + $_.Exception.Message)
    }
}

function Resolve-InputFilePath($value) {
    # WPS resolves picture paths against its own working directory, so a relative path that looks
    # right from the caller's shell fails with a bare E_FAIL. Resolve it here and report a missing
    # file as such.
    if ($null -eq $value -or "$value" -eq "") { return $null }
    $candidate = [string]$value
    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        $candidate = [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $candidate))
    }
    if (-not (Test-Path -LiteralPath $candidate)) { return $null }
    return $candidate
}

function Get-RowRefList($p) {
    # The tool layer may pass rows, row (+count), or startRow/endRow. Normalise to row refs.
    $rows = @()
    if ($null -ne $p.rows) { $rows = @($p.rows) }
    elseif ($null -ne $p.startRow -and $null -ne $p.endRow) { $rows = @("$($p.startRow):$($p.endRow)") }
    elseif ($null -ne $p.row) {
        if ($null -ne $p.count -and [int]$p.count -gt 1) { $rows = @("$($p.row):$([int]$p.row + [int]$p.count - 1)") }
        else { $rows = @([string]$p.row) }
    }
    return , $rows
}

function Get-ColumnRefList($p) {
    $cols = @()
    if ($null -ne $p.columns) { $cols = @($p.columns) }
    elseif ($null -ne $p.startColumn -and $null -ne $p.endColumn) { $cols = @("$($p.startColumn):$($p.endColumn)") }
    elseif ($null -ne $p.column) {
        if ($null -ne $p.count -and [int]$p.count -gt 1) {
            $first = $p.column
            if ($first -is [int]) { $first = Convert-ColumnNumberToLetter([int]$first) }
            $last = $p.column
            if ($last -is [int]) { $last = Convert-ColumnNumberToLetter([int]$last + [int]$p.count - 1) }
            $cols = @([string]$first + ":" + [string]$last)
        } else { $cols = @($p.column) }
    }
    return , $cols
}

function Get-WorksheetByParam($excel, $p) {
    # Most Excel actions are documented to accept a target worksheet but silently operated on
    # whichever sheet happened to be active. Only $p.sheet is read here: in several actions "name"
    # means something else entirely (a range name, a chart name), so aliasing it to a sheet would
    # send those actions to the wrong place.
    $name = Get-PropOrNull $p 'sheet'
    if ($null -ne $name -and "$name" -ne "") {
        $wb = $excel.ActiveWorkbook
        if ($null -ne $wb) { return $wb.Sheets.Item($name) }
    }
    return $excel.ActiveSheet
}

function Get-PptBackgroundSpec($p) {
    # The tool sends one background object; flat keys are accepted too. Returns a spec with either
    # rgb1/rgb2/imagePath filled in or a human-readable error, so both callers stay short.
    $bg = Get-PropOrNull $p 'background'
    $kind = if ($null -ne $bg -and $bg.type) { [string]$bg.type } elseif ($null -ne $p.type) { [string]$p.type } else { $null }
    $color = if ($null -ne $bg -and $bg.color) { [string]$bg.color } elseif ($null -ne $p.color) { [string]$p.color } else { $null }
    $colors = @()
    if ($null -ne $bg -and $null -ne $bg.colors) { $colors = @($bg.colors) } elseif ($null -ne $p.colors) { $colors = @($p.colors) }
    $image = if ($null -ne $bg -and $bg.imagePath) { [string]$bg.imagePath } elseif ($null -ne $p.imagePath) { [string]$p.imagePath } else { $null }
    # No explicit type: infer it from what was supplied, so the flat spelling
    # (color / colors / imagePath) works as well as the background object.
    if ($null -eq $kind -or "$kind" -eq "") {
        if ($null -ne $image -and "$image" -ne "") { $kind = "image" }
        elseif ($colors.Count -ge 2) { $kind = "gradient" }
        elseif ($null -ne $color -and "$color" -ne "") { $kind = "solid" }
        else { $kind = "solid" }
    }
    $spec = @{ kind = $kind; color = $color; colors = $colors; imagePath = $image; error = $null }
    if ($kind -eq "solid") {
        if ($null -eq $color -or $color -eq "") { $spec.error = "a solid background needs a color" }
        elseif ($null -eq (Convert-HexColorToRgbInt $color)) { $spec.error = ("unrecognized color '" + $color + "'") }
        else { $spec.rgb1 = Convert-HexColorToRgbInt $color }
    } elseif ($kind -eq "gradient") {
        if ($colors.Count -lt 2) { $spec.error = "a gradient background needs two colors in colors" }
        else {
            $spec.rgb1 = Convert-HexColorToRgbInt ([string]$colors[0])
            $spec.rgb2 = Convert-HexColorToRgbInt ([string]$colors[1])
        }
    } elseif ($kind -eq "image") {
        if ($null -eq $image -or $image -eq "") { $spec.error = "an image background needs imagePath" }
        else {
            $resolved = Resolve-InputFilePath $image
            if ($null -eq $resolved) { $spec.error = ("background image not found: " + $image) } else { $spec.imagePath = $resolved }
        }
    } else { $spec.error = ("unknown background type '" + $kind + "'; use solid/gradient/image") }
    return $spec
}

function Set-PptBackgroundFill($fill, $spec) {
    if ($spec.kind -eq "solid") { $fill.Solid(); $fill.ForeColor.RGB = $spec.rgb1; return }
    if ($spec.kind -eq "gradient") { $fill.TwoColorGradient(1, 1); $fill.ForeColor.RGB = $spec.rgb1; $fill.BackColor.RGB = $spec.rgb2; return }
    $fill.UserPicture($spec.imagePath)
}

function Get-PptEntryEffect($value) {
    # PpEntryEffect constants. The tools name the transition; COM needs the number, and assigning a
    # string raised a bare index-out-of-range error from WPS instead of saying what was wrong.
    $map = @{
        none = 0; cut = 257; fade = 1793; dissolve = 1537
        push = 3840; wipe = 2561; split = 1281; reveal = 2049; cover = 3854; curtains = 3586
    }
    if ($null -eq $value) { return 1793 }
    $key = ([string]$value).ToLower()
    if ($map.ContainsKey($key)) { return $map[$key] }
    $numeric = 0
    if ([int]::TryParse($key, [ref]$numeric)) { return $numeric }
    return $null
}

function Get-PptAnimEffect($value) {
    # MsoAnimEffect constants, plus the Exit variants the tools also offer.
    $map = @{
        appear = 1; flyin = 2; blinds = 3; box = 4; checkerboard = 5; circle = 6; crawl = 7; diamond = 8
        dissolve = 9; fadein = 10; fade = 10; flashonce = 11; peeko = 12; plus = 13; randombars = 14
        spiral = 15; split = 16; stretch = 17; stripc = 18; swirl = 19; wheel = 20; wipein = 22; wipe = 22
        zoomin = 23; zoom = 23; bouncein = 26; bounce = 26; spinin = 61; spin = 61; flyout = 2; fadeout = 10
    }
    if ($null -eq $value) { return 10 }
    $key = ([string]$value).ToLower()
    if ($map.ContainsKey($key)) { return $map[$key] }
    $numeric = 0
    if ([int]::TryParse($key, [ref]$numeric)) { return $numeric }
    return $null
}

function Resolve-ShapeByKind($slide, $p, [string]$kind, $ordinal) {
    # name/shapeName/shapeIndex address any shape directly. The kind-specific ordinals count only
    # pictures or only text boxes, which is what the tool schemas promise (imageIndex 1 means the
    # first picture, not shape 1 - aliasing it to shapeIndex deleted the wrong shape).
    if ($null -ne $p.name -and "$($p.name)" -ne "") { return $p.name }
    if ($null -ne $p.shapeName -and "$($p.shapeName)" -ne "") { return $p.shapeName }
    if ($null -ne $p.shapeIndex) { return $p.shapeIndex }
    if ($null -eq $ordinal) { return $null }
    $seen = 0
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        $isWanted = $false
        try {
            if ($kind -eq "picture") { $isWanted = ([int]$shape.Type -eq 13) } else { $isWanted = ([int]$shape.HasTextFrame -eq -1) }
        } catch { $isWanted = $false }
        if (-not $isWanted) { continue }
        $seen++
        if ($seen -eq [int]$ordinal) { return $i }
    }
    return $null
}

function Resolve-PictureIndex($slide, $p) { return Resolve-ShapeByKind $slide $p "picture" (Get-PropOrNull $p 'imageIndex') }

function Resolve-TextBoxIndex($slide, $p) { return Resolve-ShapeByKind $slide $p "textbox" (Get-PropOrNull $p 'textboxIndex') }

function Expand-WpsNestedParams($Action, $Params) {
    # Some tools hand over a nested object (style/shadow/border/gradient/element/rotation) while the
    # actions read flat keys. Merge the nested properties onto the parameter set. The accepted-key
    # check runs afterwards, so a nested property the action never reads still fails loudly.
    if ($null -eq $Params) { return $Params }
    $containers = $script:ActionNestedParams[$Action]
    if ($null -eq $containers) { return $Params }
    foreach ($container in $containers) {
        $obj = Get-PropOrNull $Params $container
        if ($null -eq $obj) { continue }
        foreach ($prop in @($obj.PSObject.Properties)) {
            if ($null -ne (Get-PropOrNull $Params $prop.Name)) { continue }
            Add-Member -InputObject $Params -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
        }
    }
    return $Params
}

function Add-WpsParamAliases($Action, $Params) {
    # The tool layer's public parameter names and this script's internal keys do not always match
    # (shapeIndex vs shapeName, order vs zOrder, url vs address ...). Rather than teach every case
    # two spellings, the generator emits one explicit alias table and the alias value is copied onto
    # the canonical key before the action runs. The same table feeds the accepted-key check, so an
    # alias can never be added here without the guard knowing about it.
    if ($null -eq $Params) { return $Params }
    $map = $script:ActionParamAliases[$Action]
    if ($null -eq $map) { return $Params }
    foreach ($alias in $map.Keys) {
        $canonical = $map[$alias]
        if ($null -ne (Get-PropOrNull $Params $canonical)) { continue }
        $value = Get-PropOrNull $Params $alias
        if ($null -eq $value) { continue }
        Add-Member -InputObject $Params -NotePropertyName $canonical -NotePropertyValue $value -Force
    }
    return $Params
}

function Test-WpsActionParamKeys($Action, $Params, $Accepted) {
    # Returns $null when every supplied parameter is one the action actually reads, otherwise a
    # message naming the unexpected keys.
    #
    # MCP does not validate parameter names against the bridge, and PowerShell happily ignores an
    # object property nobody reads. So a tool that spells a parameter differently from this script
    # used to be silently dropped: delete_sheet(name) deleted the active sheet, close_workbook(save)
    # ignored save=false, find_replace sent keys this file never looked at. Comparing the supplied
    # keys against the keys the action reads turns all of those into a loud error.
    if ($null -eq $Accepted) { return $null }
    $given = @()
    if ($null -ne $Params) { $given = @($Params.PSObject.Properties | ForEach-Object { $_.Name }) }
    if ($given.Count -eq 0) { return $null }
    $extra = @($given | Where-Object { $Accepted -notcontains $_ })
    if ($extra.Count -eq 0) { return $null }
    return ("unknown parameter(s) for '" + $Action + "': " + ($extra -join ', ') +
            " | accepted: " + ((@($Accepted) | Sort-Object) -join ', '))
}

function Get-WordMatchCount($doc, [string]$findText, [bool]$matchCase, [bool]$matchWholeWord) {
    # Word exposes no match-count API, so walk the document with Find. Wrap is set to wdFindStop
    # and the loop is bounded: a WPS quirk must not be able to hang the resident host.
    $count = 0
    if ([string]::IsNullOrEmpty($findText)) { return 0 }
    $contentEnd = [int]$doc.Content.End
    $search = $doc.Content
    $limit = 100000
    while ($count -lt $limit) {
        $finder = $search.Find
        $finder.ClearFormatting()
        $finder.Text = $findText
        $finder.Forward = $true
        $finder.MatchCase = $matchCase
        $finder.MatchWholeWord = $matchWholeWord
        $finder.Wrap = 0
        if (-not $finder.Execute()) { break }
        $count++
        $next = [int]$search.End
        if ($next -ge $contentEnd) { break }
        if ($next -le [int]$search.Start) { break }
        $search.SetRange($next, $contentEnd)
    }
    return $count
}

function Set-ComValue($target, [string]$member, $value) {
    $prop = Find-ComProperty $target $member
    if ($null -eq $prop) { throw ('Set-ComValue: member not found: ' + $member) }
    $prop.Value = ConvertTo-ComValue $value
}
function Get-ComValue($target, [string]$member) {
    $prop = Find-ComProperty $target $member
    if ($null -eq $prop) { throw ('Get-ComValue: member not found: ' + $member) }
    return $prop.Value
}
function Invoke-ComMethod($target, [string]$method, [object[]]$arguments) {
    if ($null -eq $target) { throw ('Invoke-ComMethod: null target for method ' + $method) }
    $m = $target.PSObject.Methods | Where-Object { $_.Name -eq $method } | Select-Object -First 1
    if ($null -eq $m) { throw ('Invoke-ComMethod: method not found: ' + $method) }
    if ($null -eq $arguments) { $arguments = @() }
    return $m.Invoke($arguments)
}

function Output-Json($obj) {
    $obj | ConvertTo-Json -Depth 10 -Compress
}

function Convert-ColumnLetterToNumber([string]$letters) {
    if (-not $letters) { return $null }
    $letters = $letters.ToUpper().Trim()
    $sum = 0
    foreach ($ch in $letters.ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { continue }
        $sum = ($sum * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $sum
}

function Convert-ColumnNumberToLetter([int]$col) {
    if ($col -le 0) { return $null }
    $letter = ""
    while ($col -gt 0) {
        $mod = ($col - 1) % 26
        $letter = [char](65 + $mod) + $letter
        $col = [Math]::Floor(($col - 1) / 26)
    }
    return $letter
}

function Convert-HexColorToRgbInt([string]$hex) {
    if (-not $hex) { return $null }
    $hex = $hex.Trim()
    if ($hex.StartsWith('#')) { $hex = $hex.Substring(1) }
    if ($hex.Length -eq 3) {
        $hex = ($hex[0] + $hex[0] + $hex[1] + $hex[1] + $hex[2] + $hex[2])
    }
    if ($hex.Length -ne 6) { return $null }
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

# Parse an A1-style address ("D1:D2", "$D$1:$E$2", "D1") into its top-left cell and span.
# Needed where WPS's range geometry (Range.Row/Column/Rows.Count) misbehaves in the resident host.
function Get-AddressSpan([string]$address) {
    if (-not $address) { return $null }
    $clean = $address.Replace('$', '').Replace(' ', '')
    if ($clean -notmatch '^[A-Za-z]+[0-9]+(:[A-Za-z]+[0-9]+)?$') { return $null }
    $parts = $clean.Split(':')
    $null = $parts[0] -match '^([A-Za-z]+)([0-9]+)$'
    $firstCol = Convert-ColumnLetterToNumber $matches[1]
    $firstRow = [int]$matches[2]
    if ($parts.Count -eq 2) {
        $null = $parts[1] -match '^([A-Za-z]+)([0-9]+)$'
        $lastCol = Convert-ColumnLetterToNumber $matches[1]
        $lastRow = [int]$matches[2]
    } else {
        $lastCol = $firstCol
        $lastRow = $firstRow
    }
    if ($null -eq $firstCol -or $null -eq $lastCol) { return $null }
    return @{
        Row = [Math]::Min($firstRow, $lastRow)
        Column = [Math]::Min($firstCol, $lastCol)
        RowCount = [Math]::Abs($lastRow - $firstRow) + 1
        ColumnCount = [Math]::Abs($lastCol - $firstCol) + 1
    }
}
# WPS's Range.Consolidate only understands R1C1 references, and only when the sheet part is quoted:
# 'Sheet'!R1C1:R4C2 consolidates, while A1-style (and an unquoted name containing a space) silently
# writes nothing at all. Excel accepts both forms. Measured with bare COM while giving this action
# its first tool (docs/FIXES.md 39), so the bridge converts instead of making callers speak R1C1.
function Get-SheetSettingsSnapshot($sheet) {
    # 一次把「页面设置 + 打印 + 页眉页脚 + 工作表外观」读全。每个字段单独 try/catch：WPS 对某些未设置过
    # 的属性会直接抛，而一个字段读不到不该让整份快照变成错误。
    $ps = $sheet.PageSetup
    $data = @{ sheet = $sheet.Name; orientation = $null; orientationName = ""; paperSize = $null }
    $fields = @(
        @('orientation', 'Orientation'), @('paperSize', 'PaperSize'), @('zoom', 'Zoom'),
        @('fitToPagesWide', 'FitToPagesWide'), @('fitToPagesTall', 'FitToPagesTall'),
        @('topMargin', 'TopMargin'), @('bottomMargin', 'BottomMargin'), @('leftMargin', 'LeftMargin'),
        @('rightMargin', 'RightMargin'), @('headerMargin', 'HeaderMargin'), @('footerMargin', 'FooterMargin'),
        @('centerHorizontally', 'CenterHorizontally'), @('centerVertically', 'CenterVertically'),
        @('printGridlines', 'PrintGridlines'), @('printHeadings', 'PrintHeadings'),
        @('printArea', 'PrintArea'), @('printTitleRows', 'PrintTitleRows'), @('printTitleColumns', 'PrintTitleColumns'),
        @('leftHeader', 'LeftHeader'), @('centerHeader', 'CenterHeader'), @('rightHeader', 'RightHeader'),
        @('leftFooter', 'LeftFooter'), @('centerFooter', 'CenterFooter'), @('rightFooter', 'RightFooter')
    )
    foreach ($field in $fields) {
        try { $data[$field[0]] = (Get-ComValue $ps $field[1]) } catch { $data[$field[0]] = $null }
    }
    if ([int]$data.orientation -eq 2) { $data.orientationName = "landscape" } else { $data.orientationName = "portrait" }
    try { $data.visible = [int]$sheet.Visible } catch { $data.visible = $null }
    try { $data.tabColor = [int]$sheet.Tab.Color } catch { $data.tabColor = $null }
    # Tab.Color 清掉之后读回的是 0（黑），不是“没设置”，只有 ColorIndex = xlColorIndexNone(-4142) 能分清。
    try { $data.tabColorIndex = [int]$sheet.Tab.ColorIndex } catch { $data.tabColorIndex = $null }
    try { $data.hPageBreaks = [int]$sheet.HPageBreaks.Count } catch { $data.hPageBreaks = 0 }
    return $data
}

function Get-RangeAddressSafe($range, [string]$fallback) {
    try {
        $addr = [string]$range.Address()
        if ($addr) { return $addr }
    } catch { }
    return $fallback
}

function ConvertTo-ConsolidateSource([string]$reference) {
    if (-not $reference) { return $null }
    $rangePart = $reference.Trim()
    $sheetPart = ''
    $bang = $rangePart.LastIndexOf('!')
    if ($bang -ge 0) {
        $sheetPart = $rangePart.Substring(0, $bang).Trim().Trim("'")
        $rangePart = $rangePart.Substring($bang + 1).Trim()
    }
    $span = Get-AddressSpan $rangePart
    if ($null -eq $span) { return $reference }
    $first = 'R' + $span.Row + 'C' + $span.Column
    $last = 'R' + ($span.Row + $span.RowCount - 1) + 'C' + ($span.Column + $span.ColumnCount - 1)
    $body = if ($first -eq $last) { $first } else { $first + ':' + $last }
    if ($sheetPart) { return "'" + $sheetPart + "'!" + $body }
    return $body
}



function Get-MainTextRange($word, $doc) {
    # 加过脚注/尾注之后，Word 会把光标留在注释正文里（Selection.StoryType 不再是正文），这时候再往
    # “光标处”插内容会插进注释而不是正文。凡是明确要动正文的动作都走这个解析器：不在正文就落到正文
    # 末尾，并且如实告警——而不是悄悄插错地方。
    $story = 1
    try { $story = [int]$word.Selection.StoryType } catch { $story = 1 }
    if ($story -eq 1) { return $word.Selection.Range }
    Add-WpsWarning "光标不在正文（可能停在脚注/尾注里），已在正文末尾插入"
    return $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
}

function Get-ListObjectByName($sheet, $table) {
    # Enumerate instead of calling Item(name) then Item(index): PowerShell caches a COM member's binder
    # after its first use, so the second argument shape fails. Enumeration sidesteps that entirely.
    $target = $null
    if ($null -ne $table -and "$table" -ne "") { $target = [string]$table }
    $seen = 0
    foreach ($lo in @($sheet.ListObjects)) {
        $seen = $seen + 1
        if ($null -eq $target) { return $lo }
        if ("$($lo.Name)" -eq $target) { return $lo }
    }
    if ($null -ne $target -and $target -match '^\d+$') {
        $idx = [int]$target
        if ($idx -ge 1 -and $idx -le $seen) { return $sheet.ListObjects.Item($idx) }
    }
    return $null
}

function Get-ListObjectAddress($lo, [string]$fallback) {
    # Range.Address() only answers for ranges the resident host reached itself, and a table must never
    # report an empty address, so fall back to whatever range the caller supplied.
    try {
        $addr = [string]$lo.Range.Address()
        if ($addr) { return $addr }
    } catch { }
    return $fallback
}

function Get-ListObjectInfo($lo, [string]$sheetName, [string]$fallback) {
    $columns = @()
    try {
        for ($i = 1; $i -le [int]$lo.ListColumns.Count; $i++) {
            $columns += @{ index = $i; name = [string]$lo.ListColumns.Item($i).Name }
        }
    } catch { Add-WpsWarning $_.Exception.Message }
    $rows = 0
    try { $rows = [int]$lo.ListRows.Count } catch { $rows = 0 }
    $hasTotals = $false
    try { $hasTotals = [bool]$lo.ShowTotals } catch { $hasTotals = $false }
    $style = ""
    try { $style = [string]$lo.TableStyle.Name } catch { $style = "" }
    $refs = @()
    foreach ($c in $columns) { $refs += ([string]$lo.Name + "[" + $c.name + "]") }
    return @{
        sheet = $sheetName
        name = [string]$lo.Name
        range = (Get-ListObjectAddress $lo $fallback)
        columns = $columns
        columnCount = $columns.Count
        rows = $rows
        hasTotals = $hasTotals
        tableStyle = $style
        structuredRefs = $refs
    }
}

function Get-RangeFromAddress($workbook, [string]$address) {
    # The leading comma on each return is load-bearing: PowerShell enumerates an enumerable return
    # value, and a multi-cell Range IS enumerable (A1:C7 comes back as 21 single-cell ranges).
    # Callers then handed that array to COM APIs - which is exactly why createPivotTable failed with
    # HRESULT 0x800A03EC at "createCache" in the resident host while the same call worked from a fresh
    # shell (FIXES 43). ",$range" emits the Range itself, not its cells.
    if ($address -match "^(?<sheet>[^!]+)!(?<range>.+)$") {
        $sheetName = $matches.sheet.Trim("'")
        return ,$workbook.Sheets.Item($sheetName).Range($matches.range)
    }
    return ,$workbook.ActiveSheet.Range($address)
}

function Get-AppTypeByExtension([string]$filePath) {
    $ext = [System.IO.Path]::GetExtension($filePath).ToLower().Trim('.')
    $wordExts = @('doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'rtf', 'txt', 'html', 'htm', 'xml', 'wps', 'wpt')
    $excelExts = @('xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm', 'csv', 'html', 'htm', 'et', 'ett')
    $pptExts = @('ppt', 'pptx', 'pptm', 'pot', 'potx', 'potm', 'pps', 'ppsx', 'ppsm', 'dps', 'dpt')
    if ($wordExts -contains $ext) { return 'word' }
    if ($excelExts -contains $ext) { return 'excel' }
    if ($pptExts -contains $ext) { return 'ppt' }
    return $null
}

function Get-WordSaveFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # Word WdSaveFormat: doc=0, docx=16, pdf=17, rtf=6, xps=18, html=8, txt=2, xml=11
    $map = @{ doc = 0; docx = 16; pdf = 17; rtf = 6; xps = 18; html = 8; htm = 8; txt = 2; xml = 11 }
    return $map[$f]
}

function Get-ExcelFileFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # Excel XlFileFormat: xls=-4143, xlsx=51, xlsm=52, xlsb=50, csv=6, html=44, xlt=17, xltx=54, xltm=53
    $map = @{ xls = -4143; xlsx = 51; xlsm = 52; xlsb = 50; csv = 6; html = 44; htm = 44; xlt = 17; xltx = 54; xltm = 53 }
    return $map[$f]
}

function Get-PptSaveFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # PowerPoint PpSaveAsFileType: ppt=1, pptx=24, pptm=25, pdf=32, png=18, jpg=17, gif=16, bmp=19
    $map = @{ ppt = 1; pptx = 24; pptm = 25; pdf = 32; png = 18; jpg = 17; jpeg = 17; gif = 16; bmp = 19 }
    return $map[$f]
}

try { $p = $Params | ConvertFrom-Json } catch { $p = @{} }

switch ($Action) {

    # ==================== Common ====================
    "ping" {
        Output-Json @{ success = $true; data = @{ message = "pong"; timestamp = [DateTimeOffset]::Now.ToUnixTimeMilliseconds() } }
    }

    "wireCheck" {
        Output-Json @{ success = $true; data = @{ message = "WPS MCP Bridge 已连接" } }
    }

    "getAppInfo" {
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $excel.Selection) } catch { Add-WpsWarning $_.Exception.Message }
            Output-Json @{ success = $true; data = @{ appType = "excel"; appName = $excel.Name; hasSelection = $hasSelection } }
            exit
        }
        $word = Get-WpsWord
        if ($null -ne $word) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $word.Selection) } catch { Add-WpsWarning $_.Exception.Message }
            Output-Json @{ success = $true; data = @{ appType = "word"; appName = $word.Name; hasSelection = $hasSelection } }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $ppt.ActiveWindow.Selection) } catch { Add-WpsWarning $_.Exception.Message }
            Output-Json @{ success = $true; data = @{ appType = "ppt"; appName = $ppt.Name; hasSelection = $hasSelection } }
            exit
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "getSelectedText" {
        $word = Get-WpsWord
        if ($null -ne $word) {
            try {
                $text = $word.Selection.Text
                Output-Json @{ success = $true; data = @{ text = $text.Trim() } }
            } catch {
                Output-Json @{ success = $false; error = "Word selection not available" }
            }
            exit
        }
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            try {
                $sel = $excel.Selection
                $text = if ($sel -is [string]) { $sel } else { $sel.Text }
                Output-Json @{ success = $true; data = @{ text = [string]$text } }
            } catch {
                Output-Json @{ success = $false; error = "Excel selection not available" }
            }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            try {
                $sel = $ppt.ActiveWindow.Selection
                if ($sel -and $sel.TextRange) {
                    $text = $sel.TextRange.Text
                    Output-Json @{ success = $true; data = @{ text = [string]$text } }
                } else {
                    Output-Json @{ success = $false; error = "PPT selection has no text" }
                }
            } catch {
                Output-Json @{ success = $false; error = "PPT selection not available" }
            }
            exit
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "setSelectedText" {
        $text = if ($null -ne $p.text) { [string]$p.text } else { "" }
        $word = Get-WpsWord
        if ($null -ne $word) {
            try {
                $word.Selection.Text = $text
                Output-Json @{ success = $true }
            } catch {
                Output-Json @{ success = $false; error = "Word selection not available" }
            }
            exit
        }
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            try {
                $sel = $excel.Selection
                Set-ComValue $sel 'Value2' $text
                Output-Json @{ success = $true }
            } catch {
                Output-Json @{ success = $false; error = "Excel selection not available" }
            }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            try {
                $sel = $ppt.ActiveWindow.Selection
                if ($sel -and $sel.TextRange) {
                    $sel.TextRange.Text = $text
                    Output-Json @{ success = $true }
                } else {
                    Output-Json @{ success = $false; error = "PPT selection has no text" }
                }
            } catch {
                Output-Json @{ success = $false; error = "PPT selection not available" }
            }
            exit
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "save" {
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $excel.ActiveWorkbook) {
            $excel.ActiveWorkbook.Save()
            Output-Json @{ success = $true; app = "excel" }
            exit
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $word.ActiveDocument) {
            $word.ActiveDocument.Save()
            Output-Json @{ success = $true; app = "word" }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $ppt.ActivePresentation.Save()
            Output-Json @{ success = $true; app = "ppt" }
            exit
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    "saveAs" {
        $path = $p.path
        if (-not $path) { Output-Json @{ success = $false; error = "Path required" }; exit }
        $appType = if ($p.appType) { $p.appType } else { Get-AppTypeByExtension $path }
        if ($appType -eq 'excel') {
            $excel = Get-WpsExcel
            if ($null -eq $excel -or $null -eq $excel.ActiveWorkbook) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
            $format = Get-ExcelFileFormat $p.format
            if ($null -ne $format) { $excel.ActiveWorkbook.SaveAs($path, $format) }
            else { $excel.ActiveWorkbook.SaveAs($path) }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "excel" } }
            exit
        }
        if ($appType -eq 'word') {
            $word = Get-WpsWord
            if ($null -eq $word -or $null -eq $word.ActiveDocument) { Output-Json @{ success = $false; error = "No active document" }; exit }
            $format = Get-WordSaveFormat $p.format
            if ($null -ne $format) { $word.ActiveDocument.SaveAs($path, $format) }
            else { $word.ActiveDocument.SaveAs($path) }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "word" } }
            exit
        }
        if ($appType -eq 'ppt') {
            $ppt = Get-WpsPpt
            if ($null -eq $ppt -or $null -eq $ppt.ActivePresentation) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
            $format = Get-PptSaveFormat $p.format
            if ($null -ne $format) { $ppt.ActivePresentation.SaveAs($path, $format) }
            else { $ppt.ActivePresentation.SaveAs($path) }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "ppt" } }
            exit
        }
        Output-Json @{ success = $false; error = "Unknown app type" }
    }

    "openFile" {
        $path = $p.path
        if (-not $path) { Output-Json @{ success = $false; error = "Path required" }; exit }
        $appType = if ($p.appType) { $p.appType } else { Get-AppTypeByExtension $path }
        if ($appType -eq 'excel') {
            $excel = Get-WpsExcel
            if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
            $prevAlerts = Set-WpsAlertsSuppressed $excel 'excel'
            $wb = $null
            $openError = $null
            try { $wb = Open-ExcelWorkbook $excel $path $null $false } catch { $openError = $_.Exception.Message }
            Restore-WpsAlerts $excel $prevAlerts
            if ($null -ne $openError) { Output-Json @{ success = $false; error = (Format-WpsOpenError $path 'excel' $openError) }; exit }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "excel"; name = $wb.Name } }
            exit
        }
        if ($appType -eq 'word') {
            $word = Get-WpsWord
            if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
            $prevAlerts = Set-WpsAlertsSuppressed $word 'word'
            $opened = $null
            $openError = $null
            try { $opened = Open-WordDocument $word $path } catch { $openError = $_.Exception.Message }
            Restore-WpsAlerts $word $prevAlerts
            if ($null -ne $openError) { Output-Json @{ success = $false; error = (Format-WpsOpenError $path 'word' $openError) }; exit }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "word"; name = $opened.Name } }
            exit
        }
        if ($appType -eq 'ppt') {
            $ppt = Get-WpsPpt
            if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
            $prevAlerts = Set-WpsAlertsSuppressed $ppt 'ppt'
            $opened = $null
            $openError = $null
            try { $opened = Open-PptPresentation $ppt $path } catch { $openError = $_.Exception.Message }
            Restore-WpsAlerts $ppt $prevAlerts
            if ($null -ne $openError) { Output-Json @{ success = $false; error = (Format-WpsOpenError $path 'ppt' $openError) }; exit }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "ppt"; name = $opened.Name } }
            exit
        }
        Output-Json @{ success = $false; error = "Unknown app type" }
    }

    "createDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $created = $null
        $createError = $null
        try { $created = $word.Documents.Add() } catch { $createError = $_.Exception.Message }
        if ($null -ne $createError) { Output-Json @{ success = $false; error = ("cannot create a document: " + $createError) }; exit }
        $data = @{}
        try {
            $newDoc = if ($null -ne $created) { $created } else { $word.ActiveDocument }
            if ($null -ne $newDoc) { $data.name = [string]$newDoc.Name }
        } catch { Add-WpsWarning ("could not read the new document name: " + $_.Exception.Message) }
        Output-Json @{ success = $true; data = $data }
    }

    "createPresentation" {
        # NOTE: the result must not be dereferenced - WPS's Presentations.Add() returns null, and its
        # new deck can have zero slides. Whether a deck is usable is answered by addSlide/getSlideCount.
        # A hollow application instance exposes Presentations but fails inside Add(); retry once
        # with a freshly acquired instance before giving up.
        $addError = $null
        for ($attempt = 1; $attempt -le 2; $attempt++) {
            try {
                $null = $ppt.Presentations.Add()
                $addError = $null
                break
            } catch {
                $addError = $_.Exception.Message
                if ($attempt -eq 1) { Reset-WpsApp 'ppt'; $ppt = Get-WpsPpt }
            }
        }
        if ($null -ne $addError) { Output-Json @{ success = $false; error = ("cannot create a presentation: " + $addError) }; exit }
        Output-Json @{ success = $true }
    }

    "convertToPDF" {
        # Without appType this used to take the first running application, and since Excel is
        # usually running, converting a Word document exported the workbook instead.
        $want = if ($null -ne $p.appType) { ([string]$p.appType).ToLower() } else { $null }
        $wantExcel = ($null -eq $want) -or $want -eq "excel" -or $want -eq "et"
        $wantWord = ($null -eq $want) -or $want -eq "word" -or $want -eq "wps"
        $wantPpt = ($null -eq $want) -or $want -eq "ppt" -or $want -eq "wpp"
        $openAfter = ($null -ne $p.openAfterExport) -and [bool]$p.openAfterExport
        $excel = if ($wantExcel) { Get-WpsExcel } else { $null }
        if ($null -ne $excel -and $null -ne $excel.ActiveWorkbook) {
            $wb = $excel.ActiveWorkbook
            $sourcePath = $wb.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $wb.ExportAsFixedFormat(0, $outputPath)
            $openResult = Open-ExportedFile $outputPath $openAfter
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "excel"; openAfterExport = $openResult } }
            exit
        }
        $word = if ($wantWord) { Get-WpsWord } else { $null }
        if ($null -ne $word -and $null -ne $word.ActiveDocument) {
            $doc = $word.ActiveDocument
            $sourcePath = $doc.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $doc.ExportAsFixedFormat($outputPath, 17)
            $openResult = Open-ExportedFile $outputPath $openAfter
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "word"; openAfterExport = $openResult } }
            exit
        }
        $ppt = if ($wantPpt) { Get-WpsPpt } else { $null }
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $pres = Get-TargetPres $ppt $p
            $sourcePath = $pres.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $pres.SaveAs($outputPath, 32)
            $openResult = Open-ExportedFile $outputPath $openAfter
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "ppt"; openAfterExport = $openResult } }
            exit
        }
        if ($null -ne $want) { Output-Json @{ success = $false; error = ("no running " + $want + " document to export") }; exit }
        Output-Json @{ success = $false; error = "No active document" }
    }

    "convertFormat" {
        $targetFormat = $p.targetFormat
        if (-not $targetFormat) { Output-Json @{ success = $false; error = "targetFormat required" }; exit }
        # Same first-running-application hazard as convertToPDF.
        $want = if ($null -ne $p.appType) { ([string]$p.appType).ToLower() } else { $null }
        $wantExcel = ($null -eq $want) -or $want -eq "excel" -or $want -eq "et"
        $wantWord = ($null -eq $want) -or $want -eq "word" -or $want -eq "wps"
        $wantPpt = ($null -eq $want) -or $want -eq "ppt" -or $want -eq "wpp"
        $excel = if ($wantExcel) { Get-WpsExcel } else { $null }
        if ($null -ne $excel -and $null -ne $excel.ActiveWorkbook) {
            $wb = $excel.ActiveWorkbook
            $sourcePath = $wb.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-ExcelFileFormat $targetFormat
            if ($null -ne $format) { $wb.SaveAs($outputPath, $format) } else { $wb.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "excel"; targetFormat = $targetFormat } }
            exit
        }
        $word = if ($wantWord) { Get-WpsWord } else { $null }
        if ($null -ne $word -and $null -ne $word.ActiveDocument) {
            $doc = $word.ActiveDocument
            $sourcePath = $doc.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-WordSaveFormat $targetFormat
            if ($null -ne $format) { $doc.SaveAs($outputPath, $format) } else { $doc.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "word"; targetFormat = $targetFormat } }
            exit
        }
        $ppt = if ($wantPpt) { Get-WpsPpt } else { $null }
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $pres = Get-TargetPres $ppt $p
            $sourcePath = $pres.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-PptSaveFormat $targetFormat
            if ($null -ne $format) { $pres.SaveAs($outputPath, $format) } else { $pres.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "ppt"; targetFormat = $targetFormat } }
            exit
        }
        if ($null -ne $want) { Output-Json @{ success = $false; error = ("no running " + $want + " document to convert") }; exit }
        Output-Json @{ success = $false; error = "No active document" }
    }

    # ==================== Excel Basic ====================
    "getActiveWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheets = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheetCount = $wb.Sheets.Count; sheets = $sheets } }
    }

    "getCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $cell = $sheet.Cells.Item([int]$p.row, [int]$p.col)
        Output-Json @{ success = $true; data = @{ value = $cell.Value2; text = $cell.Text; formula = $cell.Formula } }
    }

    "setCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $row = [int]$p.row
        $col = [int]$p.col
        # ConvertFrom-Json wraps scalars in PSObject; COM parameterised properties and Value2
        # reject the wrapped form, so every value crossing into COM is unwrapped explicitly.
        $value = $p.value
        if ($null -eq $value) { $value = $null }
        elseif ($value -is [bool]) { $value = [bool]$value }
        elseif ($value -is [int] -or $value -is [long] -or $value -is [double] -or $value -is [decimal]) { $value = [double]$value }
        else { $value = [string]$value }
        try {
            $cell = $sheet.Cells.Item($row, $col)
        # PowerShell caches the COM binder for a member after its first use, so a direct
        # ".Value2 = $v" call site fails as soon as the value type changes (Int32 vs String).
        # Going through PSPropertyInfo re-binds per call and accepts mixed types.
        $cell.PSObject.Properties['Value2'].Value = $value
        } catch {
            $valueType = if ($null -eq $value) { 'null' } else { $value.GetType().FullName }
            Output-Json @{ success = $false; error = $_.Exception.Message; errorType = $_.Exception.GetType().Name; valueType = $valueType; rowType = $p.row.GetType().FullName; colType = $p.col.GetType().FullName }
            exit
        }
        Output-Json @{ success = $true }
    }

    "getRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $range = $sheet.Range($p.range)
        $rawValue = $range.Value2
        $data = @()
        if ($null -eq $rawValue) {
            $data = @()
        } elseif ($rawValue -is [Array] -and $rawValue.Rank -eq 2) {
            $r0 = $rawValue.GetLowerBound(0); $r1 = $rawValue.GetUpperBound(0)
            $c0 = $rawValue.GetLowerBound(1); $c1 = $rawValue.GetUpperBound(1)
            for ($r = $r0; $r -le $r1; $r++) {
                $row = @()
                for ($c = $c0; $c -le $c1; $c++) { $row += $rawValue[$r, $c] }
                $data += ,@($row)
            }
        } else {
            $data = @(,@($rawValue))
        }
        Output-Json @{ success = $true; data = @{ data = $data; rows = $data.Count } }
    }

    "setRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $rows = @($p.data).Count
        if ($rows -eq 0) { Output-Json @{ success = $false; error = "data must not be empty" }; exit }
        $cols = 0
        for ($r = 0; $r -lt $rows; $r++) {
            $row = @($p.data[$r])
            if ($row.Count -gt $cols) { $cols = $row.Count }
        }
        if ($cols -eq 0) { Output-Json @{ success = $false; error = "data rows must not be empty" }; exit }
        # ConvertFrom-Json yields PSObject-wrapped numbers; assigning one straight to Value2 throws
        # (Int32 -> String cast). Unwrap into a typed matrix and write the whole block in one call.
        $matrix = New-Object 'object[,]' $rows, $cols
        for ($r = 0; $r -lt $rows; $r++) {
            $row = @($p.data[$r])
            for ($c = 0; $c -lt $row.Count; $c++) {
                $v = $row[$c]
                if ($null -eq $v) { $matrix[$r, $c] = $null }
                elseif ($v -is [bool]) { $matrix[$r, $c] = [bool]$v }
                elseif ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) { $matrix[$r, $c] = [double]$v }
                else { $matrix[$r, $c] = [string]$v }
            }
        }
        $target = $sheet.Range($p.range).Resize($rows, $cols)
        $target.PSObject.Properties['Value2'].Value = $matrix
        Output-Json @{ success = $true; data = @{ rows = $rows; columns = $cols; range = $p.range } }
    }

    "diagnoseFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $cell = $sheet.Range($p.cell)
        $value = $cell.Value2
        $formula = $cell.Formula
        $errorType = $null
        $diagnosis = ""
        $suggestion = ""
        $precedents = @()
        if ($value -is [string] -and $value.StartsWith('#')) {
            $errorType = $value
            switch ($errorType) {
                "#REF!" { $diagnosis = "引用了不存在的单元格或区域"; $suggestion = "检查引用区域是否被删除或移动" }
                "#N/A" { $diagnosis = "查找函数未找到匹配值"; $suggestion = "确认查找值存在，或检查匹配条件" }
                "#VALUE!" { $diagnosis = "参数类型不正确或运算类型不匹配"; $suggestion = "检查函数参数类型和引用单元格" }
                "#NAME?" { $diagnosis = "函数名或名称拼写错误"; $suggestion = "检查函数名是否正确" }
                "#DIV/0!" { $diagnosis = "除数为零"; $suggestion = "检查除数单元格，避免除以零" }
                "#NUM!" { $diagnosis = "数值无效或超出范围"; $suggestion = "检查函数参数范围" }
                "#NULL!" { $diagnosis = "交集为空"; $suggestion = "检查引用区域的交集是否存在" }
                default { $diagnosis = "未知错误"; $suggestion = "检查公式与引用" }
            }
        }
        try {
            $refs = $cell.DirectPrecedents
            if ($refs -ne $null) {
                foreach ($area in $refs.Areas) { $precedents += $area.Address() }
            }
        } catch { Add-WpsWarning $_.Exception.Message }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; formula = $formula; currentValue = $value; errorType = $errorType; diagnosis = $diagnosis; suggestion = $suggestion; precedents = $precedents } }
    }

    "cleanData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $range = $sheet.Range($p.range)
        $opsResult = @()
        foreach ($op in $p.operations) {
            $success = $true
            $message = ""
            switch ($op) {
                "trim" {
                    foreach ($cell in $range) {
                        if ($cell.Value2 -is [string]) { Set-ComValue $cell 'Value2' ($cell.Value2.Trim()) }
                    }
                    $message = "已去除前后空格"
                }
                "remove_duplicates" {
                    $colCount = $range.Columns.Count
                    $cols = @()
                    for ($i = 1; $i -le $colCount; $i++) { $cols += $i }
                    $range.RemoveDuplicates($cols, 1)
                    $message = "已删除重复行"
                }
                "unify_date" {
                    foreach ($cell in $range) {
                        try {
                            if ($cell.Value2) {
                                $dt = [DateTime]::FromOADate($cell.Value2)
                                Set-ComValue $cell 'Value2' $dt.ToString("yyyy-MM-dd")
                            }
                        } catch { Add-WpsWarning $_.Exception.Message }
                    }
                    $message = "已统一日期格式"
                }
                "remove_empty_rows" {
                    for ($i = $range.Rows.Count; $i -ge 1; $i--) {
                        $row = $range.Rows.Item($i)
                        $isEmpty = $true
                        foreach ($cell in $row.Cells) { if ($cell.Value2) { $isEmpty = $false; break } }
                        if ($isEmpty) { $row.Delete() }
                    }
                    $message = "已删除空行"
                }
                default { $success = $false; $message = "不支持的操作" }
            }
            $opsResult += @{ operation = $op; success = $success; message = $message }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; operations = $opsResult; message = "cleanData completed" } }
    }

    "createPivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sourceRange = Get-RangeFromAddress $wb $p.sourceRange
        $destSheet = if ($p.destinationSheet) { $wb.Sheets.Item($p.destinationSheet) } else { $excel.ActiveSheet }
        $destCell = $destSheet.Range($p.destinationCell)
        $pivotName = if ($p.tableName) { $p.tableName } else { "PivotTable" + [Guid]::NewGuid().ToString("N").Substring(0, 6) }
        # 每一步都记下来：这个 action 以前没有任何 try/catch，出错只剩一个 HRESULT，没法定位。
        $step = "createCache"
        try {
            $cache = $wb.PivotCaches().Create(1, $sourceRange)
            $step = "createTable"
            $table = $cache.CreatePivotTable($destCell, $pivotName)
            $step = "rowFields"
            foreach ($fieldName in $p.rowFields) {
                $field = $table.PivotFields($fieldName)
                $field.Orientation = 1
                $field.Position = 1
            }
            if ($p.columnFields) {
                $step = "columnFields"
                foreach ($fieldName in $p.columnFields) {
                    $field = $table.PivotFields($fieldName)
                    $field.Orientation = 2
                    $field.Position = 1
                }
            }
            if ($p.filterFields) {
                $step = "filterFields"
                foreach ($fieldName in $p.filterFields) {
                    $field = $table.PivotFields($fieldName)
                    $field.Orientation = 3
                    $field.Position = 1
                }
            }
            $step = "valueFields"
            foreach ($vf in $p.valueFields) {
                $field = $table.PivotFields($vf.field)
                $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
                $func = $funcMap[$vf.aggregation]
                if ($null -eq $func) { $func = -4157 }
                $table.AddDataField($field, $vf.field, $func) | Out-Null
            }
            $step = "readBack"
            $rowCount = [int]$table.RowRange.Rows.Count
            $columnCount = [int]$table.TableRange1.Columns.Count
            $location = $destCell.Address()
        } catch {
            Output-Json @{ success = $false; error = ($step + ": " + $_.Exception.Message) }; exit
        }
        Output-Json @{ success = $true; data = @{ pivotTableName = $pivotName; location = $location; rowCount = $rowCount; columnCount = $columnCount } }
    }

    "updatePivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $table = $null
        if ($p.pivotTableName) {
            try { $table = $sheet.PivotTables($p.pivotTableName) } catch { Add-WpsWarning $_.Exception.Message }
        }
        if ($null -eq $table -and $p.pivotTableCell) {
            try { $table = $sheet.Range($p.pivotTableCell).PivotTable } catch { Add-WpsWarning $_.Exception.Message }
        }
        if ($null -eq $table) { Output-Json @{ success = $false; error = "PivotTable not found" }; exit }
        $ops = @()
        if ($p.addRowFields) {
            foreach ($f in $p.addRowFields) {
                $field = $table.PivotFields($f); $field.Orientation = 1; $field.Position = 1
                $ops += @{ operation = "addRowFields"; success = $true; message = $f }
            }
        }
        if ($p.removeRowFields) {
            foreach ($f in $p.removeRowFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeRowFields"; success = $true; message = $f }
            }
        }
        if ($p.addColumnFields) {
            foreach ($f in $p.addColumnFields) {
                $field = $table.PivotFields($f); $field.Orientation = 2; $field.Position = 1
                $ops += @{ operation = "addColumnFields"; success = $true; message = $f }
            }
        }
        if ($p.removeColumnFields) {
            foreach ($f in $p.removeColumnFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeColumnFields"; success = $true; message = $f }
            }
        }
        if ($p.addFilterFields) {
            foreach ($f in $p.addFilterFields) {
                $field = $table.PivotFields($f); $field.Orientation = 3; $field.Position = 1
                $ops += @{ operation = "addFilterFields"; success = $true; message = $f }
            }
        }
        if ($p.removeFilterFields) {
            foreach ($f in $p.removeFilterFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeFilterFields"; success = $true; message = $f }
            }
        }
        if ($p.addValueFields) {
            foreach ($vf in $p.addValueFields) {
                $field = $table.PivotFields($vf.field)
                $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
                $func = $funcMap[$vf.aggregation]
                if ($null -eq $func) { $func = -4157 }
                $table.AddDataField($field, $vf.field, $func) | Out-Null
                $ops += @{ operation = "addValueFields"; success = $true; message = $vf.field }
            }
        }
        if ($p.updateValueFields) {
            foreach ($vf in $p.updateValueFields) {
                $field = $table.PivotFields($vf.field)
                $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
                $func = $funcMap[$vf.aggregation]
                if ($null -eq $func) { $func = -4157 }
                $field.Function = $func
                $ops += @{ operation = "updateValueFields"; success = $true; message = $vf.field }
            }
        }
        if ($p.removeValueFields) {
            foreach ($vf in $p.removeValueFields) {
                try { $table.DataFields($vf).Orientation = 0 } catch { Add-WpsWarning $_.Exception.Message }
                $ops += @{ operation = "removeValueFields"; success = $true; message = $vf }
            }
        }
        if ($p.refresh) { try { $table.RefreshTable() | Out-Null; $ops += @{ operation = "refresh"; success = $true; message = "refreshed" } } catch { Add-WpsWarning $_.Exception.Message } }
        Output-Json @{ success = $true; data = @{ pivotTableName = $table.Name; operations = $ops } }
    }

    "setFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $range = if ($p.range) { $sheet.Range($p.range) } else { $sheet.Cells.Item($p.row, $p.col) }
        $cellCount = $range.Cells.Count
        $span = Get-AddressSpan ([string]$p.range)
        if ($cellCount -eq 1) {
            $range.Formula = $p.formula
        } else {
            # A multi-cell target needs a different route, measured against the resident host:
            # assigning Formula on a multi-cell Range fails there ("this object has no property
            # Formula") where Excel broadcasts, and Range geometry (Row/Column/Rows.Count) plus
            # the one-argument Range.Cells.Item(index) also failed there. Parse the address and
            # write every cell through the worksheet's two-argument Cells accessor.
            if ($null -eq $span) {
                Output-Json @{ success = $false; error = ("cannot broadcast a formula over multiple cells for range '" + [string]$p.range + "'; use an A1 range or set one cell at a time") }
                exit
            }
            for ($r = 0; $r -lt $span.RowCount; $r++) {
                for ($c = 0; $c -lt $span.ColumnCount; $c++) {
                    $sheet.Cells.Item($span.Row + $r, $span.Column + $c).Formula = $p.formula
                }
            }
        }
        $setFormulaData = @{ formula = [string]$p.formula; cellCount = $cellCount }
        # A multi-cell range reads back as a 2D array, so read the target's first cell explicitly:
        # a value nobody read must not end up reported as null.
        try {
            $value = $null
            if ($null -ne $span) {
                $value = $sheet.Cells.Item($span.Row, $span.Column).Value2
            } elseif ($cellCount -eq 1) {
                $value = $range.Value2
            }
            if ($null -ne $value) { $setFormulaData.value = $value }
        } catch { Add-WpsWarning ("could not read back the computed value: " + $_.Exception.Message) }
        Output-Json @{ success = $true; data = $setFormulaData }
    }

    "setCellFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        try {
            $range = $sheet.Range([string]$p.range)
            # The tool may send a nested format object, flat properties, or both.
            $f = @{}
            if ($null -ne $p.format) {
                foreach ($prop in $p.format.PSObject.Properties) { $f[$prop.Name] = $prop.Value }
            }
            foreach ($name in @('bold','italic','fontSize','fontName','fontColor','bgColor','underline','strikethrough','horizontalAlignment','verticalAlignment','wrapText','numberFormat')) {
                $flat = $p.PSObject.Properties | Where-Object { $_.Name -eq $name } | Select-Object -First 1
                if ($null -ne $flat -and $null -ne $flat.Value) { $f[$name] = $flat.Value }
            }
            $applied = @()
            $font = $range.Font
            if ($null -ne $f['bold']) { Set-ComValue $font 'Bold' ([bool]$f['bold']); $applied += 'bold' }
            if ($null -ne $f['italic']) { Set-ComValue $font 'Italic' ([bool]$f['italic']); $applied += 'italic' }
            if ($null -ne $f['underline']) {
                $underlineValue = if ([bool]$f['underline']) { 2 } else { -4142 }
                Set-ComValue $font 'Underline' $underlineValue
                $applied += 'underline'
            }
            if ($null -ne $f['strikethrough']) { Set-ComValue $font 'Strikethrough' ([bool]$f['strikethrough']); $applied += 'strikethrough' }
            if ($null -ne $f['fontName'] -and "$($f['fontName'])" -ne "") { Set-ComValue $font 'Name' ([string]$f['fontName']); $applied += 'fontName' }
            if ($null -ne $f['fontSize'] -and [double]$f['fontSize'] -gt 0) { Set-ComValue $font 'Size' ([double]$f['fontSize']); $applied += 'fontSize' }
            if ($null -ne $f['fontColor']) {
                $fontRgb = Convert-HexColorToRgbInt([string]$f['fontColor'])
                if ($null -ne $fontRgb) { Set-ComValue $font 'Color' $fontRgb; $applied += 'fontColor' }
            }
            if ($null -ne $f['bgColor']) {
                $bgRgb = Convert-HexColorToRgbInt([string]$f['bgColor'])
                if ($null -ne $bgRgb) {
                    Set-ComValue $range.Interior 'Color' $bgRgb
                    Set-ComValue $range.Interior 'Pattern' 1
                    $applied += 'bgColor'
                }
            }
            if ($null -ne $f['horizontalAlignment']) {
                $alignMap = @{ left = -4131; center = -4108; right = -4152; justify = -4130 }
                $alignValue = $alignMap[[string]$f['horizontalAlignment']]
                if ($null -ne $alignValue) { Set-ComValue $range 'HorizontalAlignment' $alignValue; $applied += 'horizontalAlignment' }
            }
            if ($null -ne $f['verticalAlignment']) {
                $valignMap = @{ top = -4160; center = -4108; bottom = -4107 }
                $valignValue = $valignMap[[string]$f['verticalAlignment']]
                if ($null -ne $valignValue) { Set-ComValue $range 'VerticalAlignment' $valignValue; $applied += 'verticalAlignment' }
            }
            if ($null -ne $f['wrapText']) { Set-ComValue $range 'WrapText' ([bool]$f['wrapText']); $applied += 'wrapText' }
            if ($null -ne $f['numberFormat'] -and "$($f['numberFormat'])" -ne "") { Set-ComValue $range 'NumberFormat' ([string]$f['numberFormat']); $applied += 'numberFormat' }
            if ($applied.Count -eq 0) { Output-Json @{ success = $false; error = "no supported format property was provided" }; exit }
            Output-Json @{ success = $true; data = @{ range = [string]$p.range; sheet = $sheet.Name; applied = $applied } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "setCellStyle" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        # The tool schema also offers a named style; assign it through Excel's own Style property so
        # the workbook's real style definitions apply.
        $appliedStyle = $null
        if ($null -ne $p.style -and "$($p.style)" -ne "") {
            $appliedStyle = [string]$p.style
            try {
                Set-ComValue $range "Style" $appliedStyle
            } catch {
                Output-Json @{ success = $false; error = ("unknown cell style '" + $appliedStyle + "': " + $_.Exception.Message) }
                exit
            }
        }
        if ($p.fontSize) { $range.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $range.Font.Bold = [bool]$p.bold }
        if ($null -ne $p.italic) { $range.Font.Italic = [bool]$p.italic }
        if ($p.fontName) { $range.Font.Name = $p.fontName }
        if ($p.backgroundColor) {
            $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
            if ($null -ne $bg) { $range.Interior.Color = $bg }
        }
        if ($p.fontColor) {
            $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
            if ($null -ne $fc) { $range.Font.Color = $fc }
        }
        if ($p.horizontalAlignment) {
            $hAlignMap = @{ left = -4131; center = -4108; right = -4152 }
            $hAlign = $hAlignMap[$p.horizontalAlignment]
            if ($null -ne $hAlign) { $range.HorizontalAlignment = $hAlign }
        }
        if ($p.verticalAlignment) {
            $vAlignMap = @{ top = -4160; center = -4108; bottom = -4107 }
            $vAlign = $vAlignMap[$p.verticalAlignment]
            if ($null -ne $vAlign) { $range.VerticalAlignment = $vAlign }
        }
        if ($null -ne $p.border -and $p.border) {
            $range.Borders.LineStyle = 1
            if ($p.borderColor) {
                $bc = Convert-HexColorToRgbInt([string]$p.borderColor)
                if ($null -ne $bc) { $range.Borders.Color = $bc }
            }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; style = $appliedStyle } }
    }

    "setBorder" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $styleMap = @{ thin = 1; medium = 2; thick = 4; double = 6; none = 0 }
        $styleName = if ($null -ne $p.style) { $p.style } else { $p.borderStyle }
        $style = $styleMap[$styleName]
        if ($null -eq $style) { $style = 1 }
        $position = if ($p.position) { $p.position } else { "all" }
        $borders = @()
        if ($position -eq "all" -or $position -eq "outside") { $borders += 7, 8, 9, 10 }
        if ($position -eq "all" -or $position -eq "inside") { $borders += 11, 12 }
        if ($position -eq "left") { $borders += 7 }
        if ($position -eq "top") { $borders += 8 }
        if ($position -eq "bottom") { $borders += 9 }
        if ($position -eq "right") { $borders += 10 }
        $colorValue = $null
        if ($p.color) { $colorValue = Convert-HexColorToRgbInt([string]$p.color) }
        foreach ($b in $borders) {
            $border = $range.Borders.Item($b)
            if ($style -eq 0) {
                $border.LineStyle = -4142
            } else {
                $border.LineStyle = 1
                $border.Weight = $style
            }
            if ($null -ne $colorValue) { $border.Color = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; style = $styleName; position = $position; borders = $borders.Count } }
    }

    "copyFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $sourceRange = $sheet.Range($p.source)
        $targetRange = $sheet.Range($p.target)
        $sourceRange.Copy()
        $targetRange.PasteSpecial(-4122)
        $excel.CutCopyMode = $false
        Output-Json @{ success = $true; data = @{ source = $p.source; target = $p.target } }
    }

    "clearFormats" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $range.ClearFormats()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "addConditionalFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $formatType = if ($p.type) { $p.type } else { "cellValue" }
        # The tool schema passes a condition string (">100", "between(1,10)") and a named format,
        # neither of which this action used to read.
        $condOperator = $p.operator
        $condValue1 = if ($null -ne $p.value1) { $p.value1 } else { $p.value }
        $condValue2 = $p.value2
        if ($null -ne $p.condition -and "$($p.condition)" -ne "") {
            $conditionText = "$($p.condition)".Trim()
            $pairMatch = [regex]::Match($conditionText, "^[^,()]+\(?\s*([^,()]+)\s*,\s*([^,()]+)\s*\)?$")
            $compareMatch = [regex]::Match($conditionText, "^(>=|<=|<>|!=|>|<|=)\s*(.+)$")
            if ($conditionText -match "^between" -and $pairMatch.Success) {
                $condOperator = "between"
                $condValue1 = $pairMatch.Groups[1].Value.Trim()
                $condValue2 = $pairMatch.Groups[2].Value.Trim()
            } elseif ($compareMatch.Success) {
                $symbolMap = @{ ">" = "greaterThan"; "<" = "lessThan"; ">=" = "greaterThanOrEqual"; "<=" = "lessThanOrEqual"; "<>" = "notEqual"; "!=" = "notEqual"; "=" = "equal" }
                $condOperator = $symbolMap[$compareMatch.Groups[1].Value]
                $rest = $compareMatch.Groups[2].Value.Trim()
                $pair = [regex]::Match($rest, "^([^,]+)\s*,\s*(.+)$")
                if ($pair.Success) { $condValue1 = $pair.Groups[1].Value.Trim(); $condValue2 = $pair.Groups[2].Value.Trim() }
                else { $condValue1 = $rest }
            } else {
                $condOperator = "equal"
                $condValue1 = $conditionText
            }
        }
        $bgName = $p.backgroundColor
        $fgName = $p.fontColor
        $makeBold = $false
        if ($null -ne $p.format -and "$($p.format)" -ne "") {
            $fmtMap = @{
                red_fill    = @{ backgroundColor = "#FFC7CE"; fontColor = "#9C0006" }
                green_fill  = @{ backgroundColor = "#C6EFCE"; fontColor = "#006100" }
                yellow_fill = @{ backgroundColor = "#FFEB9C"; fontColor = "#9C6500" }
                red_font    = @{ fontColor = "#FF0000" }
                green_font  = @{ fontColor = "#006100" }
                bold        = @{ bold = $true }
            }
            $preset = $fmtMap["$($p.format)"]
            if ($null -ne $preset) {
                if ($preset.backgroundColor -and -not $bgName) { $bgName = $preset.backgroundColor }
                if ($preset.fontColor -and -not $fgName) { $fgName = $preset.fontColor }
                if ($preset.bold) { $makeBold = $true }
            } elseif ("$($p.format)" -match "^#[0-9A-Fa-f]{6}$") {
                $bgName = "$($p.format)"
            } else {
                Output-Json @{ success = $false; error = ("unknown conditional format '" + "$($p.format)" + "'; use red_fill/green_fill/yellow_fill/red_font/green_font/bold or a #RRGGBB colour") }
                exit
            }
        }
        if ($formatType -eq "cellValue") {
            $operatorMap = @{ greater = 5; greaterThan = 5; less = 6; lessThan = 6; equal = 3; notEqual = 4; greaterEqual = 7; greaterThanOrEqual = 7; lessEqual = 8; lessThanOrEqual = 8; between = 1 }
            $op = $operatorMap[$condOperator]
            if ($null -eq $op) { $op = 3 }
            $val1 = $condValue1
            $val2 = $condValue2
            if ($null -eq $val1) { Output-Json @{ success = $false; error = "condition/value1 is required for a cell-value conditional format" }; exit }
            $cf = $range.FormatConditions.Add(1, $op, $val1, $val2)
            if ($null -ne $cf -and $bgName) {
                $bg = Convert-HexColorToRgbInt([string]$bgName)
                if ($null -ne $bg) { $cf.Interior.Color = $bg }
            }
            if ($null -ne $cf -and $fgName) {
                $fc = Convert-HexColorToRgbInt([string]$fgName)
                if ($null -ne $fc) { $cf.Font.Color = $fc }
            }
            if ($null -ne $cf -and $makeBold) { $cf.Font.Bold = $true }
        } elseif ($formatType -eq "colorScale") {
            $scaleType = if ($p.colorScaleType) { $p.colorScaleType } else { 3 }
            $range.FormatConditions.AddColorScale($scaleType)
        } elseif ($formatType -eq "dataBar") {
            $range.FormatConditions.AddDatabar() | Out-Null
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $formatType } }
    }

    "removeConditionalFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        if ($p.index) {
            $range.FormatConditions.Item([int]$p.index).Delete()
        } else {
            $range.FormatConditions.Delete()
        }
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "getConditionalFormats" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $formats = @()
        $count = $range.FormatConditions.Count
        for ($i = 1; $i -le $count; $i++) {
            $cf = $range.FormatConditions.Item($i)
            $formats += @{ index = $i; type = $cf.Type }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; formats = $formats; count = $count } }
    }

    "addDataValidation" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $typeMap = @{ list = 3; whole = 1; decimal = 2; date = 4; time = 5; textLength = 6; custom = 7 }
        $typeName = if ($null -ne $p.validationType) { $p.validationType } else { $p.type }
        $validationType = $typeMap[$typeName]
        if ($null -eq $validationType) { $validationType = 3 }
        $range.Validation.Delete()
        if ($typeName -eq "list") {
            $listFormula = if ($p.formula1) { $p.formula1 } elseif ($p.list) { ($p.list -join ",") } elseif ($p.formula) { $p.formula } else { "" }
            $range.Validation.Add($validationType, 1, 1, $listFormula)
            if ($null -ne $p.showDropdown -and -not [bool]$p.showDropdown) {
                $range.Validation.InCellDropdown = $false
            } else {
                $range.Validation.InCellDropdown = $true
            }
        } else {
            $operatorMap = @{ between = 1; notBetween = 2; equal = 3; notEqual = 4; greater = 5; less = 6; greaterEqual = 7; lessEqual = 8 }
            $op = $operatorMap[$p.operator]
            if ($null -eq $op) { $op = 1 }
            $formula1 = if ($null -ne $p.formula1) { $p.formula1 } else { $p.formula }
            if ($null -eq $formula1) { Output-Json @{ success = $false; error = "formula1/formula required for validation type '$typeName'" }; exit }
            $range.Validation.Add($validationType, 1, $op, $formula1, $p.formula2)
        }
        if ($p.inputTitle -or $p.inputMessage) {
            $range.Validation.InputTitle = if ($p.inputTitle) { $p.inputTitle } else { "" }
            $range.Validation.InputMessage = if ($p.inputMessage) { $p.inputMessage } else { "" }
        }
        if ($p.errorTitle -or $p.errorMessage) {
            $range.Validation.ErrorTitle = if ($p.errorTitle) { $p.errorTitle } else { "" }
            $range.Validation.ErrorMessage = if ($p.errorMessage) { $p.errorMessage } else { "" }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $typeName } }
    }

    "removeDataValidation" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $range.Validation.Delete()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "getDataValidations" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $validation = $range.Validation
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $validation.Type; formula1 = $validation.Formula1; formula2 = $validation.Formula2; inputTitle = $validation.InputTitle; inputMessage = $validation.InputMessage } }
    }

    "getFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $cell = $sheet.Range($p.cell)
        $formula = if ($cell.Formula) { $cell.Formula } else { "" }
        $formulaLocal = ""
        try { $formulaLocal = $cell.FormulaLocal } catch { $formulaLocal = "" }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; formula = $formula; formulaLocal = $formulaLocal; hasFormula = ($formula -like "=*") } }
    }

    "setArrayFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $range.FormulaArray = $p.formula
        Output-Json @{ success = $true; data = @{ range = $p.range; formula = $p.formula } }
    }

    "getCellInfo" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        try {
            $cell = $sheet.Range([string]$p.cell)
            $value = Get-ComValue $cell 'Value2'
            $formula = [string](Get-ComValue $cell 'Formula')
            $numberFormat = [string](Get-ComValue $cell 'NumberFormat')
            $font = $cell.Font
            Output-Json @{ success = $true; data = @{
                cell = [string]$p.cell; sheet = $sheet.Name; value = $value; formula = $formula; numberFormat = $numberFormat
                font = @{ name = [string](Get-ComValue $font 'Name'); size = [double](Get-ComValue $font 'Size'); bold = [bool](Get-ComValue $font 'Bold'); italic = [bool](Get-ComValue $font 'Italic'); color = [double](Get-ComValue $font 'Color') }
                backgroundColor = [double](Get-ComValue $cell.Interior 'Color')
                horizontalAlignment = [double](Get-ComValue $cell 'HorizontalAlignment')
                verticalAlignment = [double](Get-ComValue $cell 'VerticalAlignment')
                wrapText = [bool](Get-ComValue $cell 'WrapText')
            } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "refreshLinks" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $links = $wb.LinkSources(1)
        if ($links) {
            for ($i = 1; $i -le $links.Length; $i++) {
                $wb.UpdateLink($links[$i - 1], 1)
            }
            Output-Json @{ success = $true; data = @{ refreshed = $links.Length } }
        } else {
            Output-Json @{ success = $true; data = @{ refreshed = 0; message = "没有外部链接" } }
        }
    }

    "consolidate" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $destRange = $sheet.Range($p.destination)
        # XlConsolidationFunction constants. The old map (9/2/1/4/5) was not a real enum, so every
        # call died with HRESULT 0x800A03EC - nobody saw it because no tool reached this action
        # until P2 gave it one and a test (FIXES 39).
        $funcMap = @{ sum = -4157; count = -4112; average = -4106; max = -4136; min = -4139 }
        $func = $funcMap[$p.function]
        if ($null -eq $func) { $func = -4157 }
        # Sources must reach COM as a plain string array of references: ConvertFrom-Json wraps every
        # element in a PSObject, and Range.Consolidate accepts the wrapped array without complaining
        # and then writes nothing at all (measured while giving this action its first tool).
        $sourceList = [string[]]@($p.sources | ForEach-Object { ConvertTo-ConsolidateSource ([string]$_) })
        $destRange.Consolidate($sourceList, $func, [bool]$p.topRow, [bool]$p.leftColumn, [bool]$p.createLinks)
        Output-Json @{ success = $true; data = @{ destination = $p.destination; sources = $p.sources } }
    }

    "calculateSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        if ($p.all) {
            $excel.Calculate()
            Output-Json @{ success = $true; data = @{ calculated = "all" } }
        } else {
            $sheet = Get-WorksheetByParam $excel $p
            $sheet.Calculate()
            Output-Json @{ success = $true; data = @{ calculated = $sheet.Name } }
        }
    }

    # ==================== Excel Advanced ====================
    "getExcelContext" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $usedRange = $sheet.UsedRange
        $headers = @()
        if ($usedRange.Rows.Count -gt 0) {
            $headerRow = $usedRange.Rows.Item(1)
            for ($i = 1; $i -le [Math]::Min($headerRow.Columns.Count, 26); $i++) {
                $headers += @{ column = [char](64 + $i); value = $headerRow.Cells.Item(1, $i).Value2 }
            }
        }
        $sheets = @(); for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{
            workbookName = $wb.Name; currentSheet = $sheet.Name; allSheets = $sheets
            selectedCell = $excel.Selection.Address(); headers = $headers; usedRange = $usedRange.Address(); usedRangeAddress = $usedRange.Address()
        }}
    }

    "getContext" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $usedRange = $sheet.UsedRange
        $headers = @()
        if ($usedRange.Rows.Count -gt 0) {
            $headerRow = $usedRange.Rows.Item(1)
            for ($i = 1; $i -le [Math]::Min($headerRow.Columns.Count, 26); $i++) {
                $headers += @{ column = [char](64 + $i); value = $headerRow.Cells.Item(1, $i).Value2 }
            }
        }
        $sheets = @(); for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{
            workbookName = $wb.Name; currentSheet = $sheet.Name; allSheets = $sheets
            selectedCell = $excel.Selection.Address(); headers = $headers; usedRangeAddress = $usedRange.Address(); usedRange = $usedRange.Address()
        }}
    }

    "sortRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        # The tool schema passes a 1-based column number and a boolean; this action's own keys are
        # a key range and an order string.
        $keyRef = if ($null -ne $p.keyColumn) { $p.keyColumn } else { $p.column }
        if ($null -eq $keyRef) { Output-Json @{ success = $false; error = "keyColumn/column required" }; exit }
        $keyCol = if ($keyRef -is [int]) { $range.Cells(1, [int]$keyRef) } else { $sheet.Range([string]$keyRef) }
        $descending = ($p.order -eq "desc") -or ($null -ne $p.ascending -and -not [bool]$p.ascending)
        $order = if ($descending) { 2 } else { 1 }
        $range.Sort($keyCol, $order)
        Output-Json @{ success = $true }
    }

    "autoFilter" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $field = if ($null -ne $p.field) { $p.field } else { $p.column }
        if ($null -ne $p.criteria) {
            if ($null -eq $field) { Output-Json @{ success = $false; error = "field/column is required when criteria is given" }; exit }
            $range.AutoFilter($field, $p.criteria)
        } else {
            $range.AutoFilter()
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; field = $field; criteria = $p.criteria } }
    }

    "createChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.dataRange)
        $chartType = $p.chartType
        if ($null -eq $chartType) {
            $chartTypes = @{ column = 51; column_clustered = 51; column_stacked = 52; bar = 57; bar_clustered = 57; line = 4; line_markers = 65; pie = 5; doughnut = -4120; area = 1; scatter = -4169; radar = -4151 }
            $chartType = $chartTypes[$p.chartTypeName]
        }
        if ($null -eq $chartType) { $chartType = 51 }
        $chartTypeName = if ($p.chartTypeName) { $p.chartTypeName } else { $null }
        $left = if ($p.position -and $p.position.left) { $p.position.left } elseif ($p.left) { $p.left } else { $range.Left + $range.Width + 20 }
        $top = if ($p.position -and $p.position.top) { $p.position.top } elseif ($p.top) { $p.top } else { $range.Top }
        $width = if ($p.position -and $p.position.width) { $p.position.width } else { 400 }
        $height = if ($p.position -and $p.position.height) { $p.position.height } else { 300 }
        $chartObj = $sheet.ChartObjects().Add($left, $top, $width, $height)
        $chartObj.Chart.SetSourceData($range)
        $chartObj.Chart.ChartType = $chartType
        if ($p.title) { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
        if ($null -ne $p.showLegend) { $chartObj.Chart.HasLegend = [bool]$p.showLegend }
        if ($p.showDataLabels) { $chartObj.Chart.ApplyDataLabels() }
        Output-Json @{ success = $true; data = @{ chartName = $chartObj.Name; chartIndex = $chartObj.Index; dataRange = $p.dataRange; chartType = $chartTypeName; position = @{ left = $left; top = $top; width = $width; height = $height } } }
    }

    "updateChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = if ($p.sheet) { $excel.ActiveWorkbook.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $chartObj = $null
        if ($p.chartName) {
            try { $chartObj = $sheet.ChartObjects($p.chartName) } catch { Add-WpsWarning $_.Exception.Message }
        }
        if ($null -eq $chartObj -and $p.chartIndex) {
            try { $chartObj = $sheet.ChartObjects().Item([int]$p.chartIndex) } catch { Add-WpsWarning $_.Exception.Message }
        }
        if ($null -eq $chartObj) { Output-Json @{ success = $false; error = "Chart not found" }; exit }
        $updated = @()
        if ($null -ne $p.title) {
            if ($p.title -eq "") { $chartObj.Chart.HasTitle = $false }
            else { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
            $updated += "title"
        }
        if ($null -ne $p.chartType) {
            $chartObj.Chart.ChartType = $p.chartType
            $updated += "chartType"
        }
        if ($null -ne $p.showLegend) {
            $chartObj.Chart.HasLegend = [bool]$p.showLegend
            $updated += "showLegend"
        }
        if ($null -ne $p.legendPosition) {
            $legendMap = @{ bottom = -4107; top = -4160; left = -4131; right = -4152 }
            $pos = $legendMap[$p.legendPosition]
            if ($pos) { $chartObj.Chart.Legend.Position = $pos; $updated += "legendPosition" }
        }
        if ($null -ne $p.showDataLabels) {
            if ($p.showDataLabels) {
                $chartObj.Chart.ApplyDataLabels()
            } else {
                try {
                    $series = $chartObj.Chart.SeriesCollection()
                    for ($i = 1; $i -le $series.Count; $i++) {
                        try { $series.Item($i).DataLabels().Delete() } catch { Add-WpsWarning $_.Exception.Message }
                    }
                } catch { Add-WpsWarning $_.Exception.Message }
            }
            $updated += "showDataLabels"
        }
        if ($p.dataRange) {
            $chartObj.Chart.SetSourceData($sheet.Range($p.dataRange))
            $updated += "dataRange"
        }
        if ($p.colors -and $p.colors.Count -gt 0) {
            $series = $chartObj.Chart.SeriesCollection()
            for ($i = 1; $i -le $series.Count -and $i -le $p.colors.Count; $i++) {
                $color = Convert-HexColorToRgbInt($p.colors[$i - 1])
                if ($null -ne $color) { $series.Item($i).Format.Fill.ForeColor.RGB = $color }
            }
            $updated += "colors"
        }
        Output-Json @{ success = $true; data = @{ chartName = $chartObj.Name; updatedProperties = $updated } }
    }

    "exportChartAsImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = if ($p.sheet) { $excel.ActiveWorkbook.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $outputPath = if ($p.outputPath) { $p.outputPath } else { $p.path }
        if ([string]::IsNullOrEmpty($outputPath)) { Output-Json @{ success = $false; error = "Missing outputPath" }; exit }
        $chartName = $p.chartName
        if ([string]::IsNullOrEmpty($chartName)) { Output-Json @{ success = $false; error = "Missing chartName" }; exit }
        $rawFormat = if ($p.format) { $p.format.ToString().ToUpper() } else { "PNG" }
        # JPEG 在 Excel COM 中按 JPG 滤镜处理
        $filterName = if ($rawFormat -eq "JPEG") { "JPG" } else { $rawFormat }
        $chartObj = $sheet.ChartObjects($chartName)
        $chartObj.Chart.Export($outputPath, $filterName)
        Output-Json @{ success = $true; data = @{ chartName = $chartName; outputPath = $outputPath; format = $filterName } }
    }

    "exportRangeAsImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = if ($p.sheet) { $excel.ActiveWorkbook.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $outputPath = if ($p.outputPath) { $p.outputPath } else { $p.path }
        if ([string]::IsNullOrEmpty($outputPath)) { Output-Json @{ success = $false; error = "Missing outputPath" }; exit }
        if ([string]::IsNullOrEmpty($p.range)) { Output-Json @{ success = $false; error = "Missing range" }; exit }
        $rawFormat = if ($p.format) { $p.format.ToString().ToUpper() } else { "PNG" }
        $filterName = if ($rawFormat -eq "JPEG") { "JPG" } else { $rawFormat }
        $range = $sheet.Range($p.range)
        # xlScreen=1 (Appearance), xlBitmap=2 (Format)
        $tempChart = $null
        try {
            $range.CopyPicture(1, 2)
            $tempChart = $sheet.ChartObjects().Add(0, 0, $range.Width, $range.Height)
            $tempChart.Activate()
            $tempChart.Chart.Paste()
            $tempChart.Chart.Export($outputPath, $filterName)
            $tempChart.Delete()
            $tempChart = $null
            Output-Json @{ success = $true; data = @{ range = $p.range; outputPath = $outputPath; format = $filterName } }
        } catch {
            # 异常清理：剪贴板冲突时回滚临时图表
            if ($null -ne $tempChart) { try { $tempChart.Delete() } catch { Add-WpsWarning $_.Exception.Message } }
            Output-Json @{ success = $false; error = "导出区域为图片失败: $($_.Exception.Message)" }
        }
    }

    "removeDuplicates" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $originalCount = $range.Rows.Count
        $cols = @()
        if ($p.columns -and $p.columns.Count -gt 0) {
            foreach ($col in $p.columns) {
                if ($col -is [int]) { $cols += $col }
                else {
                    $num = Convert-ColumnLetterToNumber([string]$col)
                    if ($num) { $cols += $num }
                }
            }
        }
        if ($cols.Count -eq 0) { $cols = @(1) }
        $hasHeader = if ($null -ne $p.hasHeader) { [int]([bool]$p.hasHeader) } else { 1 }
        $range.RemoveDuplicates($cols, $hasHeader)
        $remainingCount = $range.Rows.Count
        $removedCount = $originalCount - $remainingCount
        Output-Json @{ success = $true; data = @{ originalCount = $originalCount; removedCount = $removedCount; remainingCount = $remainingCount } }
    }

    "createSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        try {
            # position is 0-based, as the tool schema states.
            $position = Get-PropOrNull $p 'position'
            # Omitted position means append at the end, which is what the tool schema promises.
            $wanted = $wb.Sheets.Count
            if ($null -ne $position) { $wanted = [int]$position }
            if ($wanted -lt 0) { Output-Json @{ success = $false; error = "position must be 0 or greater (0-based sheet index)" }; exit }
            if ($wanted -eq 0) {
                $sheet = $wb.Sheets.Add($wb.Sheets.Item(1))
            } else {
                $afterIndex = [Math]::Min($wanted, $wb.Sheets.Count)
                if ($afterIndex -lt 1) { $afterIndex = 1 }
                $sheet = $wb.Sheets.Add($null, $wb.Sheets.Item($afterIndex))
            }
            $requestedName = Get-PropOrNull $p 'name'
            if ($null -ne $requestedName -and "$requestedName" -ne "") { Set-ComValue $sheet 'Name' ([string]$requestedName) }
            # index is the 1-based Excel ordinal; position is the 0-based index the caller asked in.
            Output-Json @{ success = $true; data = @{ name = $sheet.Name; index = $sheet.Index; position = [int]$sheet.Index - 1; requestedPosition = $wanted; sheetCount = $wb.Sheets.Count; sheetName = $sheet.Name; sheetIndex = $sheet.Index } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "deleteSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        # Deleting is destructive: never fall back to the active sheet, require an explicit target.
        $sheet = Resolve-Worksheet $excel $wb $p -RequireName
        if ($null -eq $sheet) { Output-Json @{ success = $false; error = "sheet name is required to delete a sheet" }; exit }
        if ($wb.Sheets.Count -le 1) { Output-Json @{ success = $false; error = "cannot delete the only sheet in a workbook" }; exit }
        try {
            $name = $sheet.Name
            $excel.DisplayAlerts = $false
            $sheet.Delete()
            $excel.DisplayAlerts = $true
            Output-Json @{ success = $true; data = @{ deletedSheet = $name; remaining = $wb.Sheets.Count } }
        } catch {
            $excel.DisplayAlerts = $true
            Output-Json @{ success = $false; error = $_.Exception.Message }
        }
    }

    "renameSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        try {
            $sheet = Resolve-Worksheet $excel $wb $p -RequireName
            if ($null -eq $sheet) { Output-Json @{ success = $false; error = "oldName is required to rename a sheet" }; exit }
            $oldName = $sheet.Name
            Set-ComValue $sheet 'Name' ([string]$p.newName)
            Output-Json @{ success = $true; data = @{ oldName = $oldName; newName = $sheet.Name } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "copySheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        try {
            $sheet = Resolve-Worksheet $excel $wb $p -RequireName
            if ($null -eq $sheet) { Output-Json @{ success = $false; error = "name is required to copy a sheet" }; exit }
            # position is 0-based: the copy should end up at that index.
            $position = Get-PropOrNull $p 'position'
            $wanted = -1
            if ($null -ne $position) { $wanted = [int]$position }
            if ($wanted -lt 0) { $wanted = $wb.Sheets.Count }
            if ($wanted -eq 0) {
                $sheet.Copy($wb.Sheets.Item(1), $null)
            } else {
                $afterIndex = [Math]::Min($wanted, $wb.Sheets.Count)
                if ($afterIndex -lt 1) { $afterIndex = $wb.Sheets.Count }
                $sheet.Copy($null, $wb.Sheets.Item($afterIndex))
            }
            $copy = $excel.ActiveSheet
            $requestedName = Get-PropOrNull $p 'newName'
            if ($null -ne $requestedName -and "$requestedName" -ne "") { Set-ComValue $copy 'Name' ([string]$requestedName) }
            Output-Json @{ success = $true; data = @{ sourceName = $sheet.Name; copiedFrom = $sheet.Name; newName = $copy.Name; index = $copy.Index; position = [int]$copy.Index - 1; requestedPosition = $wanted; sheetCount = $wb.Sheets.Count; sheetIndex = $copy.Index } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "getSheetList" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheets = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
            $sheet = $wb.Sheets.Item($i)
            $sheets += @{ name = $sheet.Name; index = $i; visible = $sheet.Visible }
        }
        Output-Json @{ success = $true; data = @{ sheets = $sheets; count = $sheets.Count; activeSheet = $excel.ActiveSheet.Name } }
    }

    "switchSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        try {
            $sheet = Resolve-Worksheet $excel $wb $p -RequireName
            if ($null -eq $sheet) { Output-Json @{ success = $false; error = "name is required to switch sheets" }; exit }
            $null = Invoke-ComMethod $sheet 'Activate' @()
            Output-Json @{ success = $true; data = @{ activeSheet = $sheet.Name } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "moveSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        try {
            $sheet = Resolve-Worksheet $excel $wb $p -RequireName
            if ($null -eq $sheet) { Output-Json @{ success = $false; error = "name is required to move a sheet" }; exit }
            $position = Get-PropOrNull $p 'position'
            if ($null -eq $position) { Output-Json @{ success = $false; error = "position is required to move a sheet" }; exit }
            # position is 0-based; Excel's Index is 1-based.
            $currentIndex = [int]$sheet.Index
            $target = [int]$position + 1
            $count = $wb.Sheets.Count
            if ($target -lt 1) { $target = 1 }
            if ($target -gt $count) { $target = $count }
            if ($target -lt $currentIndex) {
                $sheet.Move($wb.Sheets.Item($target), $null)
            } elseif ($target -gt $currentIndex) {
                $sheet.Move($null, $wb.Sheets.Item($target))
            }
            Output-Json @{ success = $true; data = @{ name = $sheet.Name; movedSheet = $sheet.Name; position = [int]$sheet.Index - 1; newPosition = [int]$sheet.Index - 1; requestedPosition = [int]$position; sheetCount = $wb.Sheets.Count; sheetIndex = $sheet.Index } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "mergeCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $across = if ($null -ne $p.across) { [bool]$p.across } else { $false }
        $range.Merge($across)
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "unmergeCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $range.UnMerge()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "setColumnWidth" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $col = $p.column
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $sheet.Range("${col}:${col}").ColumnWidth = $p.width
        Output-Json @{ success = $true; data = @{ column = $col; width = $p.width } }
    }

    "setRowHeight" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $sheet.Range("$($p.row):$($p.row)").RowHeight = $p.height
        Output-Json @{ success = $true; data = @{ row = $p.row; height = $p.height } }
    }

    "autoFitColumn" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if ($p.range) {
            $sheet.Range($p.range).Columns.AutoFit()
        } elseif ($p.column) {
            $col = $p.column
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("${col}:${col}").AutoFit()
        } else {
            $sheet.UsedRange.Columns.AutoFit()
        }
        Output-Json @{ success = $true; data = @{ message = "列宽已自动调整" } }
    }

    "autoFitRow" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if ($p.range) {
            $sheet.Range($p.range).Rows.AutoFit()
        } elseif ($p.row) {
            $sheet.Range("$($p.row):$($p.row)").AutoFit()
        } else {
            $sheet.UsedRange.Rows.AutoFit()
        }
        Output-Json @{ success = $true; data = @{ message = "行高已自动调整" } }
    }

    "autoFitAll" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $range.Columns.AutoFit()
        $range.Rows.AutoFit()
        Output-Json @{ success = $true; data = @{ message = "列宽行高已自动调整" } }
    }

    "setNumberFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        if ("$($p.format)" -eq "") { Output-Json @{ success = $false; error = "format must not be empty" }; exit }
        try {
            $sheet = Resolve-Worksheet $excel $wb $p
            $range = $sheet.Range([string]$p.range)
            Set-ComValue $range 'NumberFormat' ([string]$p.format)
            Output-Json @{ success = $true; data = @{ range = [string]$p.range; format = [string]$p.format; sheet = $sheet.Name } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "wrapText" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $wrap = if ($null -ne $p.wrap) { [bool]$p.wrap } else { $true }
        $range.WrapText = $wrap
        Output-Json @{ success = $true; data = @{ range = $p.range; wrapText = $wrap } }
    }

    "setPrintArea" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if ($p.range) { $sheet.PageSetup.PrintArea = $p.range } else { $sheet.PageSetup.PrintArea = "" }
        Output-Json @{ success = $true; data = @{ printArea = if ($p.range) { $p.range } else { "cleared" } } }
    }

    "getSelection" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sel = $excel.Selection
        if ($null -eq $sel) { Output-Json @{ success = $false; error = "No selection" }; exit }
        $addr = $sel.Address()
        Output-Json @{ success = $true; data = @{ address = $addr; rows = $sel.Rows.Count; columns = $sel.Columns.Count } }
    }

    "clearRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $clearType = if ($p.type) { $p.type } else { "all" }
        if ($clearType -eq "contents") { $range.ClearContents() }
        elseif ($clearType -eq "formats") { $range.ClearFormats() }
        elseif ($clearType -eq "comments") { $range.ClearComments() }
        else { $range.Clear() }
        Output-Json @{ success = $true; data = @{ range = $p.range; clearType = $clearType } }
    }

    "insertRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $startRow = if ($p.row) { [int]$p.row } elseif ($p.startRow) { [int]$p.startRow } else { $null }
        if ($null -eq $startRow) { Output-Json @{ success = $false; error = "row/startRow required" }; exit }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        $endRow = $startRow + $count - 1
        $sheet.Range("${startRow}:${endRow}").Insert()
        Output-Json @{ success = $true; data = @{ insertedAt = $startRow; count = $count } }
    }

    "insertColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $col = if ($p.column) { $p.column } elseif ($p.startColumn) { $p.startColumn } else { $null }
        if ($null -eq $col) { Output-Json @{ success = $false; error = "column/startColumn required" }; exit }
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        for ($i = 0; $i -lt $count; $i++) {
            $sheet.Range("${col}:${col}").Insert()
        }
        Output-Json @{ success = $true; data = @{ insertedAt = $col; count = $count } }
    }

    "deleteRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $startRow = if ($p.row) { [int]$p.row } elseif ($p.startRow) { [int]$p.startRow } else { $null }
        if ($null -eq $startRow) { Output-Json @{ success = $false; error = "row/startRow required" }; exit }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        $endRow = $startRow + $count - 1
        $sheet.Range("${startRow}:${endRow}").Delete()
        Output-Json @{ success = $true; data = @{ deletedFrom = $startRow; count = $count } }
    }

    "deleteColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $col = if ($p.column) { $p.column } elseif ($p.startColumn) { $p.startColumn } else { $null }
        if ($null -eq $col) { Output-Json @{ success = $false; error = "column/startColumn required" }; exit }
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        for ($i = 0; $i -lt $count; $i++) {
            $sheet.Range("${col}:${col}").Delete()
        }
        Output-Json @{ success = $true; data = @{ deletedFrom = $col; count = $count } }
    }

    "hideRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $rows = Get-RowRefList $p
        if ($rows.Count -eq 0) { Output-Json @{ success = $false; error = "row/rows/startRow+endRow required" }; exit }
        # The tool offers hide=false to show rows again.
        $hide = if ($null -ne $p.hide) { [bool]$p.hide } else { $true }
        # Range("2:2").Hidden raises E_FAIL on WPS; the Rows collection works.
        foreach ($r in $rows) {
            $target = if ("$r" -match ":") { "$r" } else { "$($r):$($r)" }
            $sheet.Rows($target).Hidden = $hide
        }
        Output-Json @{ success = $true; data = @{ rows = $rows; hidden = $hide } }
    }

    "hideColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $cols = Get-ColumnRefList $p
        if ($cols.Count -eq 0) { Output-Json @{ success = $false; error = "column/columns/startColumn+endColumn required" }; exit }
        $hide = if ($null -ne $p.hide) { [bool]$p.hide } else { $true }
        $applied = @()
        foreach ($c in $cols) {
            $ref = $c
            if ($ref -is [int]) { $ref = Convert-ColumnNumberToLetter([int]$ref) }
            $target = if ("$ref" -match ":") { "$ref" } else { "$($ref):$($ref)" }
            $sheet.Columns($target).Hidden = $hide
            $applied += $ref
        }
        Output-Json @{ success = $true; data = @{ columns = $applied; hidden = $hide } }
    }

    "showRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $rows = Get-RowRefList $p
        if ($rows.Count -eq 0) { Output-Json @{ success = $false; error = "row/rows/startRow+endRow required" }; exit }
        foreach ($r in $rows) {
            $target = if ("$r" -match ":") { "$r" } else { "$($r):$($r)" }
            $sheet.Rows($target).Hidden = $false
        }
        Output-Json @{ success = $true; data = @{ rows = $rows; hidden = $false } }
    }

    "showColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $cols = Get-ColumnRefList $p
        if ($cols.Count -eq 0) { Output-Json @{ success = $false; error = "column/columns/startColumn+endColumn required" }; exit }
        $applied = @()
        foreach ($c in $cols) {
            $ref = $c
            if ($ref -is [int]) { $ref = Convert-ColumnNumberToLetter([int]$ref) }
            $target = if ("$ref" -match ":") { "$ref" } else { "$($ref):$($ref)" }
            $sheet.Columns($target).Hidden = $false
            $applied += $ref
        }
        Output-Json @{ success = $true; data = @{ columns = $applied; hidden = $false } }
    }

    "groupRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if (-not $p.startRow -or -not $p.endRow) { Output-Json @{ success = $false; error = "startRow/endRow required" }; exit }
        $sheet.Range("$($p.startRow):$($p.endRow)").Group()
        Output-Json @{ success = $true; data = @{ grouped = "$($p.startRow):$($p.endRow)" } }
    }

    "groupColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if (-not $p.startColumn -or -not $p.endColumn) { Output-Json @{ success = $false; error = "startColumn/endColumn required" }; exit }
        $startCol = $p.startColumn
        $endCol = $p.endColumn
        if ($startCol -is [int]) { $startCol = Convert-ColumnNumberToLetter([int]$startCol) }
        if ($endCol -is [int]) { $endCol = Convert-ColumnNumberToLetter([int]$endCol) }
        $sheet.Range("${startCol}:${endCol}").Group()
        Output-Json @{ success = $true; data = @{ grouped = "${startCol}:${endCol}" } }
    }

    "freezePanes" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        # freeze=false is how the tool asks to unfreeze.
        if ($null -ne $p.freeze -and -not [bool]$p.freeze) {
            $excel.ActiveWindow.FreezePanes = $false
            Output-Json @{ success = $true; data = @{ frozen = $false; message = "窗格已取消冻结" } }
            exit
        }
        # "freeze through row N / column N" means the split sits after them, so the cell one past
        # the last frozen row/column is the one to select.
        $cellRef = $null
        if ($p.cell) {
            $cellRef = [string]$p.cell
        } else {
            $splitRow = if ($null -ne $p.row) { [int]$p.row + 1 } else { 1 }
            $splitCol = 1
            if ($null -ne $p.column) { $splitCol = [int]$p.column + 1 }
            $cellRef = (Convert-ColumnNumberToLetter($splitCol)) + [string]$splitRow
        }
        $null = $sheet.Activate()
        $sheet.Range($cellRef).Select()
        $excel.ActiveWindow.FreezePanes = $true
        Output-Json @{ success = $true; data = @{ frozen = $true; cell = $cellRef; message = "窗格已冻结" } }
    }

    "unfreezePanes" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $excel.ActiveWindow.FreezePanes = $false
        Output-Json @{ success = $true; data = @{ message = "窗格冻结已取消" } }
    }

# ==================== Excel 表（ListObject）====================
    # P2-2 的新 COM 代码。WPS 对 ListObject 支持完整（建表/总计行/结构化引用/Resize/Unlist 都实测可用），
    # 桥里此前一条都没有。"table" 既接受表名（表1 / Sales）也接受该表所在的 1 基序号。

    "createListObject" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if (-not $p.range) { Output-Json @{ success = $false; error = "range required" }; exit }
        # xlYes=1 / xlNo=2 / xlGuess=0. Default xlYes: a data block handed to a table tool has a header
        # row in almost every case, and guessing can silently promote the first data row to headers.
        $hasHeaders = 1
        if ($null -ne $p.hasHeaders -and -not [bool]$p.hasHeaders) { $hasHeaders = 2 }
        $source = $sheet.Range([string]$p.range)
        try {
            $lo = $sheet.ListObjects.Add(1, $source, $null, $hasHeaders)
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        if ($p.name) { try { $lo.Name = [string]$p.name } catch { Add-WpsWarning ("list object rename failed: " + $_.Exception.Message) } }
        if ($p.tableStyle) { try { $lo.TableStyle = [string]$p.tableStyle } catch { Add-WpsWarning ("list object style failed: " + $_.Exception.Message) } }
        Output-Json @{ success = $true; data = (Get-ListObjectInfo $lo $sheet.Name ([string]$p.range)) }
    }

    "getListObjects" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $targets = @()
        if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") {
            $targets += $wb.Sheets.Item($p.sheet)
        } else {
            for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $targets += $wb.Sheets.Item($i) }
        }
        $tables = @()
        foreach ($sheet in $targets) {
            $count = 0
            try { $count = [int]$sheet.ListObjects.Count } catch { $count = 0 }
            for ($i = 1; $i -le $count; $i++) {
                try { $tables += (Get-ListObjectInfo $sheet.ListObjects.Item($i) $sheet.Name "") } catch { Add-WpsWarning $_.Exception.Message }
            }
        }
        Output-Json @{ success = $true; data = @{ tables = $tables; count = $tables.Count } }
    }

    "addListRow" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $lo = Get-ListObjectByName $sheet $p.table
        if ($null -eq $lo) { Output-Json @{ success = $false; error = "table not found on this sheet" }; exit }
        try { $null = $lo.ListRows.Add() } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        if ($null -ne $p.values) {
            $vals = @($p.values)
            if ($vals.Count -gt 0) {
                $matrix = New-Object 'object[,]' 1, $vals.Count
                for ($c = 0; $c -lt $vals.Count; $c++) {
                    $v = $vals[$c]
                    if ($null -eq $v) { $matrix[0, $c] = $null }
                    elseif ($v -is [bool]) { $matrix[0, $c] = [bool]$v }
                    elseif ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) { $matrix[0, $c] = [double]$v }
                    else { $matrix[0, $c] = [string]$v }
                }
                $rowCount = [int]$lo.ListRows.Count
                $rowRange = $lo.ListRows.Item($rowCount).Range
                $rowRange.PSObject.Properties['Value2'].Value = $matrix
            }
        }
        # Same stale-geometry behaviour as deleteListRow, so re-resolve before describing the table.
        $lo = Get-ListObjectByName $sheet ([string]$lo.Name)
        Output-Json @{ success = $true; data = (Get-ListObjectInfo $lo $sheet.Name "") }
    }

    "deleteListRow" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $lo = Get-ListObjectByName $sheet $p.table
        if ($null -eq $lo) { Output-Json @{ success = $false; error = "table not found on this sheet" }; exit }
        $total = 0
        try { $total = [int]$lo.ListRows.Count } catch { $total = 0 }
        $idx = [int]$p.rowIndex
        if ($idx -lt 1 -or $idx -gt $total) { Output-Json @{ success = $false; error = ("rowIndex " + $idx + " out of range 1.." + $total) }; exit }
        $name = [string]$lo.Name
        try { $lo.ListRows.Item($idx).Delete() } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        # WPS keeps the pre-delete geometry on the object we already hold (ListRows.Count still reports
        # the old number, Range still the old address) until the ListObject is looked up again, so the
        # result would describe the table as it was. Re-resolve before reporting.
        $lo = Get-ListObjectByName $sheet $name
        Output-Json @{ success = $true; data = (Get-ListObjectInfo $lo $sheet.Name "") }
    }

    "updateListObject" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $lo = Get-ListObjectByName $sheet $p.table
        if ($null -eq $lo) { Output-Json @{ success = $false; error = "table not found on this sheet" }; exit }
        try {
            if ($p.name) { $lo.Name = [string]$p.name }
            if ($p.tableStyle) { $lo.TableStyle = [string]$p.tableStyle }
            if ($null -ne $p.showHeaders) { $lo.ShowHeaders = [bool]$p.showHeaders }
            if ($null -ne $p.showAutoFilter) { $lo.ShowAutoFilter = [bool]$p.showAutoFilter }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = (Get-ListObjectInfo $lo $sheet.Name "") }
    }

    "setListObjectTotals" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $lo = Get-ListObjectByName $sheet $p.table
        if ($null -eq $lo) { Output-Json @{ success = $false; error = "table not found on this sheet" }; exit }
        try {
            if ($null -ne $p.show) { $lo.ShowTotals = [bool]$p.show }
            if ($p.column) {
                $col = $null
                $i = 0
                foreach ($c in @($lo.ListColumns)) {
                    $i = $i + 1
                    if ("$($c.Name)" -eq [string]$p.column -or "$i" -eq [string]$p.column) { $col = $c; break }
                }
                if ($null -eq $col) { Output-Json @{ success = $false; error = "column not found in this table" }; exit }
                if (-not [bool]$lo.ShowTotals) { $lo.ShowTotals = $true }
                # XlTotalsCalculation. Measured against WPS with bare COM (the enum order is not the
                # order of the total-row menu): 1=sum 2=average 3=count 4=countNums 5=max 6=min
                # 7=stdDev 8=var; 9 and 10 are rejected.
                $calcMap = @{ sum = 1; average = 2; count = 3; countNums = 4; max = 5; min = 6; stdDev = 7; var = 8; none = 0 }
                $calc = $calcMap[[string]$p.function]
                if ($null -eq $calc) { $calc = 1 }
                $col.TotalsCalculation = $calc
            }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = (Get-ListObjectInfo $lo $sheet.Name "") }
    }

    "resizeListObject" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $lo = Get-ListObjectByName $sheet $p.table
        if ($null -eq $lo) { Output-Json @{ success = $false; error = "table not found on this sheet" }; exit }
        if (-not $p.range) { Output-Json @{ success = $false; error = "range required" }; exit }
        try { $null = $lo.Resize($sheet.Range([string]$p.range)) } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        Output-Json @{ success = $true; data = (Get-ListObjectInfo $lo $sheet.Name ([string]$p.range)) }
    }

    "unlistListObject" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $lo = Get-ListObjectByName $sheet $p.table
        if ($null -eq $lo) { Output-Json @{ success = $false; error = "table not found on this sheet" }; exit }
        $info = Get-ListObjectInfo $lo $sheet.Name ""
        try { $lo.Unlist() } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        Output-Json @{ success = $true; data = @{
            sheet = $info.sheet; name = $info.name; range = $info.range; columns = $info.columns
            columnCount = $info.columnCount; message = "表已转回普通区域（数据与格式保留）"
        } }
    }

# ==================== Excel 页面设置 / 打印 / 外观 / 公式审计（P2-3）====================
    # Excel 此前完全没有页面设置与打印能力（桥里的 setPageSetup / insertPageBreak 都是 Word 的）。
    # 打印与打印预览刻意不做成工具：前者是物理副作用，后者会开模态窗口把常驻宿主卡住。

    "getSheetSettings" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        Output-Json @{ success = $true; data = (Get-SheetSettingsSnapshot $sheet) }
    }

    "setSheetPageSetup" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $ps = $sheet.PageSetup
        $applied = @()
        try {
            if ($p.orientation) {
                if ("$($p.orientation)".ToLower().StartsWith("land")) { $ps.Orientation = 2 } else { $ps.Orientation = 1 }
                $applied += "orientation"
            }
            if ($p.paperSize) {
                # xlPaperA4=9 A3=8 A5=11 B5=13 letter=1 legal=5 tabloid=3；也接受直接给数字。
                $paperMap = @{ a4 = 9; a3 = 8; a5 = 11; b5 = 13; letter = 1; legal = 5; tabloid = 3 }
                $paper = $paperMap["$($p.paperSize)".ToLower()]
                if ($null -eq $paper -and "$($p.paperSize)" -match '^d+$') { $paper = [int]$p.paperSize }
                if ($null -eq $paper) { Output-Json @{ success = $false; error = ("unknown paperSize: " + $p.paperSize) }; exit }
                $ps.PaperSize = $paper
                $applied += "paperSize"
            }
            if ($null -ne $p.topMargin) { $ps.TopMargin = [double]$p.topMargin; $applied += "topMargin" }
            if ($null -ne $p.bottomMargin) { $ps.BottomMargin = [double]$p.bottomMargin; $applied += "bottomMargin" }
            if ($null -ne $p.leftMargin) { $ps.LeftMargin = [double]$p.leftMargin; $applied += "leftMargin" }
            if ($null -ne $p.rightMargin) { $ps.RightMargin = [double]$p.rightMargin; $applied += "rightMargin" }
            if ($null -ne $p.headerMargin) { $ps.HeaderMargin = [double]$p.headerMargin; $applied += "headerMargin" }
            if ($null -ne $p.footerMargin) { $ps.FooterMargin = [double]$p.footerMargin; $applied += "footerMargin" }
            # Zoom 与 FitToPages 互斥：设置 FitToPages 前必须先把 Zoom 置 False，否则缩放比例优先。
            if ($null -ne $p.fitToPagesWide -or $null -ne $p.fitToPagesTall) {
                $ps.Zoom = $false
                if ($null -ne $p.fitToPagesWide) { $ps.FitToPagesWide = [int]$p.fitToPagesWide }
                if ($null -ne $p.fitToPagesTall) { $ps.FitToPagesTall = [int]$p.fitToPagesTall }
                $applied += "fitToPages"
            }
            if ($null -ne $p.zoom) { $ps.Zoom = [int]$p.zoom; $applied += "zoom" }
            if ($null -ne $p.centerHorizontally) { $ps.CenterHorizontally = [bool]$p.centerHorizontally; $applied += "centerHorizontally" }
            if ($null -ne $p.centerVertically) { $ps.CenterVertically = [bool]$p.centerVertically; $applied += "centerVertically" }
            if ($null -ne $p.printGridlines) { $ps.PrintGridlines = [bool]$p.printGridlines; $applied += "printGridlines" }
            if ($null -ne $p.printHeadings) { $ps.PrintHeadings = [bool]$p.printHeadings; $applied += "printHeadings" }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ applied = $applied; settings = (Get-SheetSettingsSnapshot $sheet) } }
    }

    "setSheetPrintTitles" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $ps = $sheet.PageSetup
        $applied = @()
        try {
            if ($null -ne $p.printTitleRows) { $ps.PrintTitleRows = [string]$p.printTitleRows; $applied += "printTitleRows" }
            if ($null -ne $p.printTitleColumns) { $ps.PrintTitleColumns = [string]$p.printTitleColumns; $applied += "printTitleColumns" }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ applied = $applied; settings = (Get-SheetSettingsSnapshot $sheet) } }
    }

    "setSheetHeaderFooter" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $ps = $sheet.PageSetup
        $applied = @()
        try {
            if ($null -ne $p.leftHeader) { $ps.LeftHeader = [string]$p.leftHeader; $applied += "leftHeader" }
            if ($null -ne $p.centerHeader) { $ps.CenterHeader = [string]$p.centerHeader; $applied += "centerHeader" }
            if ($null -ne $p.rightHeader) { $ps.RightHeader = [string]$p.rightHeader; $applied += "rightHeader" }
            if ($null -ne $p.leftFooter) { $ps.LeftFooter = [string]$p.leftFooter; $applied += "leftFooter" }
            if ($null -ne $p.centerFooter) { $ps.CenterFooter = [string]$p.centerFooter; $applied += "centerFooter" }
            if ($null -ne $p.rightFooter) { $ps.RightFooter = [string]$p.rightFooter; $applied += "rightFooter" }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ applied = $applied; settings = (Get-SheetSettingsSnapshot $sheet) } }
    }

    "setSheetAppearance" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $applied = @()
        try {
            if ($null -ne $p.visible) {
                $value = $p.visible
                $code = $null
                if ($value -is [bool]) {
                    if ($value) { $code = -1 } else { $code = 0 }
                } else {
                    $text = "$value".ToLower()
                    if ($text -eq "veryhidden") { $code = 2 }
                    elseif ($text -eq "hidden" -or $text -eq "false" -or $text -eq "0") { $code = 0 }
                    elseif ($text -eq "visible" -or $text -eq "true" -or $text -eq "-1") { $code = -1 }
                }
                if ($null -eq $code) { Output-Json @{ success = $false; error = ("unknown visible value: " + $value) }; exit }
                $sheet.Visible = $code
                $applied += "visible"
            }
            # 空字符串表示“恢复默认标签色”：xlColorIndexNone = -4142，赋 Color 是做不到这件事的。
            if ($null -ne $p.tabColor) {
                if ("$($p.tabColor)" -eq "") {
                    $sheet.Tab.ColorIndex = -4142
                } else {
                    $color = Convert-HexColorToRgbInt ([string]$p.tabColor)
                    if ($null -eq $color) { Output-Json @{ success = $false; error = ("unrecognized color: " + $p.tabColor) }; exit }
                    $sheet.Tab.Color = $color
                }
                $applied += "tabColor"
            }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ applied = $applied; settings = (Get-SheetSettingsSnapshot $sheet) } }
    }

    "setOutlineLevels" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $applied = @()
        try {
            if ($null -ne $p.rowLevels -or $null -ne $p.columnLevels) {
                $rows = 1
                if ($null -ne $p.rowLevels) { $rows = [int]$p.rowLevels }
                $cols = 1
                if ($null -ne $p.columnLevels) { $cols = [int]$p.columnLevels }
                $null = $sheet.Outline.ShowLevels($rows, $cols)
                $applied += "levels"
            }
            # xlSummaryAbove=0 / xlSummaryBelow=1（默认 below）；xlSummaryOnLeft=-1 / xlSummaryOnRight=0。
            if ($p.summaryRow) {
                if ("$($p.summaryRow)".ToLower() -eq "above") { $sheet.Outline.SummaryRow = 0 } else { $sheet.Outline.SummaryRow = 1 }
                $applied += "summaryRow"
            }
            if ($p.summaryColumn) {
                if ("$($p.summaryColumn)".ToLower() -eq "left") { $sheet.Outline.SummaryColumn = -1 } else { $sheet.Outline.SummaryColumn = 0 }
                $applied += "summaryColumn"
            }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{
            applied = $applied; sheet = $sheet.Name
            summaryRow = (Get-ComValue $sheet.Outline "SummaryRow")
            summaryColumn = (Get-ComValue $sheet.Outline "SummaryColumn")
        } }
    }

    "resetPageBreaks" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        try { $sheet.ResetAllPageBreaks() } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        $count = 0
        try { $count = [int]$sheet.HPageBreaks.Count } catch { $count = 0 }
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; hPageBreaks = $count; message = "手动分页符已清除" } }
    }

    "getFormulaAudit" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if (-not $p.cell) { Output-Json @{ success = $false; error = "cell required" }; exit }
        $cell = $sheet.Range([string]$p.cell)
        try {
            if ($p.clearArrows) { $sheet.ClearArrows() }
            if ($p.showPrecedents) { $cell.ShowPrecedents() }
            if ($p.showDependents) { $cell.ShowDependents() }
        } catch { Add-WpsWarning ("arrow drawing failed: " + $_.Exception.Message) }
        $pre = 0; $dep = 0; $direct = 0
        $preAddr = ""; $depAddr = ""
        try { $pre = [int]$cell.Precedents.Count } catch { $pre = 0 }
        try { $dep = [int]$cell.Dependents.Count } catch { $dep = 0 }
        try { $direct = [int]$cell.DirectPrecedents.Count } catch { $direct = 0 }
        try { $preAddr = Get-RangeAddressSafe $cell.Precedents "" } catch { $preAddr = "" }
        try { $depAddr = Get-RangeAddressSafe $cell.Dependents "" } catch { $depAddr = "" }
        $formula = ""
        try { $formula = [string]$cell.Formula } catch { $formula = "" }
        # 常量格上 Formula 返回的是值本身（不是空串），所以“是不是公式”只能问 HasFormula。
        $hasFormula = $false
        try { $hasFormula = [bool](Get-ComValue $cell "HasFormula") } catch { $hasFormula = $false }
        Output-Json @{ success = $true; data = @{
            sheet = $sheet.Name; cell = [string]$p.cell; formula = $formula; hasFormula = $hasFormula
            precedents = $pre; dependents = $dep; directPrecedents = $direct
            precedentAddress = $preAddr; dependentAddress = $depAddr
        } }
    }

# ==================== Excel 高级项：透视表 / 单变量求解 / 迷你图 / 图表（P2-4）====================
    # 全部先用裸 COM 量过支持面。刻意推迟两项（都写进 FIXES 43）：
    #   切片器：SlicerCaches.Add2 能建出缓存，但 Slicers.Count 仍是 0，用户可见效果不确定；
    #   场景管理器：WPS 把 Worksheet.Scenarios 暴露成一个方法（$s.Scenarios() 才拿到集合），语义含糊。

    "getPivotTables" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $targets = @()
        if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $targets += $wb.Sheets.Item($p.sheet) }
        else { for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $targets += $wb.Sheets.Item($i) } }
        $tables = @()
        foreach ($sheet in $targets) {
            $count = 0
            try { $count = [int]$sheet.PivotTables.Count } catch { $count = 0 }
            for ($i = 1; $i -le $count; $i++) {
                $pt = $null
                try { $pt = $sheet.PivotTables($i) } catch { continue }
                $name = ""; try { $name = [string]$pt.Name } catch { $name = "" }
                $range = ""; try { $range = Get-RangeAddressSafe $pt.TableRange2 "" } catch { $range = "" }
                $rowField = ""; try { if ([int]$pt.RowFields().Count -ge 1) { $rowField = [string]$pt.RowFields(1).Name } } catch { $rowField = "" }
                $dataField = ""; try { if ([int]$pt.DataFields().Count -ge 1) { $dataField = [string]$pt.DataFields(1).Name } } catch { $dataField = "" }
                $tables += @{ sheet = $sheet.Name; index = $i; name = $name; range = $range; rowField = $rowField; dataField = $dataField }
            }
        }
        Output-Json @{ success = $true; data = @{ tables = $tables; count = $tables.Count } }
    }

    "refreshPivotTables" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $targets = @()
        if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $targets += $wb.Sheets.Item($p.sheet) }
        else { for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $targets += $wb.Sheets.Item($i) } }
        $refreshed = 0
        $names = @()
        try {
            foreach ($sheet in $targets) {
                if ($null -ne $p.pivotTable -and "$($p.pivotTable)" -ne "") {
                    $pt = $sheet.PivotTables([string]$p.pivotTable)
                    $pt.RefreshTable() | Out-Null
                    $refreshed = $refreshed + 1
                    $names += [string]$pt.Name
                } else {
                    $count = [int]$sheet.PivotTables.Count
                    for ($i = 1; $i -le $count; $i++) {
                        $pt = $sheet.PivotTables($i)
                        $pt.RefreshTable() | Out-Null
                        $refreshed = $refreshed + 1
                        $names += [string]$pt.Name
                    }
                }
            }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ refreshed = $refreshed; names = $names } }
    }

    "clearPivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if (-not $p.pivotTable) { Output-Json @{ success = $false; error = "pivotTable required" }; exit }
        $pt = $null
        try { $pt = $sheet.PivotTables([string]$p.pivotTable) } catch { $pt = $null }
        if ($null -eq $pt) { Output-Json @{ success = $false; error = "pivot table not found on this sheet" }; exit }
        $name = ""; try { $name = [string]$pt.Name } catch { $name = "" }
        $rangeBefore = ""; try { $rangeBefore = Get-RangeAddressSafe $pt.TableRange2 "" } catch { $rangeBefore = "" }
        try { $pt.TableRange2.Clear() } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        # WPS 清掉报表后 PivotTable 对象仍留在集合里（再读 TableRange2 会 E_FAIL），所以不谎报“已删除”。
        $remaining = 0
        try { $remaining = [int]$sheet.PivotTables.Count } catch { $remaining = 0 }
        Output-Json @{ success = $true; data = @{
            sheet = $sheet.Name; name = $name; rangeBefore = $rangeBefore; remaining = $remaining
            message = "透视表报表已清除；WPS 的 PivotTable 对象会留到保存/重开，期间集合里仍能看到它"
        } }
    }

    "refreshAllData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        try { $wb.RefreshAll() } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        Output-Json @{ success = $true; data = @{ workbook = $wb.Name; message = "已刷新工作簿的全部外部数据与透视表" } }
    }

    "goalSeek" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if (-not $p.cell) { Output-Json @{ success = $false; error = "cell required" }; exit }
        if (-not $p.changingCell) { Output-Json @{ success = $false; error = "changingCell required" }; exit }
        $target = $sheet.Range([string]$p.cell)
        $changing = $sheet.Range([string]$p.changingCell)
        $goal = [double]$p.goal
        $solved = $false
        try { $solved = [bool]$target.GoalSeek($goal, $changing) } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        $achieved = ""
        try { $achieved = $changing.Value2 } catch { $achieved = "" }
        $resultValue = ""
        try { $resultValue = $target.Value2 } catch { $resultValue = "" }
        Output-Json @{ success = $true; data = @{
            sheet = $sheet.Name; cell = [string]$p.cell; goal = $goal; changingCell = [string]$p.changingCell
            solved = $solved; changingValue = $achieved; resultValue = $resultValue
        } }
    }

    "addSparkline" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if (-not $p.dataRange) { Output-Json @{ success = $false; error = "dataRange required" }; exit }
        if (-not $p.location) { Output-Json @{ success = $false; error = "location required" }; exit }
        # xlSparkLine=1 xlSparkColumn=2 xlSparkColumnStacked100=3（实测 1/2/3 都接受）
        $typeMap = @{ line = 1; column = 2; stacked = 3; winloss = 3 }
        $typeName = if ($p.sparklineType) { "$($p.sparklineType)".ToLower() } else { "line" }
        $sparkType = $typeMap[$typeName]
        if ($null -eq $sparkType) { Output-Json @{ success = $false; error = ("unknown sparklineType: " + $p.sparklineType) }; exit }
        $target = $sheet.Range([string]$p.location)
        try { $group = $target.SparklineGroups.Add($sparkType, [string]$p.dataRange) } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        if ($null -ne $p.markers -and [bool]$p.markers) {
            try { $group.Points.Markers.Visible = $true } catch { Add-WpsWarning ("markers failed: " + $_.Exception.Message) }
        }
        $count = 0
        try { $count = [int]$target.SparklineGroups.Count } catch { $count = 0 }
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; location = [string]$p.location; dataRange = [string]$p.dataRange; sparklineType = $typeName; groups = $count } }
    }

    "clearSparkline" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if (-not $p.location) { Output-Json @{ success = $false; error = "location required" }; exit }
        try { $sheet.Range([string]$p.location).SparklineGroups.Clear() } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        $count = 0
        try { $count = [int]$sheet.Range([string]$p.location).SparklineGroups.Count } catch { $count = 0 }
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; location = [string]$p.location; groups = $count } }
    }

    "deleteChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $target = $null
        $label = ""
        if ($null -ne $p.chart -and "$($p.chart)" -ne "") {
            $label = "$($p.chart)"
            if ($label -match '^\d+$') { try { $target = $sheet.ChartObjects([int]$label) } catch { $target = $null } }
            else { try { $target = $sheet.ChartObjects([string]$label) } catch { $target = $null } }
        } else {
            $count = 0
            try { $count = [int]$sheet.ChartObjects().Count } catch { $count = 0 }
            if ($count -eq 1) { try { $target = $sheet.ChartObjects(1); $label = [string]$target.Name } catch { $target = $null } }
        }
        if ($null -eq $target) { Output-Json @{ success = $false; error = "chart not found on this sheet（给 chart 名称，或在只有一张图时省略）" }; exit }
        $name = ""; try { $name = [string]$target.Name } catch { $name = $label }
        try { $target.Delete() } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        $remaining = 0
        try { $remaining = [int]$sheet.ChartObjects().Count } catch { $remaining = 0 }
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; deleted = $name; remaining = $remaining } }
    }

    "setChartLabels" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $target = $null
        if ($null -ne $p.chart -and "$($p.chart)" -ne "") {
            $label = "$($p.chart)"
            if ($label -match '^\d+$') { try { $target = $sheet.ChartObjects([int]$label) } catch { $target = $null } }
            else { try { $target = $sheet.ChartObjects([string]$label) } catch { $target = $null } }
        } else {
            $count = 0
            try { $count = [int]$sheet.ChartObjects().Count } catch { $count = 0 }
            if ($count -eq 1) { try { $target = $sheet.ChartObjects(1) } catch { $target = $null } }
        }
        if ($null -eq $target) { Output-Json @{ success = $false; error = "chart not found on this sheet" }; exit }
        $chart = $target.Chart
        $applied = @()
        try {
            if ($null -ne $p.title) { $chart.HasTitle = $true; $chart.ChartTitle.Text = [string]$p.title; $applied += "title" }
            if ($null -ne $p.categoryAxisTitle) { $chart.Axes(1).HasTitle = $true; $chart.Axes(1).AxisTitle.Text = [string]$p.categoryAxisTitle; $applied += "categoryAxisTitle" }
            if ($null -ne $p.valueAxisTitle) { $chart.Axes(2).HasTitle = $true; $chart.Axes(2).AxisTitle.Text = [string]$p.valueAxisTitle; $applied += "valueAxisTitle" }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        $title = ""; try { if ([bool]$chart.HasTitle) { $title = [string]$chart.ChartTitle.Text } } catch { $title = "" }
        $cat = ""; try { if ([bool]$chart.Axes(1).HasTitle) { $cat = [string]$chart.Axes(1).AxisTitle.Text } } catch { $cat = "" }
        $val = ""; try { if ([bool]$chart.Axes(2).HasTitle) { $val = [string]$chart.Axes(2).AxisTitle.Text } } catch { $val = "" }
        Output-Json @{ success = $true; data = @{
            sheet = $sheet.Name; chart = [string]$target.Name; applied = $applied
            title = $title; categoryAxisTitle = $cat; valueAxisTitle = $val
        } }
    }

# ==================== Word 表格读写编辑（P3-2，新 COM 代码）====================
    # WPS 的 Word 表格对象模型可用（实测：Cell 读写/行列增删/合并拆分/边框底纹/整表文本）。

    "getDocumentTables" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $tables = @()
        for ($i = 1; $i -le $doc.Tables.Count; $i++) {
            $t = $doc.Tables.Item($i)
            $text = ""
            try { $text = ([string]$t.Range.Text).Trim() } catch { $text = "" }
            if ($text.Length -gt 200) { $text = $text.Substring(0, 200) + "..." }
            $style = ""
            try { $style = [string]$t.Style.NameLocal } catch { $style = "" }
            $tables += @{ index = $i; rows = [int]$t.Rows.Count; columns = [int]$t.Columns.Count; style = $style; text = $text }
        }
        Output-Json @{ success = $true; data = @{ tables = $tables; count = $tables.Count } }
    }

    "getTableData" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $index = if ($null -ne $p.table) { [int]$p.table } else { 1 }
        if ($index -lt 1 -or $index -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "table index out of range" }; exit }
        $t = $doc.Tables.Item($index)
        $rows = [int]$t.Rows.Count
        $cols = [int]$t.Columns.Count
        $data = @()
        for ($r = 1; $r -le $rows; $r++) {
            $line = @()
            for ($c = 1; $c -le $cols; $c++) {
                # 合并单元格的非起点格子会抛，如实留空而不是让整次读取失败。
                $cellText = ""
                try {
                    $raw = [string]$t.Cell($r, $c).Range.Text
                    $cellText = $raw.TrimEnd([char[]]@([char]13, [char]7, [char]10))
                } catch { $cellText = "" }
                $line += $cellText
            }
            $data += ,@($line)
        }
        Output-Json @{ success = $true; data = @{ table = $index; rows = $rows; columns = $cols; data = $data } }
    }

    "setTableCell" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $index = if ($null -ne $p.table) { [int]$p.table } else { 1 }
        if ($index -lt 1 -or $index -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "table index out of range" }; exit }
        if ($null -eq $p.row -or $null -eq $p.column) { Output-Json @{ success = $false; error = "row and column required" }; exit }
        $t = $doc.Tables.Item($index)
        $cell = $null
        try { $cell = $t.Cell([int]$p.row, [int]$p.column) } catch { Output-Json @{ success = $false; error = "cell not found (merged or out of range?)" }; exit }
        $value = if ($null -eq $p.text) { "" } else { [string]$p.text }
        try { $cell.Range.Text = $value } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        $readBack = ""
        try { $readBack = ([string]$cell.Range.Text).TrimEnd([char[]]@([char]13, [char]7, [char]10)) } catch { $readBack = "" }
        Output-Json @{ success = $true; data = @{ table = $index; row = [int]$p.row; column = [int]$p.column; text = $readBack } }
    }

    "addTableLines" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $index = if ($null -ne $p.table) { [int]$p.table } else { 1 }
        if ($index -lt 1 -or $index -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "table index out of range" }; exit }
        $kind = if ($p.kind) { "$($p.kind)".ToLower() } else { "row" }
        if ($kind -ne "row" -and $kind -ne "column") { Output-Json @{ success = $false; error = "kind must be row or column" }; exit }
        $count = if ($null -ne $p.count) { [int]$p.count } else { 1 }
        if ($count -lt 1) { Output-Json @{ success = $false; error = "count must be >= 1" }; exit }
        $t = $doc.Tables.Item($index)
        try {
            for ($i = 1; $i -le $count; $i++) {
                if ($kind -eq "row") {
                    # beforeRow 给了就插在它前面，否则追加到表尾。
                    if ($null -ne $p.position) { $t.Rows.Add($t.Rows.Item([int]$p.position)) | Out-Null }
                    else { $t.Rows.Add() | Out-Null }
                } else {
                    if ($null -ne $p.position) { $t.Columns.Add($t.Columns.Item([int]$p.position)) | Out-Null }
                    else { $t.Columns.Add() | Out-Null }
                }
            }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ table = $index; kind = $kind; added = $count; rows = [int]$t.Rows.Count; columns = [int]$t.Columns.Count } }
    }

    "deleteTableLine" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $index = if ($null -ne $p.table) { [int]$p.table } else { 1 }
        if ($index -lt 1 -or $index -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "table index out of range" }; exit }
        $kind = if ($p.kind) { "$($p.kind)".ToLower() } else { "row" }
        $lineIndex = [int]$p.lineIndex
        $t = $doc.Tables.Item($index)
        $count = if ($kind -eq "column") { [int]$t.Columns.Count } else { [int]$t.Rows.Count }
        if ($lineIndex -lt 1 -or $lineIndex -gt $count) { Output-Json @{ success = $false; error = ("lineIndex " + $lineIndex + " out of range 1.." + $count) }; exit }
        try {
            if ($kind -eq "column") { $t.Columns.Item($lineIndex).Delete() } else { $t.Rows.Item($lineIndex).Delete() }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ table = $index; kind = $kind; rows = [int]$t.Rows.Count; columns = [int]$t.Columns.Count } }
    }

    "mergeTableCells" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $index = if ($null -ne $p.table) { [int]$p.table } else { 1 }
        if ($index -lt 1 -or $index -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "table index out of range" }; exit }
        # 逐个显式检查：ConvertFrom-Json 得到的是 PSObject，用变量拼属性名在这里不可靠（生成器也会把它判成动态键）。
        if ($null -eq $p.startRow -or $null -eq $p.startColumn -or $null -eq $p.endRow -or $null -eq $p.endColumn) {
            Output-Json @{ success = $false; error = "startRow/startColumn/endRow/endColumn required" }; exit
        }
        $t = $doc.Tables.Item($index)
        # 合并区域必须是一块连续矩形，用起点与终点的 Cell 直接 Merge。
        try {
            $t.Cell([int]$p.startRow, [int]$p.startColumn).Merge($t.Cell([int]$p.endRow, [int]$p.endColumn))
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ table = $index; rows = [int]$t.Rows.Count; columns = [int]$t.Columns.Count; message = "已合并单元格" } }
    }

    "splitTableCell" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $index = if ($null -ne $p.table) { [int]$p.table } else { 1 }
        if ($index -lt 1 -or $index -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "table index out of range" }; exit }
        $t = $doc.Tables.Item($index)
        $rows = if ($null -ne $p.rows) { [int]$p.rows } else { 1 }
        $cols = if ($null -ne $p.columns) { [int]$p.columns } else { 2 }
        try {
            $t.Cell([int]$p.row, [int]$p.column).Split($rows, $cols)
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ table = $index; rows = [int]$t.Rows.Count; columns = [int]$t.Columns.Count; message = "已拆分单元格" } }
    }

    "setTableFormat" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $index = if ($null -ne $p.table) { [int]$p.table } else { 1 }
        if ($index -lt 1 -or $index -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "table index out of range" }; exit }
        $t = $doc.Tables.Item($index)
        $applied = @()
        try {
            if ($p.style) { $t.Style = [string]$p.style; $applied += "style" }
            if ($null -ne $p.borders) { $t.Borders.Enable = [bool]$p.borders; $applied += "borders" }
            if ($p.autoFit) {
                # wdAutoFitContent=1 / wdAutoFitWindow=2
                if ("$($p.autoFit)".ToLower() -eq "window") { $t.AutoFitBehavior(2) } else { $t.AutoFitBehavior(1) }
                $applied += "autoFit"
            }
            if ($null -ne $p.headerShading) {
                $color = Convert-HexColorToRgbInt ([string]$p.headerShading)
                if ($null -eq $color) { Output-Json @{ success = $false; error = ("unrecognized color: " + $p.headerShading) }; exit }
                # Word 的 Shading.BackgroundPatternColor 用 BGR long，和 Excel 一样。
                for ($c = 1; $c -le [int]$t.Columns.Count; $c++) { $t.Cell(1, $c).Shading.BackgroundPatternColor = $color }
                $applied += "headerShading"
            }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        $style = ""
        try { $style = [string]$t.Style.NameLocal } catch { $style = "" }
        Output-Json @{ success = $true; data = @{ table = $index; applied = $applied; style = $style; rows = [int]$t.Rows.Count; columns = [int]$t.Columns.Count } }
    }

    "convertTableToText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $index = if ($null -ne $p.table) { [int]$p.table } else { 1 }
        if ($index -lt 1 -or $index -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "table index out of range" }; exit }
        # ConvertToText 的 Separator 要一个字符串（或分隔常量），不是字符数组。
        $sep = $null
        if ($p.separator) {
            $name = "$($p.separator)".ToLower()
            if ($name -eq "tab") { $sep = [string][char]9 }
            elseif ($name -eq "comma") { $sep = "," }
            elseif ($name -eq "paragraph") { $sep = [string][char]13 }
            else { $sep = [string]$p.separator }
        }
        $t = $doc.Tables.Item($index)
        try {
            if ($null -eq $sep) { $t.ConvertToText() | Out-Null } else { $t.ConvertToText($sep, $true) | Out-Null }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        Output-Json @{ success = $true; data = @{ tablesRemaining = [int]$doc.Tables.Count } }
    }

# ==================== Word 文档生产族（P3-3，新 COM 代码）====================
    # 页码 / 分栏 / 修订（列表、接受、拒绝）/ 批注删除。全部先经裸 COM 量过。

    "insertPageNumbers" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $sectionIndex = if ($null -ne $p.section) { [int]$p.section } else { 1 }
        if ($sectionIndex -lt 1 -or $sectionIndex -gt $doc.Sections.Count) { Output-Json @{ success = $false; error = "section out of range" }; exit }
        $section = $doc.Sections.Item($sectionIndex)
        # wdAlignPageNumberLeft=0 / Center=1 / Right=2
        $alignment = 2
        $alignmentName = "right"
        if ($p.alignment) {
            $name = "$($p.alignment)".ToLower()
            if ($name -eq "left") { $alignment = 0; $alignmentName = "left" }
            elseif ($name -eq "center") { $alignment = 1; $alignmentName = "center" }
        }
        $firstPage = $true
        if ($null -ne $p.showFirstPage) { $firstPage = [bool]$p.showFirstPage }
        $position = "footer"
        $container = $section.Footers.Item(1)
        if ($p.position -and "$($p.position)".ToLower() -eq "header") { $position = "header"; $container = $section.Headers.Item(1) }
        try { $container.PageNumbers.Add($alignment, $firstPage) | Out-Null } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        $count = 0
        try { $count = [int]$container.PageNumbers.Count } catch { $count = 0 }
        Output-Json @{ success = $true; data = @{ section = $sectionIndex; position = $position; alignment = $alignmentName; showFirstPage = $firstPage; pageNumbers = $count } }
    }

    "setColumns" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $sectionIndex = if ($null -ne $p.section) { [int]$p.section } else { 1 }
        if ($sectionIndex -lt 1 -or $sectionIndex -gt $doc.Sections.Count) { Output-Json @{ success = $false; error = "section out of range" }; exit }
        $count = if ($null -ne $p.count) { [int]$p.count } else { 1 }
        if ($count -lt 1 -or $count -gt 12) { Output-Json @{ success = $false; error = "count must be between 1 and 12" }; exit }
        $ps = $doc.Sections.Item($sectionIndex).PageSetup
        $applied = @()
        try {
            $ps.TextColumns.SetCount($count)
            $applied += "count"
            if ($null -ne $p.spacing) { $ps.TextColumns.Spacing = [double]$p.spacing; $applied += "spacing" }
            if ($null -ne $p.lineBetween) { $ps.TextColumns.LineBetween = [bool]$p.lineBetween; $applied += "lineBetween" }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        $actual = 0
        try { $actual = [int]$ps.TextColumns.Count } catch { $actual = 0 }
        $spacing = 0
        try { $spacing = [double]$ps.TextColumns.Spacing } catch { $spacing = 0 }
        Output-Json @{ success = $true; data = @{ section = $sectionIndex; applied = $applied; count = $actual; spacing = $spacing } }
    }

    "getRevisions" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $tracking = $false
        try { $tracking = [bool]$doc.TrackRevisions } catch { $tracking = $false }
        $revisions = @()
        $total = 0
        try { $total = [int]$doc.Revisions.Count } catch { $total = 0 }
        for ($i = 1; $i -le $total; $i++) {
            $r = $doc.Revisions.Item($i)
            $text = ""
            try { $text = ([string]$r.Range.Text).Trim() } catch { $text = "" }
            if ($text.Length -gt 120) { $text = $text.Substring(0, 120) + "..." }
            $author = ""
            try { $author = [string]$r.Author } catch { $author = "" }
            $revisions += @{ index = $i; type = [int]$r.Type; text = $text; author = $author }
        }
        Output-Json @{ success = $true; data = @{ revisions = $revisions; count = $revisions.Count; trackChanges = $tracking } }
    }

    "acceptRevisions" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $applied = 0
        try {
            if ($null -ne $p.index) {
                $doc.Revisions.Item([int]$p.index).Accept()
                $applied = 1
            } else {
                $before = 0
                try { $before = [int]$doc.Revisions.Count } catch { $before = 0 }
                $doc.AcceptAllRevisions()
                $applied = $before
            }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        $remaining = 0
        try { $remaining = [int]$doc.Revisions.Count } catch { $remaining = 0 }
        Output-Json @{ success = $true; data = @{ accepted = $applied; remaining = $remaining } }
    }

    "rejectRevisions" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $applied = 0
        try {
            if ($null -ne $p.index) {
                $doc.Revisions.Item([int]$p.index).Reject()
                $applied = 1
            } else {
                $before = 0
                try { $before = [int]$doc.Revisions.Count } catch { $before = 0 }
                $doc.RejectAllRevisions()
                $applied = $before
            }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        $remaining = 0
        try { $remaining = [int]$doc.Revisions.Count } catch { $remaining = 0 }
        Output-Json @{ success = $true; data = @{ rejected = $applied; remaining = $remaining } }
    }

    "deleteComment" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $deleted = 0
        try {
            if ($null -ne $p.index) {
                $doc.Comments.Item([int]$p.index).Delete()
                $deleted = 1
            } else {
                # 从后往前删，避免删掉一个之后后面的序号整体前移。
                $count = 0
                try { $count = [int]$doc.Comments.Count } catch { $count = 0 }
                for ($i = $count; $i -ge 1; $i--) { $doc.Comments.Item($i).Delete() }
                $deleted = $count
            }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        $remaining = 0
        try { $remaining = [int]$doc.Comments.Count } catch { $remaining = 0 }
        Output-Json @{ success = $true; data = @{ deleted = $deleted; remaining = $remaining } }
    }

# ==================== Word 长尾（P3-4，新 COM 代码）====================
    # 内容控件 / 脚注尾注 / 索引 / 交叉引用 / 邮件合并。支持面已用裸 COM 量过（word-probe4/5）。

    "getContentControls" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $controls = @()
        $total = 0
        try { $total = [int]$doc.ContentControls.Count } catch { $total = 0 }
        for ($i = 1; $i -le $total; $i++) {
            $cc = $doc.ContentControls.Item($i)
            $title = ""; try { $title = [string]$cc.Title } catch { $title = "" }
            $tag = ""; try { $tag = [string]$cc.Tag } catch { $tag = "" }
            $text = ""; try { $text = ([string]$cc.Range.Text).Trim() } catch { $text = "" }
            if ($text.Length -gt 120) { $text = $text.Substring(0, 120) + "..." }
            $controls += @{ index = $i; type = [int]$cc.Type; title = $title; tag = $tag; text = $text }
        }
        Output-Json @{ success = $true; data = @{ controls = $controls; count = $controls.Count } }
    }

    "addContentControl" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        # wdContentControlRichText=0 plainText=1 checkBox=2 comboBox=3 dropDownList=4 datePicker=5 picture=7
        $typeMap = @{ richText = 0; plainText = 1; checkBox = 2; comboBox = 3; dropDownList = 4; datePicker = 5; picture = 7 }
        $typeName = if ($p.type) { "$($p.type)" } else { "richText" }
        $ccType = $typeMap[$typeName]
        if ($null -eq $ccType) { Output-Json @{ success = $false; error = ("unknown content control type: " + $p.type) }; exit }
        $range = Get-MainTextRange $word $doc
        $cc = $null
        try { $cc = $doc.ContentControls.Add($ccType, $range) } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        if ($null -ne $p.text -and "$($p.text)" -ne "") {
            try { $cc.Range.Text = [string]$p.text } catch { Add-WpsWarning ("content control text failed: " + $_.Exception.Message) }
        }
        if ($p.title) { try { $cc.Title = [string]$p.title } catch { Add-WpsWarning ("content control title failed: " + $_.Exception.Message) } }
        if ($p.tag) { try { $cc.Tag = [string]$p.tag } catch { Add-WpsWarning ("content control tag failed: " + $_.Exception.Message) } }
        $readBack = ""
        try { $readBack = [string]$cc.Range.Text } catch { $readBack = "" }
        $title = ""; try { $title = [string]$cc.Title } catch { $title = "" }
        # 返回数字类型：工具层用它映射中文名；只给字符串名字会让映射失效。
        Output-Json @{ success = $true; data = @{ type = [int]$ccType; typeName = $typeName; title = $title; text = $readBack; total = [int]$doc.ContentControls.Count } }
    }

    "addFootnote" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        if (-not $p.text) { Output-Json @{ success = $false; error = "text required" }; exit }
        $range = Get-MainTextRange $word $doc
        try { $doc.Footnotes.Add($range, [string]$p.text) | Out-Null } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        Output-Json @{ success = $true; data = @{ text = [string]$p.text; total = [int]$doc.Footnotes.Count } }
    }

    "addEndnote" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        if (-not $p.text) { Output-Json @{ success = $false; error = "text required" }; exit }
        $range = Get-MainTextRange $word $doc
        try { $doc.Endnotes.Add($range, [string]$p.text) | Out-Null } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        Output-Json @{ success = $true; data = @{ text = [string]$p.text; total = [int]$doc.Endnotes.Count } }
    }

    "getNotes" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $footnotes = @()
        $footnoteCount = 0
        try { $footnoteCount = [int]$doc.Footnotes.Count } catch { $footnoteCount = 0 }
        for ($i = 1; $i -le $footnoteCount; $i++) {
            # 脚注正文不在 .Range.Text 里（那里是空的），要读 .Reference.Text —— 裸 COM 量出来的差异。
            $text = ""
            try { $text = ([string]$doc.Footnotes.Item($i).Reference.Text).Trim() } catch { $text = "" }
            $footnotes += @{ index = $i; text = $text }
        }
        $endnotes = @()
        $endnoteCount = 0
        try { $endnoteCount = [int]$doc.Endnotes.Count } catch { $endnoteCount = 0 }
        for ($i = 1; $i -le $endnoteCount; $i++) {
            $text = ""
            try { $text = ([string]$doc.Endnotes.Item($i).Reference.Text).Trim() } catch { $text = "" }
            $endnotes += @{ index = $i; text = $text }
        }
        Output-Json @{ success = $true; data = @{ footnotes = $footnotes; endnotes = $endnotes; footnoteCount = $footnoteCount; endnoteCount = $endnoteCount } }
    }

    "insertIndex" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $target = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        try { $doc.Indexes.Add($target, $null, $true) | Out-Null } catch { Output-Json @{ success = $false; error = $_.Exception.Message }; exit }
        $count = 0
        try { $count = [int]$doc.Indexes.Count } catch { $count = 0 }
        Output-Json @{ success = $true; data = @{ indexes = $count; message = "已在文档末尾插入索引（没有索引项时 Word 会写「未找到索引项」）" } }
    }

    "insertCrossReference" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        # 参数直接对应 Word 的 Range.InsertCrossReference(ReferenceType, ReferenceKind, ReferenceItem)：
        # 这两个枚举在 WPS 上没有可靠的友好映射，所以照实透传，不做猜测性翻译。
        $referenceType = if ($null -ne $p.referenceType) { [int]$p.referenceType } else { 1 }
        $referenceKind = if ($null -ne $p.referenceKind) { [int]$p.referenceKind } else { -1 }
        if ($null -eq $p.referenceItem) { Output-Json @{ success = $false; error = "referenceItem required" }; exit }
        $range = $word.Selection.Range
        $item = $p.referenceItem
        try {
            if ("$item" -match '^\d+$') { $range.InsertCrossReference($referenceType, $referenceKind, [int]$item) | Out-Null }
            else { $range.InsertCrossReference($referenceType, $referenceKind, [string]$item) | Out-Null }
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        Output-Json @{ success = $true; data = @{ referenceType = $referenceType; referenceKind = $referenceKind; referenceItem = "$item" } }
    }

    "mailMerge" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        if (-not $p.dataFile) { Output-Json @{ success = $false; error = "dataFile required" }; exit }
        $path = Resolve-InputFilePath $p.dataFile
        if ($null -eq $path) { Output-Json @{ success = $false; error = ("data file not found: " + $p.dataFile) }; exit }
        $fields = @()
        if ($null -ne $p.fields) { $fields = @($p.fields) }
        $original = [string]$doc.Name
        $inserted = 0
        try {
            $doc.MailMerge.MainDocumentType = 0
            $doc.MailMerge.OpenDataSource($path)
            $target = Get-MainTextRange $word $doc
            foreach ($field in $fields) {
                $doc.MailMerge.Fields.Add($target, [string]$field) | Out-Null
                $inserted = $inserted + 1
            }
            $doc.MailMerge.Destination = 0
            $doc.MailMerge.Execute()
        } catch {
            Output-Json @{ success = $false; error = $_.Exception.Message }; exit
        }
        $merged = $word.ActiveDocument
        $name = ""
        try { $name = [string]$merged.Name } catch { $name = "" }
        $preview = ""
        try { $preview = ([string]$merged.Content.Text).Trim() } catch { $preview = "" }
        if ($preview.Length -gt 200) { $preview = $preview.Substring(0, 200) + "..." }
        Output-Json @{ success = $true; data = @{
            mergedDocument = $name; sourceDocument = $original; fields = $fields; fieldsInserted = $inserted; preview = $preview
            message = "已按数据源合并到一个新文档（原文档未改动）"
        } }
    }

# P4-2：把 setShapeShadow / setShapeGradient / setShapeBorder / setShapeTransparency 四个碎片 setter
    # 合并成一个结构化工具（见 docs/FIXES.md 48）。目标解析与四个碎片完全一致。
    "setShapeEffect" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $applied = @()
        if ($null -ne $p.shadowEnabled) { $shape.Shadow.Visible = $(if ([bool]$p.shadowEnabled) { -1 } else { 0 }); $applied += "shadowEnabled" }
        if ($p.shadowColor) {
            $shadowColor = Convert-HexColorToRgbInt([string]$p.shadowColor)
            if ($null -ne $shadowColor) { $shape.Shadow.ForeColor.RGB = $shadowColor; $applied += "shadowColor" }
        }
        if ($null -ne $p.shadowTransparency) { $shape.Shadow.Transparency = [double]$p.shadowTransparency; $applied += "shadowTransparency" }
        if ($null -ne $p.shadowBlur) { $shape.Shadow.Blur = [double]$p.shadowBlur; $applied += "shadowBlur" }
        if ($null -ne $p.shadowOffsetX) { $shape.Shadow.OffsetX = [double]$p.shadowOffsetX; $applied += "shadowOffsetX" }
        if ($null -ne $p.shadowOffsetY) { $shape.Shadow.OffsetY = [double]$p.shadowOffsetY; $applied += "shadowOffsetY" }
        if ($null -ne $p.borderEnabled) { $shape.Line.Visible = $(if ([bool]$p.borderEnabled) { -1 } else { 0 }); $applied += "borderEnabled" }
        if ($p.borderColor) {
            $borderColor = Convert-HexColorToRgbInt([string]$p.borderColor)
            if ($null -ne $borderColor) { $shape.Line.ForeColor.RGB = $borderColor; $applied += "borderColor" }
        }
        if ($null -ne $p.borderWidth) { $shape.Line.Weight = [single]$p.borderWidth; $applied += "borderWidth" }
        if ($p.borderStyle) {
            $dashMap = @{ solid = 1; dash = 4; dot = 3; dash_dot = 5; dash_dot_dot = 6 }
            $dash = $dashMap[[string]$p.borderStyle]
            if ($null -eq $dash) { Output-Json @{ success = $false; error = ("unknown borderStyle '" + $p.borderStyle + "'; use solid/dash/dot/dash_dot/dash_dot_dot") }; exit }
            $shape.Line.DashStyle = $dash
            $applied += "borderStyle"
        }
        if ($p.gradientColor1 -or $p.gradientColor2) {
            $shape.Fill.TwoColorGradient(1, 1)
            if ($p.gradientColor1) {
                $c1 = Convert-HexColorToRgbInt([string]$p.gradientColor1)
                if ($null -ne $c1) { $shape.Fill.ForeColor.RGB = $c1; $applied += "gradientColor1" }
            }
            if ($p.gradientColor2) {
                $c2 = Convert-HexColorToRgbInt([string]$p.gradientColor2)
                if ($null -ne $c2) { $shape.Fill.BackColor.RGB = $c2; $applied += "gradientColor2" }
            }
        }
        if ($null -ne $p.transparency) { $shape.Fill.Transparency = [double]$p.transparency; $applied += "transparency" }
        # 一个效果都没给就是空操作——直接报错，而不是回一个“成功”却什么都没做。
        if ($applied.Count -eq 0) { Output-Json @{ success = $false; error = "nothing to apply: give at least one effect property" }; exit }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; applied = $applied } }
    }

# P4 第二波：三组合并（FIXES 49）。
    #   addAnimation + addAnimationPreset + addEmphasisAnimation -> setAnimation（preset / effectKind 分派）
    #   setPptTableStyle + setPptTableCellStyle + setPptTableRowStyle -> setPptTableFormat（row/col 定作用域）
    #   setSlideNumber + setPptFooter + setPptDateTime -> setSlideFooter（一次设完页脚三件套）

    "setAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        if (-not (Test-WpsSlideIndex $pres $slideIndex)) { Output-Json @{ success = $false; error = "slideIndex is out of range" }; exit }
        $slide = $pres.Slides.Item($slideIndex)
        $timeline = $slide.TimeLine
        # 预设模式：给 preset 就是"整页（或指定形状）按顺序出场"，不指定形状则作用于每个形状。
        if ($p.preset) {
            $presets = @{ fadeIn = @{ effect = 10; duration = 0.5 }; flyIn = @{ effect = 2; duration = 0.5 }; zoomIn = @{ effect = 53; duration = 0.4 }; wipeIn = @{ effect = 22; duration = 0.5 }; appear = @{ effect = 1; duration = 0 } }
            $config = $presets[[string]$p.preset]
            if ($null -eq $config) { Output-Json @{ success = $false; error = ("unknown preset '" + $p.preset + "'; use fadeIn/flyIn/zoomIn/wipeIn/appear") }; exit }
            $delay = 0
            $delayIncrement = if ($p.delayIncrement) { [double]$p.delayIncrement } else { 0.3 }
            $onlyShape = $null
            if ($null -ne $p.shapeName -and "$($p.shapeName)" -ne "") { $onlyShape = $p.shapeName }
            elseif ($null -ne $p.shapeIndex) { $onlyShape = $p.shapeIndex }
            $firstShape = 1
            $lastShape = [int]$slide.Shapes.Count
            if ($null -ne $onlyShape) { $firstShape = [int]$onlyShape; $lastShape = [int]$onlyShape }
            if ($firstShape -lt 1 -or $firstShape -gt $slide.Shapes.Count) { Output-Json @{ success = $false; error = "shapeIndex is out of range" }; exit }
            $animated = 0
            for ($i = $firstShape; $i -le $lastShape; $i++) {
                $shape = $slide.Shapes.Item($i)
                try {
                    $effectObject = $timeline.MainSequence.AddEffect($shape, $config.effect, 0, 1)
                    $effectObject.Timing.Duration = $config.duration
                    $effectObject.Timing.TriggerDelayTime = $delay
                    $delay += $delayIncrement
                    $animated++
                } catch { Add-WpsWarning $_.Exception.Message }
            }
            Output-Json @{ success = $true; data = @{ mode = "preset"; preset = [string]$p.preset; animatedShapes = $animated; shapeIndex = $onlyShape } }
            exit
        }
        if ($null -eq $p.shapeName -and $null -eq $p.shapeIndex) { Output-Json @{ success = $false; error = "shapeIndex/shapeName is required (or give preset to animate the whole slide)" }; exit }
        $shape = $slide.Shapes.Item($(if ($null -ne $p.shapeName) { $p.shapeName } else { $p.shapeIndex }))
        if ("$($p.effectKind)".ToLower() -eq "emphasis") {
            $effectType = if ($p.effect) { "$($p.effect)" } else { "pulse" }
            $effectMap = @{ pulse = 63; spin = 15; grow = 53; teeter = 28 }
            $effectId = if ($effectMap[$effectType]) { $effectMap[$effectType] } else { 63 }
            $effectObject = $timeline.MainSequence.AddEffect($shape, $effectId, 0, 2)
            $effectObject.Timing.Duration = if ($p.duration) { [double]$p.duration } else { 0.5 }
            Output-Json @{ success = $true; data = @{ mode = "emphasis"; shape = $shape.Name; effect = $effectType } }
            exit
        }
        $effect = Get-PptAnimEffect $p.effect
        if ($null -eq $effect) { Output-Json @{ success = $false; error = ("unknown animation '" + $p.effect + "'; use fadeIn/flyIn/wipeIn/zoomIn/bounceIn/spinIn/fadeOut/flyOut or a MsoAnimEffect number") }; exit }
        $triggerMap = @{ onClick = 1; withPrevious = 2; afterPrevious = 3 }
        $trigger = if ($null -ne $p.trigger -and $triggerMap[[string]$p.trigger]) { $triggerMap[[string]$p.trigger] } else { 1 }
        $effectObject = $slide.TimeLine.MainSequence.AddEffect($shape, $effect, 1, $trigger)
        $exitNames = @("fadeout", "flyout")
        $isExit = ($exitNames -contains ([string]$p.effect).ToLower())
        if ($isExit) { try { $effectObject.Exit = -1 } catch { Add-WpsWarning $_.Exception.Message } }
        Output-Json @{ success = $true; data = @{ mode = "entrance"; shape = $shape.Name; effect = $effect; trigger = $trigger; isExit = $isExit } }
    }

    "setPptTableFormat" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { [int]$p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) { $tableCount++; if ($tableCount -eq $targetIndex) { $shape = $s; break } }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; exit }
        $applied = @()
        if ($null -ne $p.left) { $shape.Left = $p.left; $applied += "left" }
        if ($null -ne $p.top) { $shape.Top = $p.top; $applied += "top" }
        if ($null -ne $p.width) { $shape.Width = $p.width; $applied += "width" }
        if ($null -ne $p.height) { $shape.Height = $p.height; $applied += "height" }
        # 作用域：给了 row+col 就是那一格；只给 row 是整行；都没给但有样式键就是整张表。
        $cells = @()
        $scope = "none"
        if ($null -ne $p.row -and $null -ne $p.col) {
            $cells += $shape.Table.Cell([int]$p.row, [int]$p.col).Shape
            $scope = "cell"
        } elseif ($null -ne $p.row) {
            $row = $shape.Table.Rows.Item([int]$p.row)
            for ($c = 1; $c -le $row.Cells.Count; $c++) { $cells += $row.Cells.Item($c).Shape }
            $scope = "row"
        } elseif ($null -ne $p.backgroundColor -or $null -ne $p.fontColor -or $null -ne $p.fontSize -or $null -ne $p.bold) {
            for ($r = 1; $r -le $shape.Table.Rows.Count; $r++) {
                $row = $shape.Table.Rows.Item($r)
                for ($c = 1; $c -le $row.Cells.Count; $c++) { $cells += $row.Cells.Item($c).Shape }
            }
            $scope = "table"
        }
        foreach ($cellShape in $cells) {
            if ($p.backgroundColor) {
                $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
                if ($null -ne $bg) {
                    $cellShape.Fill.Visible = $true
                    $cellShape.Fill.Solid()
                    $cellShape.Fill.ForeColor.RGB = $bg
                    if ($applied -notcontains "backgroundColor") { $applied += "backgroundColor" }
                }
            }
            if ($p.fontColor) {
                $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
                if ($null -ne $fc) { $cellShape.TextFrame.TextRange.Font.Color.RGB = $fc; if ($applied -notcontains "fontColor") { $applied += "fontColor" } }
            }
            if ($p.fontSize) { $cellShape.TextFrame.TextRange.Font.Size = $p.fontSize; if ($applied -notcontains "fontSize") { $applied += "fontSize" } }
            if ($null -ne $p.bold) { $cellShape.TextFrame.TextRange.Font.Bold = $p.bold; if ($applied -notcontains "bold") { $applied += "bold" } }
        }
        if ($applied.Count -eq 0) { Output-Json @{ success = $false; error = "nothing to apply: give geometry or cell style properties" }; exit }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; scope = $scope; applied = $applied } }
    }

    "setSlideFooter" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $headers = $pres.SlideMaster.HeadersFooters
        $applied = @()
        if ($null -ne $p.showSlideNumber) { $headers.SlideNumber.Visible = [bool]$p.showSlideNumber; $applied += "slideNumber" }
        if ($null -ne $p.showFooter -or ($null -ne $p.footerText -and "$($p.footerText)" -ne "")) {
            $visible = if ($null -ne $p.showFooter) { [bool]$p.showFooter } else { $true }
            $headers.Footer.Visible = $visible
            if ($null -ne $p.footerText) { $headers.Footer.Text = [string]$p.footerText }
            $applied += "footer"
        }
        if ($null -ne $p.showDate -or $null -ne $p.dateFormat -or $null -ne $p.autoUpdate) {
            $headers.DateAndTime.Visible = if ($null -ne $p.showDate) { [bool]$p.showDate } else { $true }
            $applied += "dateTime"
        }
        $dateTime = $headers.DateAndTime
        # autoUpdate=true 表示让程序自己刷新日期；false 就把文本钉住。
        $useFixed = $null
        if ($null -ne $p.autoUpdate) { $useFixed = -not [bool]$p.autoUpdate }
        if ($null -ne $p.dateFormat -and "$($p.dateFormat)" -ne "") {
            $pattern = "$($p.dateFormat)".ToUpper().Replace("YYYY", "yyyy").Replace("DD", "dd")
            try {
                $dateTime.UseFormat = $false
                $dateTime.Text = (Get-Date).ToString($pattern)
                $dateTime.Visible = $true
            } catch {
                Output-Json @{ success = $false; error = ("unsupported date format '" + "$($p.dateFormat)" + "': " + $_.Exception.Message) }; exit
            }
        }
        if ($useFixed -eq $true) { $dateTime.UseFormat = $false; $dateTime.Text = "" }
        if ($applied.Count -eq 0) { Output-Json @{ success = $false; error = "nothing to apply: give footer/date/slide-number properties" }; exit }
        Output-Json @{ success = $true; data = @{ applied = $applied; footerText = [string]$headers.Footer.Text; footerVisible = [bool]$headers.Footer.Visible; slideNumberVisible = [bool]$headers.SlideNumber.Visible; dateVisible = [bool]$headers.DateAndTime.Visible } }
    }

    "findInSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $searchRange = if ($p.range) { $sheet.Range([string]$p.range) } else { $sheet.UsedRange }
        # The old Find()+Address() version could never work here: Address() answers only for ranges the
        # resident host reaches directly, and a Find() result is not one of them (measured in P2 wave 1).
        # UsedRange.Address() does work - getExcelContext relies on it - and the cells are then scanned
        # from a single Value2 read, the same accessor getRangeData uses.
        $rangeRef = if ($p.range) { [string]$p.range } else { $sheet.UsedRange.Address() }
        $span = Get-AddressSpan $rangeRef
        if ($null -eq $span) { Output-Json @{ success = $false; error = "findInSheet could not resolve a searchable range" }; exit }
        $rawValue = $searchRange.Value2
        $matrix = $null; $flatArr = @()
        if ($rawValue -is [Array] -and $rawValue.Rank -eq 2) { $matrix = $rawValue }
        elseif ($null -ne $rawValue) { $flatArr = @($rawValue) }
        $mRow = 0; $mCol = 0; $mRows = 1; $mCols = $span.ColumnCount
        if ($null -ne $matrix) {
            $mRow = $matrix.GetLowerBound(0)
            $mCol = $matrix.GetLowerBound(1)
            $mRows = $matrix.GetUpperBound(0) - $mRow + 1
            $mCols = $matrix.GetUpperBound(1) - $mCol + 1
        } elseif ($flatArr.Count -gt 0) {
            $mRows = [Math]::Ceiling($flatArr.Count / $mCols)
        } else {
            $mRows = 0
        }
        $needle = [string]$p.searchText
        $caseSensitive = [bool]$p.matchCase
        $results = @()
        for ($r = 0; $r -lt $span.RowCount; $r++) {
            for ($c = 0; $c -lt $span.ColumnCount; $c++) {
                if ($r -ge $mRows -or $c -ge $mCols) { continue }
                $v = $null
                # The index list must be parenthesised: comma binds tighter than "+" in PowerShell,
                # so "$matrix[$mRow + $r, $mCol + $c]" parses as "$mRow + ($r, $mCol) + $c" and every
                # access throws "Object[] has no op_Addition" - which a bare catch would swallow.
                if ($null -ne $matrix) { $v = $matrix[($mRow + $r), ($mCol + $c)] }
                else { $v = $flatArr[($r * $mCols + $c)] }
                if ($null -eq $v) { continue }
                $cellText = [string]$v
                $hit = if ($caseSensitive) { $cellText.Contains($needle) } else { $cellText.ToLower().Contains($needle.ToLower()) }
                if (-not $hit) { continue }
                $cellAddress = (Convert-ColumnNumberToLetter($span.Column + $c)) + [string]($span.Row + $r)
                $results += @{ address = $cellAddress; value = $v }
            }
        }
        Output-Json @{ success = $true; data = @{ searchText = $needle; results = $results; count = $results.Count } }
    }

    "replaceInSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $searchRange = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $lookAt = if ($p.matchCase) { 1 } else { 2 }
        # Signature is deliberately identical to the findReplaceExcel call site: PowerShell caches a COM
        # member's binder after its first use, so every Range.Replace call must look the same.
        $replaced = $searchRange.Replace([string]$p.searchText, [string]$p.replaceText, [int]$lookAt, 1, $false, $false)
        Output-Json @{ success = $true; data = @{ searchText = $p.searchText; replaceText = $p.replaceText; success = $replaced } }
    }

    "copyRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        # The tool schema is source/destination. The old body only ever copied "range" to the
        # clipboard, so source and destination were dropped and nothing was pasted anywhere.
        $sourceRef = if ($null -ne $p.source) { $p.source } else { $p.range }
        if ($null -eq $sourceRef) { Output-Json @{ success = $false; error = "source/range required" }; exit }
        $sourceRange = $sheet.Range([string]$sourceRef)
        $destRef = if ($null -ne $p.destination) { [string]$p.destination } else { $null }
        if ($null -eq $destRef -or $destRef -eq "") {
            $sourceRange.Copy()
            Output-Json @{ success = $true; data = @{ source = $sourceRef; message = "已复制到剪贴板" } }
            exit
        }
        $sourceRange.Copy($sheet.Range($destRef))
        $excel.CutCopyMode = $false
        Output-Json @{ success = $true; data = @{ source = $sourceRef; destination = $destRef } }
    }

    "pasteRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $destRange = $sheet.Range($p.destination)
        if ($p.pasteType -eq "values") {
            $destRange.PasteSpecial(-4163)
        } elseif ($p.pasteType -eq "formats") {
            $destRange.PasteSpecial(-4122)
        } elseif ($p.pasteType -eq "formulas") {
            $destRange.PasteSpecial(-4123)
        } else {
            $sheet.Paste($destRange)
        }
        Output-Json @{ success = $true; data = @{ destination = $p.destination } }
    }

    "fillSeries" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        # wps_excel_auto_fill sends sourceRange/targetRange and means "extend the source pattern".
        if ($null -ne $p.sourceRange -and $null -ne $p.targetRange) {
            $srcRange = $sheet.Range([string]$p.sourceRange)
            $tgtRange = $sheet.Range([string]$p.targetRange)
            $srcRange.AutoFill($tgtRange, 0)
            Output-Json @{ success = $true; data = @{ sourceRange = $p.sourceRange; targetRange = $p.targetRange } }
            exit
        }
        if ($null -eq $p.range -or "$($p.range)" -eq "") { Output-Json @{ success = $false; error = "range or sourceRange+targetRange is required" }; exit }
        $range = $sheet.Range($p.range)
        $startCell = $range.Cells.Item(1, 1)
        $startValue = if ($null -ne $p.startValue) { $p.startValue } else { 1 }
        Set-ComValue $startCell 'Value2' $startValue
        $typeMap = @{ linear = 0; growth = 1; date = 2; autoFill = 3 }
        $fillType = $typeMap[$p.type]
        if ($null -eq $fillType) { $fillType = 0 }
        $step = if ($null -ne $p.step) { $p.step } else { 1 }
        # The tool schema carries a direction (down/right/up/left). Filling the other way is the
        # same series with a negated step.
        $direction = if ($null -ne $p.direction) { [string]$p.direction } else { $null }
        if ($null -ne $direction -and ($direction -eq "up" -or $direction -eq "left")) {
            $step = -1 * [double]$step
        }
        # Rowcol must stay null so WPS infers the axis from the range shape. Passing 1 (xlRows) for a
        # vertical range silently fills nothing on this build, while 2 and null both work.
        $range.DataSeries($null, -4132, $fillType, $step)
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $p.type; direction = $direction; step = $step } }
    }

    "transpose" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $sourceRef = if ($null -ne $p.sourceRange) { $p.sourceRange } else { $p.source }
        if ($null -eq $sourceRef) { Output-Json @{ success = $false; error = "sourceRange/source required" }; exit }
        $sourceRange = $sheet.Range($sourceRef)
        $destCell = if ($p.destinationCell) { $p.destinationCell } elseif ($p.targetCell) { $p.targetCell } elseif ($p.destination) { $p.destination } else { $null }
        if ($null -eq $destCell) { Output-Json @{ success = $false; error = "destinationCell/targetCell required" }; exit }
        $destRange = $sheet.Range($destCell)
        $sourceRange.Copy()
        $destRange.PasteSpecial(-4163, -4142, $false, $true)
        $excel.CutCopyMode = $false
        Output-Json @{ success = $true; data = @{ source = $sourceRef; destination = $destCell } }
    }

    "textToColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $delimiter = if ($p.delimiter) { $p.delimiter } else { "," }
        $tab = $false; $semicolon = $false; $comma = $false; $space = $false; $other = $false; $otherChar = $null
        if ($delimiter -eq "`t") { $tab = $true }
        elseif ($delimiter -eq ";") { $semicolon = $true }
        elseif ($delimiter -eq ",") { $comma = $true }
        elseif ($delimiter -eq " ") { $space = $true }
        else { $other = $true; $otherChar = $delimiter }
        $range.TextToColumns($null, 1, 1, $false, $tab, $semicolon, $comma, $space, $other, $otherChar)
        Output-Json @{ success = $true; data = @{ range = $p.range; delimiter = $delimiter } }
    }

    "subtotal" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        # Same XlConsolidationFunction constants as consolidate; the old 9/2/1/4/5 failed every call.
        $funcMap = @{ sum = -4157; count = -4112; average = -4106; max = -4136; min = -4139 }
        $func = $funcMap[$p.function]
        if ($null -eq $func) { $func = -4157 }
        $totalCols = $p.totalColumns
        if ($null -eq $totalCols) { $totalCols = $p.totalColumn }
        if ($null -eq $totalCols) { $totalCols = $p.columns }
        if ($totalCols -isnot [System.Array]) { $totalCols = @($totalCols) }
        $replace = if ($null -ne $p.replace) { [bool]$p.replace } else { $true }
        $range.Subtotal([int]$p.groupBy, $func, $totalCols, $replace, $false, $true)
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "createNamedRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $addr = $range.Address()
        $wb.Names.Add($p.name, "=" + $sheet.Name + "!" + $addr)
        Output-Json @{ success = $true; data = @{ name = $p.name; range = $p.range } }
    }

    "deleteNamedRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $wb.Names.Item($p.name).Delete()
        Output-Json @{ success = $true; data = @{ deletedName = $p.name } }
    }

    "getNamedRanges" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $names = @()
        for ($i = 1; $i -le $wb.Names.Count; $i++) {
            $n = $wb.Names.Item($i)
            $names += @{ name = $n.Name; refersTo = $n.RefersTo; visible = $n.Visible }
        }
        Output-Json @{ success = $true; data = @{ names = $names; count = $names.Count } }
    }

    "addCellComment" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        if ($null -eq $p.cell -or "$($p.cell)" -eq "") { Output-Json @{ success = $false; error = "cell is required to add a comment" }; exit }
        # The tool schema calls this parameter comment; accept both spellings.
        $commentText = if ($null -ne $p.comment) { [string]$p.comment } elseif ($null -ne $p.text) { [string]$p.text } else { "" }
        if ($commentText -eq "") { Output-Json @{ success = $false; error = "comment text is required" }; exit }
        $cell = $sheet.Range([string]$p.cell)
        if ($cell.Comment) { $cell.Comment.Delete() }
        $cell.AddComment($commentText)
        if ($p.visible) { $cell.Comment.Visible = $true }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; sheet = $sheet.Name; text = $commentText } }
    }

    "deleteCellComment" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $cell = $sheet.Range($p.cell)
        if ($cell.Comment) { $cell.Comment.Delete() }
        Output-Json @{ success = $true; data = @{ cell = $p.cell } }
    }

    "getCellComments" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        # The tool may narrow the result to a range; without it every comment on the sheet is returned.
        $filter = $null
        if ($null -ne $p.range -and "$($p.range)" -ne "") { $filter = $sheet.Range([string]$p.range) }
        $comments = @()
        for ($i = 1; $i -le $sheet.Comments.Count; $i++) {
            $c = $sheet.Comments.Item($i)
            $addr = $c.Parent.Address()
            $flat = ($addr -replace "\$", "")
            if ($null -ne $filter) {
                $target = $sheet.Range($flat)
                if ($target.Row -lt $filter.Row) { continue }
                if ($target.Row -gt ($filter.Row + $filter.Rows.Count - 1)) { continue }
                if ($target.Column -lt $filter.Column) { continue }
                if ($target.Column -gt ($filter.Column + $filter.Columns.Count - 1)) { continue }
            }
            $comments += @{ cell = $flat; text = $c.Text(); author = if ($c.Author) { $c.Author } else { "" } }
        }
        Output-Json @{ success = $true; data = @{ comments = $comments; count = $comments.Count; range = $p.range } }
    }

    "protectSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = if ($p.sheet) { $excel.ActiveWorkbook.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $password = if ($p.password) { $p.password } else { "" }
        # protect=false is how the tool asks to remove protection.
        $wantProtect = if ($null -ne $p.protect) { [bool]$p.protect } else { $true }
        if ($wantProtect) {
            $sheet.Protect($password, $p.drawingObjects, $p.contents, $p.scenarios)
        } else {
            $sheet.Unprotect($password)
        }
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; protected = $wantProtect } }
    }

    "unprotectSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = if ($p.sheet) { $excel.ActiveWorkbook.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $password = if ($p.password) { $p.password } else { "" }
        $sheet.Unprotect($password)
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; protected = $false } }
    }

    "protectWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $password = if ($p.password) { $p.password } else { "" }
        $wantProtect = if ($null -ne $p.protect) { [bool]$p.protect } else { $true }
        if ($wantProtect) {
            $structure = if ($null -ne $p.structure) { [bool]$p.structure } else { $true }
            $wb.Protect($password, $structure, $p.windows)
        } else {
            $wb.Unprotect($password)
        }
        Output-Json @{ success = $true; data = @{ workbook = $wb.Name; protected = $wantProtect } }
    }

    "insertExcelImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        # Anchoring at a cell is what the tool schema promises; explicit left/top still win.
        $anchorLeft = $null
        $anchorTop = $null
        if ($null -ne $p.cell -and "$($p.cell)" -ne "") {
            $anchor = $sheet.Range([string]$p.cell)
            $anchorLeft = $anchor.Left
            $anchorTop = $anchor.Top
        }
        $left = if ($null -ne $p.left) { $p.left } elseif ($null -ne $anchorLeft) { $anchorLeft } else { 100 }
        $top = if ($null -ne $p.top) { $p.top } elseif ($null -ne $anchorTop) { $anchorTop } else { 100 }
        $width = if ($null -ne $p.width) { $p.width } else { -1 }
        $height = if ($null -ne $p.height) { $p.height } else { -1 }
        if ($null -eq $p.path -or "$($p.path)" -eq "") { Output-Json @{ success = $false; error = "path is required to insert an image" }; exit }
        $imagePath = Resolve-InputFilePath $p.path
        if ($null -eq $imagePath) { Output-Json @{ success = $false; error = ("image file not found: " + "$($p.path)") }; exit }
        $pic = $sheet.Shapes.AddPicture($imagePath, $false, $true, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $pic.Name; path = $imagePath; left = $left; top = $top } }
    }

    "setHyperlink" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.cell)
        # The tool schema spells these url/text; accept both spellings so neither layer is dropped.
        $address = if ($p.address) { $p.address } elseif ($p.url) { $p.url } else { "" }
        $subAddress = if ($p.subAddress) { $p.subAddress } else { "" }
        $screenTip = if ($p.screenTip) { $p.screenTip } else { "" }
        $textToDisplay = if ($p.textToDisplay) { $p.textToDisplay } elseif ($p.text) { $p.text } else { "" }
        $sheet.Hyperlinks.Add($range, $address, $subAddress, $screenTip, $textToDisplay)
        Output-Json @{ success = $true; data = @{ cell = $p.cell; address = $address } }
    }

    "lockCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = Get-WorksheetByParam $excel $p
        $range = $sheet.Range($p.range)
        $locked = if ($null -ne $p.locked) { [bool]$p.locked } else { $true }
        $range.Locked = $locked
        Output-Json @{ success = $true; data = @{ range = $p.range; locked = $locked } }
    }

    "openWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        if (-not $p.path) { Output-Json @{ success = $false; error = "path required" }; exit }
        $prevAlerts = Set-WpsAlertsSuppressed $excel 'excel'
        # AskToUpdateLinks 也是弹框来源，和 UpdateLinks 参数一起双重保险。
        $prevAskLinks = $null
        try { $prevAskLinks = [bool]$excel.AskToUpdateLinks; $excel.AskToUpdateLinks = $false } catch { }
        $wb = $null
        $openError = $null
        try { $wb = Open-ExcelWorkbook $excel $p.path $p.updateLinks ([bool]$p.readOnly) } catch { $openError = $_.Exception.Message }
        try { if ($null -ne $prevAskLinks) { $excel.AskToUpdateLinks = $prevAskLinks } } catch { }
        Restore-WpsAlerts $excel $prevAlerts
        if ($null -ne $openError) { Output-Json @{ success = $false; error = (Format-WpsOpenError $p.path 'excel' $openError) }; exit }
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count } }
    }

    "getOpenWorkbooks" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $workbooks = @()
        for ($i = 1; $i -le $excel.Workbooks.Count; $i++) {
            $wb = $excel.Workbooks.Item($i)
            $workbooks += @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count; active = ($wb.Name -eq $excel.ActiveWorkbook.Name) }
        }
        Output-Json @{ success = $true; data = @{ workbooks = $workbooks; count = $workbooks.Count } }
    }

    "switchWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.Workbooks.Item($(if ($p.name) { $p.name } else { $p.index }))
        $wb.Activate()
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName } }
    }

    "closeWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = if ($p.name) { $excel.Workbooks.Item($p.name) } else { $excel.ActiveWorkbook }
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No workbook" }; exit }
        $name = $wb.Name
        $path = ""
        try { $path = [string]$wb.Path } catch { Add-WpsWarning $_.Exception.Message }
        # Tools spell this parameter "save" while this action used to read only "saveChanges".
        # The mismatch meant save=false was silently ignored and every close asked to save.
        $saveWanted = $true
        if ($null -ne $p.save) { $saveWanted = [bool]$p.save }
        elseif ($null -ne $p.saveChanges) { $saveWanted = [bool]$p.saveChanges }
        $saveChanges = $saveWanted
        $warning = $null
        if ($saveWanted -and $path -eq "") {
            # Close(save) on a workbook that has no file yet opens a modal Save As dialog and
            # blocks the COM call until something answers it. Never allow that.
            $saveChanges = $false
            $warning = "workbook was never saved to disk; closed without saving (use save_as first if you need it on disk)"
        }
        $prevAlerts = $false
        try { $prevAlerts = [bool]$excel.DisplayAlerts } catch { }
        try { $excel.DisplayAlerts = $false } catch { }
        try {
            $wb.Close($saveChanges)
        } catch {
            try { $excel.DisplayAlerts = $prevAlerts } catch { }
            Output-Json @{ success = $false; error = $_.Exception.Message }
            exit
        }
        try { $excel.DisplayAlerts = $prevAlerts } catch { }
        $data = @{ closed = $name; saved = [bool]$saveChanges; saveRequested = $saveWanted }
        if ($null -ne $warning) { $data.warning = $warning }
        Output-Json @{ success = $true; data = $data }
    }

    "createWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.Workbooks.Add()
        if ($p.name) { $wb.SaveAs($p.name) }
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count } }
    }

    # ==================== Word ====================
    "getActiveDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        Output-Json @{ success = $true; data = @{
            name = $doc.Name; path = $doc.FullName
            paragraphCount = $doc.Paragraphs.Count; wordCount = $doc.Words.Count; characterCount = $doc.Characters.Count
            # wdStatisticPages = 2. Not every WPS build answers this; a failed call warns instead of
            # reporting a page count nobody computed.
            pageCount = $(try { [int]$doc.ComputeStatistics(2) } catch { Add-WpsWarning ("page count unavailable: " + $_.Exception.Message); $null })
        }}
    }

    "generateTOC" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $position = if ($p.position) { $p.position } else { "start" }
        $levels = if ($p.levels) { [int]$p.levels } else { 3 }
        $includePageNumbers = if ($null -ne $p.includePageNumbers) { [bool]$p.includePageNumbers } else { $true }
        $range = if ($position -eq "cursor") { $word.Selection.Range } else { $doc.Range(0, 0) }
        $doc.TablesOfContents.Add($range, $true, 1, $levels, $false, "", $true, $true, $includePageNumbers, $true)
        Output-Json @{ success = $true; data = @{ levels = $levels; position = $position; includePageNumbers = $includePageNumbers } }
    }

    "getDocumentText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $text = $doc.Content.Text
        if ($text.Length -gt 10000) { $text = $text.Substring(0, 10000) + "...(truncated)" }
        Output-Json @{ success = $true; data = @{ text = $text; length = $doc.Content.Text.Length } }
    }

    "getOpenDocuments" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $documents = @()
        for ($i = 1; $i -le $word.Documents.Count; $i++) {
            $docItem = $word.Documents.Item($i)
            $documents += @{ name = $docItem.Name; path = $docItem.FullName; paragraphs = $docItem.Paragraphs.Count; active = ($word.ActiveDocument.Name -eq $docItem.Name) }
        }
        Output-Json @{ success = $true; data = @{ documents = $documents; count = $documents.Count } }
    }

    "switchDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $docItem = $word.Documents.Item($(if ($p.name) { $p.name } else { $p.index }))
        $docItem.Activate()
        Output-Json @{ success = $true; data = @{ name = $docItem.Name; path = $docItem.FullName } }
    }

    "openDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        if (-not $p.path) { Output-Json @{ success = $false; error = "path required" }; exit }
        $prevAlerts = Set-WpsAlertsSuppressed $word 'word'
        $docItem = $null
        $openError = $null
        try { $docItem = Open-WordDocument $word $p.path } catch { $openError = $_.Exception.Message }
        Restore-WpsAlerts $word $prevAlerts
        if ($null -ne $openError) { Output-Json @{ success = $false; error = (Format-WpsOpenError $p.path 'word' $openError) }; exit }
        Output-Json @{ success = $true; data = @{ name = $docItem.Name; path = $docItem.FullName; paragraphs = $docItem.Paragraphs.Count } }
    }

    "closeDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $docItem = if ($p.name) { $word.Documents.Item($p.name) } else { $word.ActiveDocument }
        if ($null -eq $docItem) { Output-Json @{ success = $false; error = "No document" }; exit }
        $name = $docItem.Name
        $path = ""
        try { $path = [string]$docItem.Path } catch { Add-WpsWarning $_.Exception.Message }
        # Same "save" vs "saveChanges" mismatch as closeWorkbook.
        $saveWanted = $true
        if ($null -ne $p.save) { $saveWanted = [bool]$p.save }
        elseif ($null -ne $p.saveChanges) { $saveWanted = [bool]$p.saveChanges }
        $saveChanges = $saveWanted
        $warning = $null
        if ($saveWanted -and $path -eq "") {
            $saveChanges = $false
            $warning = "document was never saved to disk; closed without saving (use save_as first if you need it on disk)"
        }
        $prevAlerts = $false
        try { $prevAlerts = [bool]$word.DisplayAlerts } catch { }
        try { $word.DisplayAlerts = 0 } catch { }
        try {
            $docItem.Close($saveChanges)
        } catch {
            try { $word.DisplayAlerts = $prevAlerts } catch { }
            Output-Json @{ success = $false; error = $_.Exception.Message }
            exit
        }
        try { $word.DisplayAlerts = $prevAlerts } catch { }
        $data = @{ closed = $name; saved = [bool]$saveChanges; saveRequested = $saveWanted }
        if ($null -ne $warning) { $data.warning = $warning }
        Output-Json @{ success = $true; data = $data }
    }

    "insertText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $position = if ($p.position) { $p.position } else { "cursor" }
        switch ($position) {
            "start" { $range = $doc.Range(0, 0); $range.InsertBefore($p.text) }
            "end" { $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1); $range.InsertAfter($p.text) }
            default { $word.Selection.TypeText($p.text) }
        }
        if ($p.style) {
            try { $word.Selection.Range.Style = $p.style } catch { Add-WpsWarning $_.Exception.Message }
        }
        Output-Json @{ success = $true; data = @{ position = $position; textLength = $p.text.Length } }
    }

    "setFont" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $range = if ($p.range -eq "all") { $doc.Content } else { $word.Selection.Range }
        if ($p.fontName) { $range.Font.Name = $p.fontName }
        if ($p.fontSize) { $range.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $range.Font.Bold = $p.bold }
        if ($null -ne $p.italic) { $range.Font.Italic = $p.italic }
        if ($null -ne $p.underline) { $range.Font.Underline = $p.underline }
        if ($p.color) {
            $colorMap = @{ red = 255; blue = 16711680; green = 65280; black = 0 }
            $colorValue = $colorMap[$p.color.ToLower()]
            if ($null -eq $colorValue) { $colorValue = Convert-HexColorToRgbInt([string]$p.color) }
            if ($null -ne $colorValue) { $range.Font.Color = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ settings = @{ fontName = $p.fontName; fontSize = $p.fontSize; bold = $p.bold; italic = $p.italic; underline = $p.underline; color = $p.color } } }
    }

    "findReplace" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $findText = if ($null -ne $p.findText) { [string]$p.findText } else { [string]$p.find }
        $hasReplaceText = ($null -ne $p.replaceText) -or ($null -ne $p.replace)
        $replaceText = if ($null -ne $p.replaceText) { [string]$p.replaceText } else { [string]$p.replace }
        $matchCase = if ($null -ne $p.matchCase) { [bool]$p.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $p.matchWholeWord) { [bool]$p.matchWholeWord } else { $false }
        $replaceAll = if ($null -ne $p.replaceAll) { [bool]$p.replaceAll } else { $true }

        # A caller that only wants to search must never erase what it found. The tool layer signals
        # search with replaceMode=false. Without the flag, an empty replacement means "search" too:
        # replacing matches with an empty string deletes them, which is not what a search asks for.
        $isReplace = if ($null -ne $p.replaceMode) { [bool]$p.replaceMode } else { ($hasReplaceText -and $replaceText -ne "") }

        # Count first, so the result reports how much the document actually had to match.
        $count = Get-WordMatchCount $doc $findText $matchCase $matchWholeWord

        if (-not $isReplace) {
            Output-Json @{ success = $true; data = @{ count = $count; replaced = $false; found = $count; find = $findText; replaceMode = $false } }
            exit
        }

        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $findText
        $find.Replacement.Text = $replaceText
        $replaceType = if ($replaceAll) { 2 } else { 1 }
        $result = $find.Execute($findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, $replaceText, $replaceType)
        $replacedCount = if ($replaceAll) { $count } else { if ($result) { 1 } else { 0 } }
        Output-Json @{ success = $true; data = @{ count = $replacedCount; replaced = [bool]$result; found = $count; find = $findText; replace = $replaceText; replaceMode = $true } }
    }

    "insertTable" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $rows = if ($p.rows) { [int]$p.rows } else { 3 }
        $cols = if ($p.cols) { [int]$p.cols } else { 3 }
        $range = $word.Selection.Range
        $table = $doc.Tables.Add($range, $rows, $cols)
        if ($p.data) {
            for ($r = 0; $r -lt [Math]::Min($p.data.Count, $rows); $r++) {
                for ($c = 0; $c -lt [Math]::Min($p.data[$r].Count, $cols); $c++) {
                    $table.Cell($r + 1, $c + 1).Range.Text = [string]$p.data[$r][$c]
                }
            }
        }
        $table.Borders.Enable = $true
        Output-Json @{ success = $true }
    }

    "setParagraph" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $range = if ($p.range -eq "all") { $doc.Content } else { $word.Selection.Range }
        $para = $range.ParagraphFormat
        if ($null -ne $p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2; justify = 3 }
            $align = $alignMap[$p.alignment]
            if ($null -eq $align) { $align = 0 }
            $para.Alignment = $align
        }
        if ($p.lineSpacing) {
            $para.LineSpacingRule = 4
            $para.LineSpacing = [double]$p.lineSpacing * 12
        }
        if ($null -ne $p.spaceBefore) { $para.SpaceBefore = $p.spaceBefore }
        if ($null -ne $p.spaceAfter) { $para.SpaceAfter = $p.spaceAfter }
        if ($null -ne $p.firstLineIndent) { $para.FirstLineIndent = [double]$p.firstLineIndent * 28.35 }
        if ($null -ne $p.leftIndent) { $para.LeftIndent = [double]$p.leftIndent * 28.35 }
        if ($null -ne $p.rightIndent) { $para.RightIndent = [double]$p.rightIndent * 28.35 }
        Output-Json @{ success = $true }
    }

    "setPageSetup" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $ps = $doc.PageSetup
        # The tool schema documents points, so the values are used as points (they used to be
        # multiplied by 28.35 as if they were centimetres, which pushed them out of range).
        if ($null -ne $p.topMargin) {
            $margin = ConvertTo-WordMarginPoints $p.topMargin
            if ($null -eq $margin) { Output-Json @{ success = $false; error = ("topMargin must be 0-1584 points (got " + $p.topMargin + ")") }; exit }
            $ps.TopMargin = $margin
        }
        if ($null -ne $p.bottomMargin) {
            $margin = ConvertTo-WordMarginPoints $p.bottomMargin
            if ($null -eq $margin) { Output-Json @{ success = $false; error = ("bottomMargin must be 0-1584 points (got " + $p.bottomMargin + ")") }; exit }
            $ps.BottomMargin = $margin
        }
        if ($null -ne $p.leftMargin) {
            $margin = ConvertTo-WordMarginPoints $p.leftMargin
            if ($null -eq $margin) { Output-Json @{ success = $false; error = ("leftMargin must be 0-1584 points (got " + $p.leftMargin + ")") }; exit }
            $ps.LeftMargin = $margin
        }
        if ($null -ne $p.rightMargin) {
            $margin = ConvertTo-WordMarginPoints $p.rightMargin
            if ($null -eq $margin) { Output-Json @{ success = $false; error = ("rightMargin must be 0-1584 points (got " + $p.rightMargin + ")") }; exit }
            $ps.RightMargin = $margin
        }
        if ($null -ne $p.orientation) {
            $ps.Orientation = if ($p.orientation -eq "landscape") { 1 } else { 0 }
        }
        if ($null -ne $p.paperSize) {
            $sizeMap = @{ A4 = 7; A3 = 6; Letter = 1; Legal = 5 }
            $size = $sizeMap[$p.paperSize]
            if ($null -eq $size) { $size = 7 }
            $ps.PaperSize = $size
        }
        # Read the values back so the caller can verify what actually took effect, and so the tool
        # does not have to treat a missing payload as a failure.
        $applied = @{}
        if ($null -ne $p.orientation) { $applied.orientation = $(if ([int]$ps.Orientation -eq 1) { "landscape" } else { "portrait" }) }
        if ($null -ne $p.topMargin) { $applied.marginTop = [int]$ps.TopMargin }
        if ($null -ne $p.bottomMargin) { $applied.marginBottom = [int]$ps.BottomMargin }
        if ($null -ne $p.leftMargin) { $applied.marginLeft = [int]$ps.LeftMargin }
        if ($null -ne $p.rightMargin) { $applied.marginRight = [int]$ps.RightMargin }
        if ($null -ne $p.paperSize) { $applied.paperSize = $p.paperSize }
        Output-Json @{ success = $true; data = @{ settings = $applied } }
    }

    "insertPageBreak" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $breakType = if ($p.type) { $p.type } else { "page" }
        $breakTypeMap = @{ page = 7; column = 8; section = 2; sectionContinuous = 3 }
        $bt = $breakTypeMap[$breakType]
        if ($null -eq $bt) { $bt = 7 }
        $word.Selection.InsertBreak($bt)
        Output-Json @{ success = $true; data = @{ type = $breakType } }
    }

    "insertHeader" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $sectionNo = if ($null -ne $p.section) { [int]$p.section } else { 1 }
        if ($sectionNo -lt 1 -or $sectionNo -gt $doc.Sections.Count) {
            Output-Json @{ success = $false; error = ("section " + $sectionNo + " does not exist; the document has " + $doc.Sections.Count) }
            exit
        }
        $section = $doc.Sections.Item($sectionNo)
        $header = $section.Headers.Item(1)
        $header.Range.Text = if ($p.text) { $p.text } else { "" }
        if ($p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2 }
            $header.Range.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        Output-Json @{ success = $true; data = @{ section = $sectionNo; text = $header.Range.Text } }
    }

    "insertFooter" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $sectionNo = if ($null -ne $p.section) { [int]$p.section } else { 1 }
        if ($sectionNo -lt 1 -or $sectionNo -gt $doc.Sections.Count) {
            Output-Json @{ success = $false; error = ("section " + $sectionNo + " does not exist; the document has " + $doc.Sections.Count) }
            exit
        }
        $section = $doc.Sections.Item($sectionNo)
        $footer = $section.Footers.Item(1)
        $footer.Range.Text = if ($p.text) { $p.text } else { "" }
        if ($p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2 }
            $footer.Range.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        if ($p.includePageNumber) {
            $footer.Range.InsertAfter(" - 第 ")
            $footer.Range.Fields.Add($footer.Range, -1, "PAGE", $false)
            $footer.Range.InsertAfter(" 页 ")
        }
        Output-Json @{ success = $true; data = @{ section = $sectionNo } }
    }

    "insertHyperlink" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $range = Get-MainTextRange $word $doc
        $url = if ($p.url) { $p.url } else { $p.address }
        $text = if ($p.text) { $p.text } elseif ($p.displayText) { $p.displayText } else { $url }
        if ($range.Text -and $range.Text.Trim() -ne "") {
            $doc.Hyperlinks.Add($range, $url)
        } else {
            $range.Text = $text
            $doc.Hyperlinks.Add($doc.Range($range.Start, $range.Start + $text.Length), $url)
        }
        Output-Json @{ success = $true; data = @{ url = $url; text = $text } }
    }

    "insertBookmark" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if (-not $p.name) { Output-Json @{ success = $false; error = "name required" }; exit }
        $range = $word.Selection.Range
        $doc.Bookmarks.Add($p.name, $range)
        Output-Json @{ success = $true; data = @{ name = $p.name } }
    }

    "getBookmarks" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $bookmarks = @()
        for ($i = 1; $i -le $doc.Bookmarks.Count; $i++) {
            $bm = $doc.Bookmarks.Item($i)
            $bookmarks += @{ name = $bm.Name; start = $bm.Start; end = $bm.End }
        }
        Output-Json @{ success = $true; data = @{ bookmarks = $bookmarks; count = $bookmarks.Count } }
    }

    "getDocumentParagraphs" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $startIdx = if ($null -ne $p.startParagraph) { [int]$p.startParagraph } else { 1 }
        $endIdx = if ($null -ne $p.endParagraph) { [int]$p.endParagraph } else { [Math]::Min($doc.Paragraphs.Count, $startIdx + 49) }
        if ($endIdx -gt $doc.Paragraphs.Count) { $endIdx = $doc.Paragraphs.Count }
        $paragraphs = @()
        for ($i = $startIdx; $i -le $endIdx; $i++) {
            $para = $doc.Paragraphs.Item($i)
            $text = $para.Range.Text
            $text = $text.TrimEnd([char[]]@([char]13, [char]10))
            if ($text.Length -gt 200) { $text = $text.Substring(0, 200) + "..." }
            $styleName = ""
            try { $styleName = $para.Range.Style.NameLocal } catch { $styleName = "" }
            $paragraphs += @{ index = $i; text = $text; style = $styleName; start = $para.Range.Start; end = $para.Range.End }
        }
        Output-Json @{ success = $true; data = @{ paragraphs = $paragraphs; totalCount = $doc.Paragraphs.Count; returnedCount = $paragraphs.Count } }
    }

    "findInDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        if (-not $p.findText) { Output-Json @{ success = $false; error = "findText required" }; exit }
        $matchCase = if ($null -ne $p.matchCase) { [bool]$p.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $p.matchWholeWord) { [bool]$p.matchWholeWord } else { $false }
        $maxResults = if ($null -ne $p.maxResults) { [int]$p.maxResults } else { 20 }
        $results = @()
        $searchRange = $doc.Content.Duplicate
        $searchRange.Find.ClearFormatting()
        $searchRange.Find.Text = $p.findText
        $found = $searchRange.Find.Execute($p.findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, "", 0)
        while ($found -and $results.Count -lt $maxResults) {
            $matchStart = $searchRange.Start
            $matchEnd = $searchRange.End
            $matchText = $searchRange.Text
            $paraIdx = 0
            for ($pi = 1; $pi -le $doc.Paragraphs.Count; $pi++) {
                $pr = $doc.Paragraphs.Item($pi).Range
                if ($matchStart -ge $pr.Start -and $matchStart -le $pr.End) { $paraIdx = $pi; break }
            }
            $ctxStart = [Math]::Max(0, $matchStart - 50)
            $ctxEnd = [Math]::Min($doc.Content.End, $matchEnd + 50)
            $ctxRange = $doc.Range($ctxStart, $ctxEnd)
            $context = $ctxRange.Text
            $results += @{ text = $matchText; start = $matchStart; end = $matchEnd; paragraphIndex = $paraIdx; context = $context }
            $searchRange = $doc.Range($matchEnd, $doc.Content.End)
            $searchRange.Find.ClearFormatting()
            $searchRange.Find.Text = $p.findText
            $found = $searchRange.Find.Execute($p.findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, "", 0)
        }
        Output-Json @{ success = $true; data = @{ results = $results; count = $results.Count; findText = $p.findText } }
    }

    "smartFillField" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        if (-not $p.keyword) { Output-Json @{ success = $false; error = "keyword required" }; exit }
        if ($null -eq $p.value) { Output-Json @{ success = $false; error = "value required（空字符串将清除该字段内容）" }; exit }
        $fillMode = if ($p.fillMode) { $p.fillMode.ToLower() } else { "auto" }

        $searchRange = $doc.Content.Duplicate
        $searchRange.Find.ClearFormatting()
        $searchRange.Find.Text = $p.keyword
        $found = $searchRange.Find.Execute($p.keyword, $false, $false, $false, $false, $false, $true, 1, $false, "", 0)
        if (-not $found) { Output-Json @{ success = $false; error = "Keyword '$($p.keyword)' not found" }; exit }

        $matchStart = $searchRange.Start
        $matchEnd = $searchRange.End
        $paraRange = $searchRange.Paragraphs.Item(1).Range
        $paraText = $paraRange.Text
        $detectedMode = $fillMode

        if ($fillMode -eq "auto") {
            if ($paraText -match '[\{\【\[]' + [regex]::Escape($p.keyword) + '[\}\】\]]') {
                $detectedMode = "placeholder"
            }
            elseif ($paraText -match [regex]::Escape($p.keyword) + '\s*[：:]\s*_+') {
                $detectedMode = "underline"
            }
            elseif ($paraText -match [regex]::Escape($p.keyword) + '\s*[：:]') {
                $detectedMode = "afterColon"
            }
            else {
                $afterKeyword = $paraText.Substring([Math]::Min($paraText.IndexOf($p.keyword) + $p.keyword.Length, $paraText.Length))
                $afterKeyword = $afterKeyword.TrimStart().TrimEnd([char[]]@([char]13, [char]10)).TrimEnd()
                if ($afterKeyword.Length -eq 0 -or $afterKeyword -match '^[\s\r\n]*$') {
                    $detectedMode = "afterLabel"
                } else {
                    $detectedMode = "afterColon"
                }
            }
        }

        $fillResult = ""
        switch ($detectedMode) {
            "placeholder" {
                $fullPlaceholder = ""
                if ($paraText -match '([\{\【\[])' + [regex]::Escape($p.keyword) + '([\}\】\]])') {
                    $fullPlaceholder = $Matches[0]
                } else {
                    $fullPlaceholder = $p.keyword
                }
                $replaceRange = $doc.Content.Duplicate
                $replaceRange.Find.ClearFormatting()
                $replaceRange.Find.Text = $fullPlaceholder
                $found2 = $replaceRange.Find.Execute($fullPlaceholder, $false, $false, $false, $false, $false, $true, 1, $false, $p.value, 1)
                if ($found2) { $fillResult = "Replaced placeholder '$fullPlaceholder' with '$($p.value)'" }
                else { $fillResult = "Failed to replace placeholder" }
            }
            "underline" {
                $keywordEndPos = $matchEnd
                $fillDone = $false
                for ($ri = 1; $ri -le $paraRange.Words.Count; $ri++) {
                    $wordObj = $paraRange.Words.Item($ri)
                    if ($wordObj.Start -gt $keywordEndPos -and $wordObj.Font.Underline -ne 0 -and $wordObj.Text -match '_+') {
                        $wordObj.Text = $p.value
                        $wordObj.Font.Underline = 1
                        $fillDone = $true
                        $fillResult = "Filled underlined field after '$($p.keyword)' with '$($p.value)'"
                        break
                    }
                }
                if (-not $fillDone) {
                    $underscorePattern = "_+"
                    $afterKeyRange = $doc.Range($keywordEndPos, $paraRange.End)
                    $afterKeyRange.Find.ClearFormatting()
                    $afterKeyRange.Find.Text = $underscorePattern
                    $afterKeyRange.Find.MatchWildcards = $true
                    $foundUl = $afterKeyRange.Find.Execute($underscorePattern, $false, $false, $false, $false, $false, $true, 1, $false, $p.value, 1)
                    if ($foundUl) {
                        $fillResult = "Filled underline after '$($p.keyword)' with '$($p.value)'"
                    } else {
                        $insertRange = $doc.Range($keywordEndPos, $keywordEndPos)
                        $insertRange.InsertAfter($p.value)
                        $fillResult = "Inserted '$($p.value)' after keyword '$($p.keyword)' (no underline found)"
                    }
                }
            }
            "afterColon" {
                $afterKeyRange = $doc.Range($matchEnd, $paraRange.End)
                $afterKeyText = $afterKeyRange.Text
                $colonOffset = $afterKeyText.IndexOfAny([char[]]@('：', ':'))
                if ($colonOffset -ge 0) {
                    $insertPos = $matchEnd + $colonOffset + 1
                    $afterColonRange = $doc.Range($insertPos, $paraRange.End - 1)
                    $afterColonText = $afterColonRange.Text.TrimStart().TrimEnd([char[]]@([char]13, [char]10)).TrimEnd()
                    if ($afterColonText.Length -gt 0 -and $afterColonText -notmatch '^[\s\r\n_　]+$') {
                        $afterColonRange.Text = $p.value
                        $fillResult = "Replaced content after colon with '$($p.value)'"
                    } else {
                        $insertRange = $doc.Range($insertPos, $insertPos)
                        $insertRange.InsertAfter($p.value)
                        $fillResult = "Inserted '$($p.value)' after colon following '$($p.keyword)'"
                    }
                } else {
                    $insertRange = $doc.Range($matchEnd, $matchEnd)
                    $insertRange.InsertAfter("：" + $p.value)
                    $fillResult = "No colon found, inserted '：$($p.value)' after '$($p.keyword)'"
                }
            }
            "afterLabel" {
                $insertRange = $doc.Range($matchEnd, $matchEnd)
                $insertRange.InsertAfter($p.value)
                $fillResult = "Inserted '$($p.value)' after label '$($p.keyword)'"
            }
            default {
                $insertRange = $doc.Range($matchEnd, $matchEnd)
                $insertRange.InsertAfter($p.value)
                $fillResult = "Inserted '$($p.value)' after '$($p.keyword)' (default mode)"
            }
        }
        Output-Json @{ success = $true; data = @{ keyword = $p.keyword; value = $p.value; fillMode = $detectedMode; result = $fillResult } }
    }

    "replaceBookmarkContent" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        if (-not $p.name) { Output-Json @{ success = $false; error = "name required" }; exit }
        if ($null -eq $p.text) { Output-Json @{ success = $false; error = "text required" }; exit }
        try {
            $bm = $doc.Bookmarks.Item($p.name)
            $bmStart = $bm.Start
            $bmRange = $bm.Range
            $bmRange.Text = $p.text
            $newEnd = $bmStart + $p.text.Length
            $newRange = $doc.Range($bmStart, $newEnd)
            $doc.Bookmarks.Add($p.name, $newRange) | Out-Null
            Output-Json @{ success = $true; data = @{ name = $p.name; text = $p.text; start = $bmStart; end = $newEnd } }
        } catch {
            Output-Json @{ success = $false; error = "Bookmark '$($p.name)' not found or failed to replace: $($_.Exception.Message)" }
        }
    }

    "addComment" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $text = if ($p.text) { $p.text } else { $p.comment }
        if (-not $text) { Output-Json @{ success = $false; error = "text required" }; exit }
        $range = $word.Selection.Range
        $doc.Comments.Add($range, $text)
        Output-Json @{ success = $true }
    }

    "getComments" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $comments = @()
        for ($i = 1; $i -le $doc.Comments.Count; $i++) {
            $c = $doc.Comments.Item($i)
            $comments += @{ index = $i; text = $c.Range.Text; author = if ($c.Author) { $c.Author } else { "" }; date = if ($c.Date) { $c.Date.ToString() } else { "" } }
        }
        Output-Json @{ success = $true; data = @{ comments = $comments; count = $comments.Count } }
    }

    "getDocumentStats" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $stats = @{ name = $doc.Name; path = $doc.FullName; pages = $doc.ComputeStatistics(2); words = $doc.ComputeStatistics(0); characters = $doc.ComputeStatistics(3); paragraphs = $doc.ComputeStatistics(4); lines = $doc.ComputeStatistics(1) }
        Output-Json @{ success = $true; data = $stats }
    }

    "insertImage" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $path = if ($p.path) { $p.path } else { $p.filePath }
        if (-not $path) { Output-Json @{ success = $false; error = "path required" }; exit }
        $range = $word.Selection.Range
        $shape = $doc.InlineShapes.AddPicture($path, $false, $true, $range)
        if ($p.width) { $shape.Width = $p.width }
        if ($p.height) { $shape.Height = $p.height }
        if ($p.scale) { $shape.ScaleWidth = $p.scale; $shape.ScaleHeight = $p.scale }
        Output-Json @{ success = $true; data = @{ width = $shape.Width; height = $shape.Height } }
    }

    "applyStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        if ($p.range -and $null -ne $p.range.start -and $null -ne $p.range.end) {
            $range = $doc.Range([int]$p.range.start, [int]$p.range.end)
        } else {
            $range = $word.Selection.Range
        }
        $range.Style = $p.styleName
        Output-Json @{ success = $true; data = @{ affectedText = $range.Text } }
    }

    # ==================== Word: Proofreading ====================
    "enableTrackChanges" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $enable = if ($null -ne $p.enable) { [bool]$p.enable } else { $true }
        $doc.TrackRevisions = $enable
        Output-Json @{ success = $true; data = @{ trackChanges = $enable; active = $doc.TrackRevisions } }
    }

    "getTrackChangesStatus" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $revisionCount = 0
        try { $revisionCount = $doc.Revisions.Count } catch { Add-WpsWarning $_.Exception.Message }
        Output-Json @{ success = $true; data = @{ trackChanges = [bool]$doc.TrackRevisions; revisionCount = $revisionCount } }
    }

    "replaceRange" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        if ($null -eq $p.startPos -or $null -eq $p.endPos) { Output-Json @{ success = $false; error = "startPos and endPos required" }; exit }
        if ($null -eq $p.text) { Output-Json @{ success = $false; error = "text required" }; exit }
        try {
            $range = $doc.Range([int]$p.startPos, [int]$p.endPos)
            $originalText = $range.Text
            $range.Text = $p.text
            $actualEndPos = $range.End
            Output-Json @{ success = $true; data = @{ startPos = [int]$p.startPos; originalEndPos = [int]$p.endPos; endPos = $actualEndPos; originalText = $originalText.TrimEnd("`r`n"); newText = $p.text } }
        } catch {
            Output-Json @{ success = $false; error = "Failed to replace range: $($_.Exception.Message)" }
        }
    }

    # ==================== PPT ====================
    "getActivePresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $slides = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $shapes = @()
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                $text = ""
                try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $text = $shape.TextFrame.TextRange.Text.Substring(0, [Math]::Min(50, $shape.TextFrame.TextRange.Text.Length)) } } catch { Add-WpsWarning $_.Exception.Message }
                $shapes += @{ name = $shape.Name; type = $shape.Type; text = $text }
            }
            $slides += @{ index = $i; shapeCount = $slide.Shapes.Count; shapes = $shapes }
        }
        Output-Json @{ success = $true; data = @{ name = $pres.Name; path = $pres.FullName; slideCount = $pres.Slides.Count; slides = $slides } }
    }

    "openPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        if (-not $p.path) { Output-Json @{ success = $false; error = "path required" }; exit }
        $prevAlerts = Set-WpsAlertsSuppressed $ppt 'ppt'
        $pres = $null
        $openError = $null
        try { $pres = Open-PptPresentation $ppt $p.path } catch { $openError = $_.Exception.Message }
        Restore-WpsAlerts $ppt $prevAlerts
        if ($null -ne $openError) { Output-Json @{ success = $false; error = (Format-WpsOpenError $p.path 'ppt' $openError) }; exit }
        Output-Json @{ success = $true; data = @{ name = $pres.Name; path = $pres.FullName; slideCount = $pres.Slides.Count } }
    }

    "closePresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = if ($p.name) { $ppt.Presentations.Item($p.name) } else { $ppt.ActivePresentation }
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No presentation" }; exit }
        $name = $pres.Name
        $path = ""
        try { $path = [string]$pres.Path } catch { Add-WpsWarning $_.Exception.Message }
        # Same "save" vs "saveChanges" mismatch as closeWorkbook.
        $saveWanted = $true
        if ($null -ne $p.save) { $saveWanted = [bool]$p.save }
        elseif ($null -ne $p.saveChanges) { $saveWanted = [bool]$p.saveChanges }
        $saved = $false
        $warning = $null
        if ($saveWanted) {
            if ($path -eq "") {
                # Presentation.Close() takes no save argument, so it only stays quiet if the
                # document is already marked clean.
                $warning = "presentation was never saved to disk; closed without saving (use save_as first if you need it on disk)"
            } else {
                $pres.Save()
                $saved = $true
            }
        }
        if (-not $saved) { $pres.Saved = $true }
        $pres.Close()
        $data = @{ closed = $name; saved = $saved; saveRequested = $saveWanted }
        if ($null -ne $warning) { $data.warning = $warning }
        Output-Json @{ success = $true; data = $data }
    }

    "getOpenPresentations" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $presentations = @()
        for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
            $presItem = $ppt.Presentations.Item($i)
            $presentations += @{ name = $presItem.Name; path = $presItem.FullName; slideCount = $presItem.Slides.Count }
        }
        Output-Json @{ success = $true; data = @{ presentations = $presentations; count = $presentations.Count } }
    }

    "switchPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $presItem = $ppt.Presentations.Item($(if ($p.name) { $p.name } else { $p.index }))
        $presItem.Windows.Item(1).Activate()
        Output-Json @{ success = $true; data = @{ name = $presItem.Name } }
    }

    "deleteSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        $pres.Slides.Item($index).Delete()
        Output-Json @{ success = $true; data = @{ deleted = $index } }
    }

    "duplicateSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $newSlide = $slide.Duplicate()
        $newIndex = $newSlide.Item(1).SlideIndex
        if ($null -ne $p.targetIndex) {
            $target = [int]$p.targetIndex
            if ($target -lt 1 -or $target -gt $pres.Slides.Count) { Output-Json @{ success = $false; error = "targetIndex is out of range" }; exit }
            # Move the fresh copy to the requested position instead of leaving it next to the source.
            $pres.Slides.Item($newIndex).MoveTo($target)
            $newIndex = $target
        }
        Output-Json @{ success = $true; data = @{ sourceIndex = $index; newIndex = $newIndex } }
    }

    "insertSlidesFromFile" {
        # 从另一个 PPT 文件把整页幻灯片插入当前演示文稿，保留来源格式（跨PPT整合）
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $src = if ($p.filePath) { $p.filePath } elseif ($p.path) { $p.path } else { $null }
        if (-not $src) { Output-Json @{ success = $false; error = "filePath required" }; exit }
        if (-not (Test-Path $src)) { Output-Json @{ success = $false; error = "source file not found: $src" }; exit }
        $before = $pres.Slides.Count
        $afterIndex = if ($null -ne $p.afterIndex) { [int]$p.afterIndex } else { $before }
        if ($afterIndex -lt 0) { $afterIndex = 0 }
        if ($afterIndex -gt $before) { $afterIndex = $before }
        if ($null -ne $p.slideStart -and $null -ne $p.slideEnd) {
            [void]$pres.Slides.InsertFromFile($src, $afterIndex, [int]$p.slideStart, [int]$p.slideEnd)
        } elseif ($null -ne $p.slideStart) {
            [void]$pres.Slides.InsertFromFile($src, $afterIndex, [int]$p.slideStart)
        } else {
            [void]$pres.Slides.InsertFromFile($src, $afterIndex)
        }
        $after = $pres.Slides.Count
        Output-Json @{ success = $true; data = @{ inserted = ($after - $before); afterIndex = $afterIndex; totalSlides = $after; source = $src } }
    }

    "moveSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $fromIndex = if ($p.from) { $p.from } else { $p.fromIndex }
        $toIndex = if ($p.to) { $p.to } else { $p.toIndex }
        $pres.Slides.Item($fromIndex).MoveTo($toIndex)
        Output-Json @{ success = $true; data = @{ from = $fromIndex; to = $toIndex } }
    }

    "getSlideCount" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        Output-Json @{ success = $true; data = @{ count = $pres.Slides.Count } }
    }

    "getSlideInfo" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $txt = ""
            try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $txt = $shape.TextFrame.TextRange.Text } } catch { Add-WpsWarning $_.Exception.Message }
            try {
                if ($shape.HasTable -and [string]::IsNullOrEmpty($txt)) {
                    $tbl = $shape.Table; $rws = @()
                    for ($rr = 1; $rr -le $tbl.Rows.Count; $rr++) {
                        $cs = @()
                        for ($cc = 1; $cc -le $tbl.Columns.Count; $cc++) {
                            $cv = ""
                            try { $cv = $tbl.Cell($rr, $cc).Shape.TextFrame.TextRange.Text } catch { Add-WpsWarning $_.Exception.Message }
                            $cs += $cv
                        }
                        $rws += ($cs -join "|")
                    }
                    $txt = "[表格] " + ($rws -join " ;; ")
                }
            } catch { Add-WpsWarning $_.Exception.Message }
            $txt = ($txt -replace "[`r`n`v]", " ")
            if ($txt.Length -gt 90) { $txt = $txt.Substring(0, 90) + "…" }
            $hasTxt = $false
            if ($shape.HasTextFrame) { $hasTxt = $true }
            $shapes += @{ name = $shape.Name; type = "$($shape.Type)"; hasText = $hasTxt; text = $txt }
        }
        Output-Json @{ success = $true; data = @{ index = $index; slideIndex = $index; shapeCount = $slide.Shapes.Count; shapesCount = $slide.Shapes.Count; shapes = $shapes; layout = "$($slide.Layout)" } }
    }

    "switchSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        $ppt.ActiveWindow.View.GotoSlide($index)
        Output-Json @{ success = $true; data = @{ currentSlide = $index } }
    }

    "setSlideLayout" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $layoutMap = @{ title = 1; titleContent = 2; blank = 12; twoColumn = 4; comparison = 5; titleOnly = 11 }
        $layout = $layoutMap[$p.layout]
        if ($null -eq $layout) { $layout = if ($p.layout) { $p.layout } else { 2 } }
        $pres.Slides.Item($index).Layout = $layout
        Output-Json @{ success = $true; data = @{ slideIndex = $index; layout = $layout } }
    }

    "getSlideNotes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $notes = ""
        try { $notes = $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text } catch { $notes = "" }
        Output-Json @{ success = $true; data = @{ slideIndex = $index; notes = $notes } }
    }

    "setSlideNotes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = if ($p.notes) { $p.notes } else { "" }
        Output-Json @{ success = $true; data = @{ slideIndex = $index } }
    }


    "addSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $layouts = @{ title = 1; title_content = 2; blank = 12; two_column = 3; comparison = 34 }
        $layoutKey = if ($p.layout) { $p.layout } else { "title_content" }
        $layoutType = $layouts[$layoutKey]
        if ($null -eq $layoutType) { $layoutType = 2 }
        $position = if ($p.position) { $p.position } else { $pres.Slides.Count + 1 }
        $slide = $pres.Slides.Add($position, $layoutType)
        if ($p.title -and $slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        if ($p.content) {
            try {
                foreach ($shape in $slide.Shapes) {
                    try {
                        if ($shape.PlaceholderFormat.Type -eq 2) {
                            $shape.TextFrame.TextRange.Text = $p.content
                        }
                    } catch { Add-WpsWarning $_.Exception.Message }
                }
            } catch { Add-WpsWarning $_.Exception.Message }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $position; layout = $layoutKey } }
    }

    "addTextBox" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 50 }
        $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
        $shape.TextFrame.TextRange.Text = $p.text
        if ($p.fontSize) { $shape.TextFrame.TextRange.Font.Size = $p.fontSize }
        if ($p.fontName) { $shape.TextFrame.TextRange.Font.Name = $p.fontName }
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name } }
    }

    "deleteTextBox" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $textboxIndex = Resolve-TextBoxIndex $slide $p
        if ($null -eq $textboxIndex) { Output-Json @{ success = $false; error = "no text box matched textboxIndex/shapeIndex/name" }; exit }
        $shape = $slide.Shapes.Item($textboxIndex)
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $textboxIndex; deletedShape = $shape.Name } }
    }

    "getTextBoxes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $textBoxes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            if ($shape.HasTextFrame) {
                $text = ""
                try { $text = $shape.TextFrame.TextRange.Text } catch { Add-WpsWarning $_.Exception.Message }
                $textBoxes += @{ name = $shape.Name; index = $i; text = $text; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
            }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; textBoxes = $textBoxes; count = $textBoxes.Count } }
    }

    "setTextBoxText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $textboxIndex = Resolve-TextBoxIndex $slide $p
        if ($null -eq $textboxIndex) { Output-Json @{ success = $false; error = "no text box matched textboxIndex/shapeIndex/name" }; exit }
        $shape = $slide.Shapes.Item($textboxIndex)
        $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; text = $p.text } }
    }

    "setTextBoxStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $textboxIndex = Resolve-TextBoxIndex $slide $p
        if ($null -eq $textboxIndex) { Output-Json @{ success = $false; error = "no text box matched textboxIndex/shapeIndex/name" }; exit }
        $shape = $slide.Shapes.Item($textboxIndex)
        $tr = $shape.TextFrame.TextRange
        if ($p.fontSize) { $tr.Font.Size = $p.fontSize }
        if ($p.fontName) { $tr.Font.Name = $p.fontName }
        if ($null -ne $p.bold) { $tr.Font.Bold = $p.bold }
        if ($null -ne $p.italic) { $tr.Font.Italic = $p.italic }
        if ($p.color) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $colorValue) { $tr.Font.Color = $colorValue }
        }
        if ($p.alignment) {
            $alignMap = @{ left = 1; center = 2; right = 3 }
            $tr.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setSlideTitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        if ($slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; title = $p.title } }
    }

    "getSlideTitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $title = ""
        if ($slide.Shapes.HasTitle) { $title = $slide.Shapes.Title.TextFrame.TextRange.Text }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; title = $title } }
    }

    "setSlideSubtitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $placeholderType = $null
            try { $placeholderType = $shape.PlaceholderFormat.Type } catch { Add-WpsWarning $_.Exception.Message }
            if ($placeholderType -eq 2) {
                $shape.TextFrame.TextRange.Text = if ($p.subtitle) { $p.subtitle } else { "" }
                Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; subtitle = $p.subtitle } }
                return
            }
        }
        Output-Json @{ success = $false; error = "Subtitle placeholder not found" }
    }

    "setSlideContent" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $placeholderType = $null
            try { $placeholderType = $shape.PlaceholderFormat.Type } catch { Add-WpsWarning $_.Exception.Message }
            if ($placeholderType -eq 7) {
                $shape.TextFrame.TextRange.Text = if ($p.content) { $p.content } else { "" }
                Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
                return
            }
        }
        Output-Json @{ success = $false; error = "Content placeholder not found" }
    }

    "addShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shapeTypeMap = @{ rectangle = 1; oval = 9; triangle = 7; diamond = 4; pentagon = 51; hexagon = 52; arrow = 13; star = 12; heart = 21; cloud = 179 }
        $shapeType = $shapeTypeMap[$p.type]
        if ($null -eq $shapeType) { $shapeType = if ($p.type) { $p.type } else { 1 } }
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 100 }
        $height = if ($p.height) { $p.height } else { 100 }
        $shape = $slide.Shapes.AddShape($shapeType, $left, $top, $width, $height)
        if ($p.text) { $shape.TextFrame.TextRange.Text = $p.text }
        if ($p.fillColor) {
            $fillColor = Convert-HexColorToRgbInt([string]$p.fillColor)
            if ($null -ne $fillColor) { $shape.Fill.ForeColor.RGB = $fillColor }
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; slideIndex = $slideIndex } }
    }

    "deleteShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "getShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $shapes += @{ name = $shape.Name; index = $i; type = $shape.Type; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; shapes = $shapes; count = $shapes.Count } }
    }

    "setShapeStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.fillColor) {
            $fillColor = Convert-HexColorToRgbInt([string]$p.fillColor)
            if ($null -ne $fillColor) { $shape.Fill.ForeColor.RGB = $fillColor }
        }
        if ($p.lineColor) {
            $lineColor = Convert-HexColorToRgbInt([string]$p.lineColor)
            if ($null -ne $lineColor) { $shape.Line.ForeColor.RGB = $lineColor }
        }
        if ($p.lineWidth) { $shape.Line.Weight = $p.lineWidth }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; text = $p.text } }
    }

    "setShapePosition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height } }
    }

    "insertPptImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($null -ne $p.width) { $p.width } else { -1 }
        $height = if ($null -ne $p.height) { $p.height } else { -1 }
        $pic = $slide.Shapes.AddPicture($p.path, $false, $true, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $pic.Name; path = $p.path } }
    }

    "deletePptImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $pictureIndex = Resolve-PictureIndex $slide $p
        if ($null -eq $pictureIndex) { Output-Json @{ success = $false; error = "no picture matched imageIndex/shapeIndex/name" }; exit }
        $shape = $slide.Shapes.Item($pictureIndex)
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $pictureIndex; deletedShape = $shape.Name } }
    }

    "replacePptImage" {
        # 原位替换图片：保留原图位置/尺寸/旋转，删除旧图后在同一矩形插入新图（换图不动版式）
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $sel = if ($p.name) { $p.name } else { $p.shapeIndex }
        if ($null -eq $sel) { Output-Json @{ success = $false; error = "shapeIndex or name required" }; exit }
        $newPath = if ($p.path) { $p.path } elseif ($p.filePath) { $p.filePath } else { $null }
        if (-not $newPath) { Output-Json @{ success = $false; error = "path required" }; exit }
        if (-not (Test-Path $newPath)) { Output-Json @{ success = $false; error = "image not found: $newPath" }; exit }
        $old = $slide.Shapes.Item($sel)
        $l = $old.Left; $t = $old.Top; $w = $old.Width; $h = $old.Height
        $rot = 0
        try { $rot = $old.Rotation } catch { Add-WpsWarning $_.Exception.Message }
        $old.Delete()
        $pic = $slide.Shapes.AddPicture($newPath, $false, $true, $l, $t, $w, $h)
        try { $pic.Rotation = $rot } catch { Add-WpsWarning $_.Exception.Message }
        Output-Json @{ success = $true; data = @{ name = $pic.Name; left = $l; top = $t; width = $w; height = $h; path = $newPath } }
    }

    "setImageStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $pictureIndex = Resolve-PictureIndex $slide $p
        if ($null -eq $pictureIndex) { Output-Json @{ success = $false; error = "no picture matched imageIndex/shapeIndex/name" }; exit }
        $shape = $slide.Shapes.Item($pictureIndex)
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        if ($null -ne $p.rotation) { $shape.Rotation = $p.rotation }
        # style carries the border/shadow/opacity/crop vocabulary; each piece is applied or reported.
        $styleErrors = @()
        if ($null -ne $p.border) {
            $border = $p.border
            $borderOn = if ($null -ne $border.enabled) { [bool]$border.enabled } else { $true }
            try {
                $shape.Line.Visible = $(if ($borderOn) { -1 } else { 0 })
                if ($borderOn -and $null -ne $border.weight) { $shape.Line.Weight = [single]$border.weight }
                if ($borderOn -and $null -ne $border.color) {
                    $borderRgb = Convert-HexColorToRgbInt ([string]$border.color)
                    if ($null -ne $borderRgb) { $shape.Line.ForeColor.RGB = $borderRgb }
                }
            } catch { $styleErrors += ("border: " + $_.Exception.Message) }
        }
        if ($null -ne $p.shadow) {
            $shadowSpec = $p.shadow
            $shadowOn = if ($null -ne $shadowSpec.enabled) { [bool]$shadowSpec.enabled } else { $true }
            try {
                $shape.Shadow.Visible = $(if ($shadowOn) { -1 } else { 0 })
                if ($shadowOn -and $null -ne $shadowSpec.blur) { $shape.Shadow.Blur = [single]$shadowSpec.blur }
                if ($shadowOn -and $null -ne $shadowSpec.offsetX) { $shape.Shadow.OffsetX = [single]$shadowSpec.offsetX }
                if ($shadowOn -and $null -ne $shadowSpec.offsetY) { $shape.Shadow.OffsetY = [single]$shadowSpec.offsetY }
            } catch { $styleErrors += ("shadow: " + $_.Exception.Message) }
        }
        if ($null -ne $p.opacity) {
            try { $shape.Fill.Transparency = [double]$p.opacity } catch { $styleErrors += ("opacity: " + $_.Exception.Message) }
        }
        try {
            $cropHeight = [double]$shape.Height
            if ($null -ne $p.cropTop) { Set-ComValue $shape.PictureFormat "CropTop" ([double]$p.cropTop * $cropHeight) }
            if ($null -ne $p.cropBottom) { Set-ComValue $shape.PictureFormat "CropBottom" ([double]$p.cropBottom * $cropHeight) }
            if ($null -ne $p.cropLeft) { Set-ComValue $shape.PictureFormat "CropLeft" ([double]$p.cropLeft * [double]$shape.Width) }
            if ($null -ne $p.cropRight) { Set-ComValue $shape.PictureFormat "CropRight" ([double]$p.cropRight * [double]$shape.Width) }
        } catch { $styleErrors += ("crop: " + $_.Exception.Message) }
        if ($styleErrors.Count -gt 0) { Output-Json @{ success = $false; error = ("some image style properties failed: " + ($styleErrors -join "; ")) }; exit }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height } }
    }

    "exportSlideAsImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No target presentation (presentationName not found or no active presentation)" }; exit }
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        if ($slideIndex -lt 1 -or $slideIndex -gt $pres.Slides.Count) { Output-Json @{ success = $false; error = "slideIndex $slideIndex out of range (1..$($pres.Slides.Count)) in '$($pres.Name)'" }; exit }
        $slide = $pres.Slides.Item($slideIndex)
        $outputPath = if ($p.outputPath) { $p.outputPath } else { $p.path }
        if ([string]::IsNullOrEmpty($outputPath)) { Output-Json @{ success = $false; error = "Missing outputPath" }; exit }
        # 确保父目录存在
        $dir = Split-Path -Parent $outputPath
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $rawFormat = if ($p.format) { $p.format.ToString().ToUpper() } else { "PNG" }
        # JPEG 在 PowerPoint COM 中按 JPG 滤镜处理
        $filterName = if ($rawFormat -eq "JPEG") { "JPG" } else { $rawFormat }
        $width = if ($p.width) { [int]$p.width } else { 1280 }
        $height = if ($p.height) { [int]$p.height } else { 720 }
        try {
            $slide.Export($outputPath, $filterName, $width, $height)
        } catch {
            Output-Json @{ success = $false; error = "Slide.Export failed: $($_.Exception.Message)" }; exit
        }
        # 关键修复：校验文件是否真生成，避免 Export 静默失败却报成功
        if (Test-Path $outputPath) {
            $bytes = (Get-Item $outputPath).Length
            Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; outputPath = $outputPath; format = $filterName; width = $width; height = $height; bytes = $bytes; presentation = $pres.Name } }
        } else {
            Output-Json @{ success = $false; error = "Export 未抛错但文件未生成: $outputPath（WPS Slide.Export 可能不支持该 filter/尺寸参数或该路径）" }
        }
    }

    "insertPptTable" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $rows = if ($p.rows) { $p.rows } else { 3 }
        $cols = if ($p.cols) { $p.cols } else { 3 }
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 200 }
        $table = $slide.Shapes.AddTable($rows, $cols, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $table.Name; rows = $rows; cols = $cols } }
    }

    "setPptTableCell" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; exit }
        $cell = $shape.Table.Cell($p.row, $p.col)
        $textValue = if ($p.text) { $p.text } elseif ($p.value) { $p.value } else { "" }
        $cell.Shape.TextFrame.TextRange.Text = $textValue
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; text = $textValue } }
    }

    "getPptTableCell" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; exit }
        $table = $shape.Table
        $rowCount = $table.Rows.Count
        $colCount = $table.Columns.Count
        $value = $null
        if ($null -ne $p.row -and $null -ne $p.col) {
            try { $value = $table.Cell($p.row, $p.col).Shape.TextFrame.TextRange.Text } catch { $value = $null }
        }
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; value = $value; rowCount = $rowCount; colCount = $colCount } }
    }









    "setBackgroundGradient" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $bg = $slide.Background.Fill
        $bg.Visible = $true
        $bg.TwoColorGradient(1, 1)
        if ($p.color1) {
            $c1 = Convert-HexColorToRgbInt([string]$p.color1)
            if ($null -ne $c1) { $bg.ForeColor.RGB = $c1 }
        }
        if ($p.color2) {
            $c2 = Convert-HexColorToRgbInt([string]$p.color2)
            if ($null -ne $c2) { $bg.BackColor.RGB = $c2 }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }







    "setShapeRoundness" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.roundness) { $shape.Adjustments.Item(1) = $p.roundness }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }


    "alignShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $alignMap = @{ left = 1; center = 2; right = 3; top = 4; middle = 5; bottom = 6 }
        $align = $alignMap[$p.alignment]
        if ($null -eq $align) { $align = 1 }
        # ShapeRange.Align is unreliable here: PowerShell either cannot resolve the member in a
        # long-lived host or unwraps the ShapeRange into an Object[], so align by geometry instead.
        $indices = @()
        if ($null -ne $p.names) {
            foreach ($n in @($p.names)) { $indices += [int]$n }
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) { $indices += $i }
        }
        if ($indices.Count -eq 0) { Output-Json @{ success = $false; error = "no shapes to align" }; exit }
        $targets = @()
        foreach ($n in $indices) {
            if ($n -lt 1 -or $n -gt $slide.Shapes.Count) { Output-Json @{ success = $false; error = ("shape index " + $n + " is out of range") }; exit }
            $targets += $slide.Shapes.Item($n)
        }
        $lefts = @($targets | ForEach-Object { [double]$_.Left })
        $tops = @($targets | ForEach-Object { [double]$_.Top })
        foreach ($shape in $targets) {
            switch ($align) {
                1 { Set-ComValue $shape 'Left' $lefts[0] }
                2 { Set-ComValue $shape 'Left' ((($lefts | Measure-Object -Average).Average) - [double]$shape.Width / 2) }
                3 { Set-ComValue $shape 'Left' ((($lefts | Measure-Object -Maximum).Maximum) - [double]$shape.Width) }
                4 { Set-ComValue $shape 'Top' $tops[0] }
                5 { Set-ComValue $shape 'Top' ((($tops | Measure-Object -Average).Average) - [double]$shape.Height / 2) }
                6 { Set-ComValue $shape 'Top' ((($tops | Measure-Object -Maximum).Maximum) - [double]$shape.Height) }
            }
        }



        Output-Json @{ success = $true; data = @{ alignment = $p.alignment; shapes = $targets.Count; scope = $(if ($null -ne $p.names) { "selected" } else { "all-shapes" }) } }
    }

    "distributeShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $distMap = @{ horizontal = 0; vertical = 1 }
        $dist = $distMap[$p.direction]
        if ($null -eq $dist) { $dist = 0 }
        $range = $slide.Shapes.Range($p.names)
        $range.Distribute($dist, 1)
        Output-Json @{ success = $true; data = @{ direction = $p.direction } }
    }

    "groupShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $range = $slide.Shapes.Range($p.names)
        $group = $range.Group()
        Output-Json @{ success = $true; data = @{ name = $group.Name } }
    }

    "duplicateShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $dup = $shape.Duplicate()
        Output-Json @{ success = $true; data = @{ name = $dup.Name } }
    }

    "setShapeZOrder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $zMap = @{ front = 0; forward = 2; backward = 3; back = 1 }
        $z = $zMap[$p.zOrder]
        if ($null -eq $z) { $z = 0 }
        $shape.ZOrder($z)
        Output-Json @{ success = $true; data = @{ name = $shape.Name; zOrder = $p.zOrder } }
    }

    "insertPptChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 300 }
        # AddChart wants an XlChartType number while the tools name the type.
        $chartTypes = @{ bar = 57; bar_clustered = 57; column = 51; column_clustered = 51; column_stacked = 52; line = 4; line_markers = 65; pie = 5; doughnut = -4120; area = 1; scatter = -4169; radar = -4151 }
        $chartType = $p.type
        if ($null -eq $chartType) { Output-Json @{ success = $false; error = "type/chartType is required" }; exit }
        if ($chartType -is [string]) {
            $mapped = $chartTypes[([string]$chartType).ToLower()]
            if ($null -eq $mapped) {
                $numeric = 0
                if (-not [int]::TryParse([string]$chartType, [ref]$numeric)) { Output-Json @{ success = $false; error = ("unknown chart type '" + $chartType + "'; use " + (($chartTypes.Keys | Sort-Object) -join "/") + " or an XlChartType number") }; exit }
                $mapped = $numeric
            }
            $chartType = $mapped
        }
        $shape = $slide.Shapes.AddChart($chartType, $left, $top, $width, $height)
        # Chart data is not injected: doing so means opening the chart's embedded workbook, which can
        # leave a hidden document behind in a resident host. The tool says so instead of pretending.
        $titleError = $null
        if ($null -ne $p.title -and "$($p.title)" -ne "") {
            try { $shape.Chart.HasTitle = $true; $shape.Chart.ChartTitle.Text = [string]$p.title } catch { $titleError = $_.Exception.Message }
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; chartType = $chartType; title = $p.title; titleError = $titleError } }
    }



    "setSlideTransition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        if (-not (Test-WpsSlideIndex $pres $slideIndex)) { Output-Json @{ success = $false; error = "slideIndex is out of range" }; exit }
        $slide = $pres.Slides.Item($slideIndex)
        $effect = Get-PptEntryEffect $p.effect
        if ($null -eq $effect) { Output-Json @{ success = $false; error = ("unknown transition '" + $p.effect + "'; use none/cut/fade/dissolve/push/wipe/split/reveal/cover/curtains or a PpEntryEffect number") }; exit }
        $slide.SlideShowTransition.EntryEffect = $effect
        if ($null -ne $p.duration) { $slide.SlideShowTransition.Duration = [single]$p.duration }
        if ($null -ne $p.sound -and "$($p.sound)" -ne "") {
            $soundPath = Resolve-InputFilePath $p.sound
            if ($null -eq $soundPath) { Output-Json @{ success = $false; error = ("sound file not found: " + "$($p.sound)") }; exit }
            $slide.SlideShowTransition.SoundEffect.ImportFromFile($soundPath)
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; effect = $p.effect; entryEffect = $effect; duration = $p.duration } }
    }

    "setSlideBackground" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $spec = Get-PptBackgroundSpec $p
        if ($null -ne $spec.error) { Output-Json @{ success = $false; error = $spec.error }; exit }
        $slide.FollowMasterBackground = $false
        Set-PptBackgroundFill $slide.Background.Fill $spec
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; type = $spec.kind } }
    }

























    "getSlideMaster" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $master = $pres.SlideMaster
        $masters = @()
        for ($i = 1; $i -le $master.Shapes.Count; $i++) {
            $shape = $master.Shapes.Item($i)
            $masters += @{ name = $shape.Name; type = $shape.Type; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
        }
        Output-Json @{ success = $true; data = @{ shapeCount = $masters.Count; shapes = $masters } }
    }

    "setMasterBackground" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $master = $pres.SlideMaster
        $spec = Get-PptBackgroundSpec $p
        if ($null -ne $spec.error) { Output-Json @{ success = $false; error = $spec.error }; exit }
        Set-PptBackgroundFill $master.Background.Fill $spec
        Output-Json @{ success = $true; data = @{ type = $spec.kind } }
    }

    "addMasterElement" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $master = $pres.SlideMaster

        $type = if ($p.type) { $p.type } else { "textbox" }
        $left = if ($p.left) { $p.left } else { 600 }
        $top = if ($p.top) { $p.top } else { 20 }
        $width = if ($p.width) { $p.width } else { 100 }
        $height = if ($p.height) { $p.height } else { 40 }

        if ($type -eq "textbox") {
            $shape = $master.Shapes.AddTextbox(1, $left, $top, $width, $height)
            $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "Logo" }
            $shape.TextFrame.TextRange.Font.Size = if ($p.fontSize) { $p.fontSize } else { 14 }
            $shape.TextFrame.TextRange.Font.Bold = $true
        } elseif ($type -eq "shape") {
            $shape = $master.Shapes.AddShape($(if ($p.shapeType) { $p.shapeType } else { 5 }), $left, $top, $width, $height)
            $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
        } else {
            Output-Json @{ success = $false; error = "Unsupported master element type" }
            break
        }

        Output-Json @{ success = $true; data = @{ name = $shape.Name; type = $type } }
    }








    "setPptChartData" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.chartName) { $p.chartName } else { $p.chartIndex }))
        $chart = $shape.Chart
        $chart.ChartData.Activate()
        $dataSheet = $chart.ChartData.Workbook.Worksheets.Item(1)
        if ($p.data) {
            for ($r = 0; $r -lt $p.data.Count; $r++) {
                $rowData = $p.data[$r]
                for ($c = 0; $c -lt $rowData.Count; $c++) {
                    Set-ComValue $dataSheet.Cells.Item($r + 1, $c + 1) 'Value2' $rowData[$c]
                }
            }
        }
        Output-Json @{ success = $true; data = @{ chartName = $shape.Name } }
    }

    "setPptChartStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.chartName) { $p.chartName } else { $p.chartIndex }))
        $chart = $shape.Chart
        if ($p.title) {
            $chart.HasTitle = $true
            $chart.ChartTitle.Text = $p.title
        }
        if ($null -ne $p.hasLegend) {
            $chart.HasLegend = $p.hasLegend
        }
        Output-Json @{ success = $true; data = @{ chartName = $shape.Name } }
    }

    "removeAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        if ($p.index) {
            $seq.Item($p.index).Delete()
        } else {
            while ($seq.Count -gt 0) { $seq.Item(1).Delete() }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "getAnimations" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        $animations = @()
        for ($i = 1; $i -le $seq.Count; $i++) {
            $effect = $seq.Item($i)
            $animations += @{ index = $i; shapeName = $effect.Shape.Name; effectType = $effect.EffectType; duration = $effect.Timing.Duration }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; animations = $animations; count = $animations.Count } }
    }

    "setAnimationOrder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        $effect = $seq.Item($p.from)
        $effect.MoveTo($p.to)
        Output-Json @{ success = $true; data = @{ from = $p.from; to = $p.to } }
    }

    "removeSlideTransition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.SlideShowTransition.EntryEffect = 0
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "applyTransitionToAll" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $effect = Get-PptEntryEffect $p.transition
        if ($null -eq $effect) { Output-Json @{ success = $false; error = ("unknown transition '" + $p.transition + "'; use none/cut/fade/dissolve/push/wipe/split/reveal/cover/curtains or a PpEntryEffect number") }; exit }
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $slide.SlideShowTransition.EntryEffect = $effect
            if ($null -ne $p.duration) { $slide.SlideShowTransition.Duration = [single]$p.duration }
        }
        Output-Json @{ success = $true; data = @{ transition = $p.transition; entryEffect = $effect; appliedTo = $pres.Slides.Count } }
    }

    "setBackgroundColor" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $colorValue = Convert-HexColorToRgbInt([string]$p.color)
        if ($p.applyToAll) {
            for ($i = 1; $i -le $pres.Slides.Count; $i++) {
                $slide = $pres.Slides.Item($i)
                $slide.FollowMasterBackground = $false
                $slide.Background.Fill.Solid()
                if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
            }
            Output-Json @{ success = $true; data = @{ color = $p.color; appliedTo = $pres.Slides.Count } }
        } else {
            $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
            $slide = $pres.Slides.Item($slideIndex)
            $slide.FollowMasterBackground = $false
            $slide.Background.Fill.Solid()
            if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
            Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; color = $p.color } }
        }
    }

    "setBackgroundImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.FollowMasterBackground = $false
        $slide.Background.Fill.UserPicture($p.path)
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; path = $p.path } }
    }

    "addPptHyperlink" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))
        $actionSettings = $shape.ActionSettings.Item(1)
        $actionSettings.Hyperlink.Address = if ($p.address) { $p.address } else { "" }
        if ($p.subAddress) { $actionSettings.Hyperlink.SubAddress = $p.subAddress }
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name; address = $p.address } }
    }

    "removePptHyperlink" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))
        $shape.ActionSettings.Item(1).Hyperlink.Address = ""
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name } }
    }







    "findPptText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $results = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $text = $shape.TextFrame.TextRange.Text
                        if ($text -and $text.Contains($p.text)) {
                            $results += @{ slideIndex = $i; shapeName = $shape.Name; text = $text }
                        }
                    }
                } catch { Add-WpsWarning $_.Exception.Message }
            }
        }
        Output-Json @{ success = $true; data = @{ searchText = $p.text; results = $results; count = $results.Count } }
    }

    "replacePptText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No target presentation" }; exit }
        # 兼容两种参数名：TS 工具发送 find/replace，旧实现读 findText/replaceText（此前不匹配导致 0 替换）
        $findText = if ($null -ne $p.find) { [string]$p.find } else { [string]$p.findText }
        $replaceText = if ($null -ne $p.replace) { [string]$p.replace } elseif ($null -ne $p.replaceText) { [string]$p.replaceText } else { "" }
        if ([string]::IsNullOrEmpty($findText)) { Output-Json @{ success = $false; error = "find text required" }; exit }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                # 递归处理：文本框 + 表格单元格 + 组合(Group)内所有子形状
                $count += Replace-InShapeTree $shape $findText $replaceText
            }
        }
        Output-Json @{ success = $true; data = @{ findText = $findText; replaceText = $replaceText; count = $count } }
    }

    "startSlideShow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        if ($null -ne $p.fromSlide) {
            $from = [int]$p.fromSlide
            if ($from -lt 1 -or $from -gt $pres.Slides.Count) { Output-Json @{ success = $false; error = "fromSlide is out of range" }; exit }
            # Run() takes no start argument; the slide show settings carry it.
            $pres.SlideShowSettings.StartingSlide = $from
        }
        $pres.SlideShowSettings.Run()
        Output-Json @{ success = $true }
    }

    "endSlideShow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        if ($ppt.SlideShowWindows.Count -gt 0) { $ppt.SlideShowWindows.Item(1).View.Exit() }
        Output-Json @{ success = $true }
    }

    "unifyFont" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $fontName = if ($p.fontName) { $p.fontName } else { "微软雅黑" }
        $includeTitle = if ($null -ne $p.includeTitle) { [bool]$p.includeTitle } else { $true }
        $includeBody = if ($null -ne $p.includeBody) { [bool]$p.includeBody } else { $true }
        $firstSlide = if ($null -ne $p.slideIndex) { [int]$p.slideIndex } else { 1 }
        $lastSlide = if ($null -ne $p.slideIndex) { [int]$p.slideIndex } else { $pres.Slides.Count }
        if (-not (Test-WpsSlideIndex $pres $firstSlide)) { Output-Json @{ success = $false; error = "slideIndex is out of range" }; exit }
        $count = 0
        for ($i = $firstSlide; $i -le $lastSlide; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        # Title placeholders are ppPlaceholderTitle(13)/ppPlaceholderCenterTitle(14).
                        $isTitle = $false
                        try { $phType = [int]$shape.PlaceholderFormat.Type; $isTitle = ($phType -eq 13 -or $phType -eq 14) } catch { $isTitle = $false }
                        if (($isTitle -and $includeTitle) -or ((-not $isTitle) -and $includeBody)) {
                            $shape.TextFrame.TextRange.Font.Name = $fontName
                            $count++
                        }
                    }
                } catch { Add-WpsWarning $_.Exception.Message }
            }
        }
        Output-Json @{ success = $true; data = @{ fontName = $fontName; count = $count } }
    }






    "beautifySlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $schemes = @{
            business = @{ title = 0x2F5496; body = 0x333333 }
            tech = @{ title = 0x00B0F0; body = 0x404040 }
            creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }
            minimal = @{ title = 0x000000; body = 0x666666 }
        }
        $scheme = $schemes[$p.style]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
            $shape = $slide.Shapes.Item($j)
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $textRange = $shape.TextFrame.TextRange
                    if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title }
                    else { $textRange.Font.Color.RGB = $scheme.body }
                    $count++
                }
            } catch { Add-WpsWarning $_.Exception.Message }
        }
        Output-Json @{ success = $true; data = @{ style = $p.style; count = $count } }
    }

    # ==================== dsh-plugin-wps-office-next additions ====================
    # These 10 actions were referenced by upstream tools but never implemented here, so
    # those tools always answered "Unknown action" on Windows.

    "setZoom" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $percent = [int]$p.percent
        if ($percent -lt 10 -or $percent -gt 400) { Output-Json @{ success = $false; error = "percent must be between 10 and 400" }; exit }
        try {
            if ($null -ne $p.sheet) { $wb.Sheets.Item($p.sheet).Activate() | Out-Null }
            $excel.ActiveWindow.Zoom = $percent
            Output-Json @{ success = $true; data = @{ zoom = $excel.ActiveWindow.Zoom } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "setLineSpacing" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $multiple = [double]$p.lineSpacing
        if ($multiple -le 0) { Output-Json @{ success = $false; error = "lineSpacing must be greater than 0" }; exit }
        try {
            if ($null -ne $p.paragraphIndex) {
                $index = [int]$p.paragraphIndex + 1
                if ($index -lt 1 -or $index -gt $doc.Paragraphs.Count) { Output-Json @{ success = $false; error = "paragraphIndex out of range" }; exit }
                $paras = @($doc.Paragraphs.Item($index))
            } else {
                $paras = @($doc.Paragraphs)
            }
            foreach ($para in $paras) {
                $para.Format.LineSpacingRule = 5
                $para.Format.LineSpacing = $multiple * 12
            }
            Output-Json @{ success = $true; data = @{ lineSpacing = $multiple; applied = $paras.Count } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "autoSum" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($null -ne $p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        try {
            $target = $sheet.Range([string]$p.targetCell)
            $target.Formula = "=SUM(" + [string]$p.range + ")"
            Output-Json @{ success = $true; data = @{ targetCell = [string]$p.targetCell; formula = $target.Formula; value = $target.Value2 } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "evaluateFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $formula = [string]$p.formula
        if (-not $formula.StartsWith('=')) { $formula = '=' + $formula }
        try {
            $result = $excel.Evaluate($formula)
            Output-Json @{ success = $true; data = @{ formula = $formula; result = $result; method = 'application-evaluate' } }
        } catch {
            try {
                $wb = $excel.ActiveWorkbook
                if ($null -eq $wb) { throw }
                $probe = $excel.ActiveSheet.Range('XFD1048576')
                $probe.Formula = $formula
                $value = $probe.Value2
                $probe.ClearContents() | Out-Null
                Output-Json @{ success = $true; data = @{ formula = $formula; result = $value; method = 'scratch-cell' } }
            } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
        }
    }

    "setSlideTheme" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $theme = [string]$p.theme
        if (-not $theme) { Output-Json @{ success = $false; error = "theme must be a template path" }; exit }
        if (-not (Test-Path $theme)) {
            Output-Json @{ success = $false; error = ("theme must be an existing .thmx/.potx/.pptx template path; not found: " + $theme) }
            exit
        }
        try {
            $pres.ApplyTemplate($theme) | Out-Null
            Output-Json @{ success = $true; data = @{ applied = $theme } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "setSlideSize" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        try {
            $ps = $pres.PageSetup
            if ($null -ne $p.width) { $ps.SlideWidth = [double]$p.width * 0.75 }
            if ($null -ne $p.height) { $ps.SlideHeight = [double]$p.height * 0.75 }
            Output-Json @{ success = $true; data = @{ slideWidth = $ps.SlideWidth; slideHeight = $ps.SlideHeight; requestedPixels = @{ width = $p.width; height = $p.height } } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "setShapeFill" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $color = Convert-HexColorToRgbInt([string]$p.color)
        if ($null -eq $color) { Output-Json @{ success = $false; error = "color must be a hex value such as #FF0000" }; exit }
        try {
            $slide = $pres.Slides.Item([int]$p.slideIndex)
            $shape = $slide.Shapes.Item([int]$p.shapeIndex)
            $shape.Fill.Visible = -1
            $shape.Fill.Solid() | Out-Null
            $shape.Fill.ForeColor.RGB = $color
            Output-Json @{ success = $true; data = @{ shape = $shape.Name; color = [string]$p.color } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "setFontColor" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "no presentation is open" }; exit }
        $color = Convert-HexColorToRgbInt([string]$p.color)
        if ($null -eq $color) { Output-Json @{ success = $false; error = "color must be a hex value such as #FF0000" }; exit }
        try {
            $slide = $pres.Slides.Item([int]$p.slideIndex)
            $shape = $slide.Shapes.Item([int]$p.shapeIndex)
            $shape.TextFrame.TextRange.Font.Color.RGB = $color
            Output-Json @{ success = $true; data = @{ shape = $shape.Name; color = [string]$p.color } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "setTextColor" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $color = Convert-HexColorToRgbInt([string]$p.color)
        if ($null -eq $color) {
            $nameMap = @{ black = '000000'; white = 'FFFFFF'; red = 'FF0000'; green = '00B050'; blue = '0000FF'; yellow = 'FFFF00'; orange = 'FFA500'; purple = '800080'; gray = '808080'; grey = '808080'; pink = 'FFC0CB'; brown = '8B4513' }
            $key = ([string]$p.color).Trim().ToLower()
            if ($nameMap.ContainsKey($key)) { $color = Convert-HexColorToRgbInt($nameMap[$key]) }
        }
        if ($null -eq $color) { Output-Json @{ success = $false; error = "color must be a hex value such as #FF0000 or a known color name" }; exit }
        try {
            $range = if ($p.range -eq 'all') { $word.ActiveDocument.Content } else { $word.Selection.Range }
            $range.Font.Color = $color
            Output-Json @{ success = $true; data = @{ color = [string]$p.color } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }

    "insertSectionBreak" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $breakMap = @{ nextPage = 2; continuous = 3; evenPage = 4; oddPage = 5 }
        $requested = 'nextPage'
        if ($p.breakType) { $requested = [string]$p.breakType }
        $breakType = $breakMap[$requested]
        if ($null -eq $breakType) { Output-Json @{ success = $false; error = "breakType must be nextPage, continuous, evenPage or oddPage" }; exit }
        try {
            $word.Selection.InsertBreak($breakType)
            Output-Json @{ success = $true; data = @{ breakType = $requested; sections = $doc.Sections.Count } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }
    # Excel has its own find/replace: the shared action above operates on Word only.
    "findReplaceExcel" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($null -ne $p.sheet -and "$($p.sheet)" -ne "") { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $findText = if ($null -ne $p.findText) { [string]$p.findText } else { [string]$p.find }
        $replaceText = if ($null -ne $p.replaceText) { [string]$p.replaceText } else { [string]$p.replace }
        if (-not $findText) { Output-Json @{ success = $false; error = "findText must not be empty" }; exit }
        $matchCase = if ($null -ne $p.matchCase) { [bool]$p.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $p.matchWholeWord) { [bool]$p.matchWholeWord } else { $false }
        $lookAt = if ($matchWholeWord) { 1 } else { 2 }
        try {
            $scope = if ($null -ne $p.range -and "$($p.range)" -ne "") { $sheet.Range([string]$p.range) } else { $sheet.UsedRange }
            # Count matching cells from a single Value2 read. Excel's Find/FindNext binding is
            # unreliable under PowerShell's COM binder cache; reading values is exact and cheap.
            $cells = 0
            $values = $scope.Value2
            if ($null -ne $values) {
                # foreach enumerates every element of a COM 2D array, so this works without
                # depending on Rank/lower-bound behaviour differing between WPS builds.
                $isArray = $false
                try { $isArray = $values.GetType().IsArray } catch { $isArray = $false }
                $items = if ($isArray) { $values } else { @($values) }
                foreach ($v in $items) {
                    if ($null -eq $v) { continue }
                    $s = [string]$v
                    $hit = if ($matchCase) { $s.Contains($findText) } else { $s.IndexOf($findText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 }
                    if ($hit) { $cells++ }
                }
            }
            $didReplace = $false
            if ($cells -gt 0) {
                $didReplace = $scope.Replace([string]$findText, [string]$replaceText, [int]$lookAt, 1, [bool]$matchCase, $false)
            }
            Output-Json @{ success = $true; data = @{ cells = $cells; find = $findText; replace = $replaceText; sheet = $sheet.Name; changed = [bool]$didReplace } }
        } catch { Output-Json @{ success = $false; error = $_.Exception.Message } }
    }
    default {
        Output-Json @{ success = $false; error = "Unknown action: $Action" }
    }
}

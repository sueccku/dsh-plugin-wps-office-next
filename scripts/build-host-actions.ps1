# Input: mcp/scripts/wps-com.ps1 (upstream one-shot COM bridge)
# Output: host/wps-actions.ps1 (dot-sourceable action module exposing Invoke-WpsAction)
# Pos: Reproducible transformation of the one-shot bridge into a resident-host module.
#      Re-run after any upstream re-sync of wps-com.ps1.

param(
    [string]$Source = (Join-Path $PSScriptRoot '..\mcp\scripts\wps-com.ps1'),
    [string]$OutFile = (Join-Path $PSScriptRoot '..\host\wps-actions.ps1')
)

$ErrorActionPreference = 'Stop'

$src = [System.IO.File]::ReadAllText($Source)
# ReadAllText keeps the source BOM as a leading U+FEFF; drop it so we can write exactly one.
$src = $src.TrimStart([char]0xFEFF)

$before = $src
$src = $src -replace '(?s)param\(\s*\[string\]\$Action,\s*\[string\]\$Params\s*=\s*"\{\}"\s*\)', ''
if ($src -eq $before) { throw 'param block not found - upstream layout changed' }

$marker = 'try { $p = $Params | ConvertFrom-Json }'
$idx = $src.IndexOf($marker)
if ($idx -lt 0) { throw 'dispatch boundary not found - upstream layout changed' }
$head = $src.Substring(0, $idx)
$body = $src.Substring($idx)

$beforeHead = $head
$head = $head -replace '(?s)function Output-Json\(\$obj\)\s*\{.*?\}', 'function Output-Json($obj) { $script:WpsResult = $obj }'
if ($head -eq $beforeHead) { throw 'Output-Json not found - upstream layout changed' }

$body = $body -replace '(?m)^\s*exit\s*}\s*$', 'return }'
$body = $body -replace '(?m)^(\s*)exit\s*$', '$1return'
$body = $body -replace '(?m);\s*exit\s*}\s*$', '; return }'
$body = $body -replace '(?m);\s*exit\s*$', '; return'
$body = $body -replace '\{\s*exit\s*\}', '{ return }'

$crlf = [string][char]13 + [string][char]10

# ---- parameter key table -------------------------------------------------------
# Every action reads its parameters as properties of $p, and those property names ARE its contract.
# Extract them from the source rather than keeping a second hand-written list that would drift.
# An action whose parameter access cannot be read statically is left out of the table and thereby
# skips the check, so the guard can never reject a key that is genuinely read.
$helperKeys = @{
    'Resolve-Worksheet'    = @('sheet', 'name', 'oldName')
    'Get-TargetPres'       = @('presentationName')
    'Get-WorksheetByParam' = @('sheet')
}
$dynamicPatterns = @('\$p\.\$', '\$p\[', '\$p\.PSObject', "Get-PropOrNull\s+\`$p\s+(?!')")
$paramKeys = @{}
$dynamicActions = New-Object System.Collections.Generic.List[string]
$caseRx = [regex]'(?m)^ {4}"([A-Za-z][A-Za-z0-9]*)" \{\s*$'
$caseHits = $caseRx.Matches($body)
for ($i = 0; $i -lt $caseHits.Count; $i++) {
    $caseName = $caseHits[$i].Groups[1].Value
    $from = $caseHits[$i].Index
    $to = if ($i + 1 -lt $caseHits.Count) { $caseHits[$i + 1].Index } else { $body.Length }
    $caseText = $body.Substring($from, $to - $from)

    $isDynamic = $false
    foreach ($pattern in $dynamicPatterns) { if ([regex]::IsMatch($caseText, $pattern)) { $isDynamic = $true } }

    $keys = New-Object System.Collections.Generic.HashSet[string]
    foreach ($m in [regex]::Matches($caseText, "\`$p\.'([^']+)'")) { [void]$keys.Add($m.Groups[1].Value) }
    foreach ($m in [regex]::Matches($caseText, '\$p\.([A-Za-z_][A-Za-z0-9_]*)')) { [void]$keys.Add($m.Groups[1].Value) }
    foreach ($m in [regex]::Matches($caseText, "Get-PropOrNull\s+\`$p\s+'([^']+)'")) { [void]$keys.Add($m.Groups[1].Value) }

    # A bare $p handed to one of the shared resolvers contributes that resolver's keys.
    foreach ($line in ($caseText -split "\r?\n")) {
        if (-not [regex]::IsMatch($line, "\`$p(?![.\w'])")) { continue }
        if ([regex]::IsMatch($line, "Get-PropOrNull\s+\`$p\s*'")) { continue }
        $calleeText = $line.Trim()
        $calleeText = $calleeText -replace '^\$[\w:\.\[\]'']+\s*=\s*', ''
        $calleeText = $calleeText -replace '^(if|elseif)\s*\([^)]*\)\s*\{\s*', ''
        $calleeText = $calleeText -replace '^\{\s*', ''
        $calleeMatch = [regex]::Match($calleeText, '^([A-Za-z][A-Za-z0-9_-]*)')
        if (-not $calleeMatch.Success) { continue }
        $callee = $calleeMatch.Groups[1].Value
        if ($helperKeys.ContainsKey($callee)) { foreach ($k in $helperKeys[$callee]) { [void]$keys.Add($k) } }
        elseif ($callee -ne 'return') { $isDynamic = $true }
    }

    if ($isDynamic) { [void]$dynamicActions.Add($caseName) }
    else { $paramKeys[$caseName] = @($keys | Sort-Object) }
}
$tableLines = foreach ($k in ($paramKeys.Keys | Sort-Object)) {
    $vals = ($paramKeys[$k] | ForEach-Object { "'$_'" }) -join ', '
    ("    '" + $k + "' = @(" + $vals + ")")
}
$keyTable = '# Accepted parameter names per action, derived from the switch below by the generator.' + $crlf +
            '$script:ActionParamKeys = @{' + $crlf + ($tableLines -join $crlf) + $crlf + '}' + $crlf + $crlf

# Reject parameters the action does not read, right after $p is materialised.
$guard = '    $__paramError = Test-WpsActionParamKeys $Action $p $script:ActionParamKeys[$Action]' + $crlf +
         '    if ($null -ne $__paramError) { Output-Json @{ success = $false; error = $__paramError }; return }' + $crlf
$guardBoundary = 'try { $p = $Params | ConvertFrom-Json }'
if ($body.IndexOf($guardBoundary) -lt 0) { throw 'parameter guard boundary not found - upstream layout changed' }
$firstNl = $body.IndexOf($crlf)
$body = $body.Substring(0, $firstNl + $crlf.Length) + $guard + $body.Substring($firstNl + $crlf.Length)

$header = '# GENERATED by scripts/build-host-actions.ps1 - do not edit by hand.' + $crlf +
          '# Source: mcp/scripts/wps-com.ps1 (upstream lc2panda/wps-skills).' + $crlf + $crlf
$wrapperOpen = 'function Invoke-WpsAction {' + $crlf + '    param([string]$Action, [string]$Params = ''{}'')' + $crlf + $crlf
$module = $header + $head + $keyTable + $wrapperOpen + $body + $crlf + '}' + $crlf

$enc = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutFile, $module, $enc)

$bytes = [System.IO.File]::ReadAllBytes($OutFile)
$bom = ($bytes[0..2] | ForEach-Object { $_.ToString('X2') }) -join ' '
$text = [System.IO.File]::ReadAllText($OutFile)
$exitLeft = ([regex]::Matches($text, '\bexit\b')).Count
$cases = ([regex]::Matches($text, '(?m)^\s{4}"[A-Za-z]')).Count
$funcs = ([regex]::Matches($text, '(?m)^function ')).Count
$invoke = ([regex]::Matches($text, 'function Invoke-WpsAction')).Count
$paramsInHead = ([regex]::Matches($head, '\$Action\b|\$Params\b')).Count

'generated: ' + $OutFile
'bytes=' + $bytes.Length + ' bom=' + $bom
'exit_remaining=' + $exitLeft + ' switch_cases=' + $cases + ' functions=' + $funcs + ' invokeDefs=' + $invoke
'head_refs_Action_or_Params=' + $paramsInHead
'param_keys=' + $paramKeys.Count + ' dynamic=' + $dynamicActions.Count
if ($dynamicActions.Count -gt 0) { 'param_keys_skipped=' + (($dynamicActions | Sort-Object) -join ',') }
'guard_installed=' + ([regex]::Matches($text, 'Test-WpsActionParamKeys \$Action')).Count

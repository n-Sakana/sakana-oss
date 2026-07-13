# test/test-actions.ps1 -- stable action IDs and shared launch core.
. (Join-Path $PSScriptRoot "_assert.ps1")
. (Join-Path $PSScriptRoot "..\common\install.ps1")

$root = Split-Path $PSScriptRoot -Parent
$core = Join-Path $root "common\launch-core.ps1"
$launch = Join-Path $root "common\launch.ps1"
Assert-True (Test-Path -LiteralPath $core -PathType Leaf) "shared launch core exists"
if (-not (Test-Path -LiteralPath $core -PathType Leaf)) { Exit-Test }

$dotOutput = @(. $core)
Assert-True ($dotOutput.Count -eq 0) "dot-sourcing launch core has no output"

$expected = [ordered]@{
    "capture" = @("default")
    "clip" = @("codex", "codex-resume", "claude", "claude-resume")
    "keysend" = @("default", "custom")
    "md-extract" = @("default", "timeout-300", "custom")
    "md-mirror" = @("create", "restore-file", "restore-clipboard", "primer")
    "md-patch" = @("apply-clipboard", "apply-file")
    "timer" = @("5-min", "10-min", "25-min", "custom")
    "timestamp" = @("date", "datetime")
    "transcribe" = @("start")
    "tree" = @("depth-3", "depth-5", "unlimited", "custom")
    "vba-devkit" = @("analyze", "analyze-settings", "extract", "diff", "sanitize", "unlock")
}

foreach ($toolId in $expected.Keys) {
    $toolDir = Join-Path $root ("tool\" + $toolId)
    $read = Read-Manifest $toolDir
    Assert-True $read.Ok ("manifest parses: " + $toolId)
    if (-not $read.Ok) { continue }
    $errors = @(Test-Manifest $read.Data $toolId $toolDir)
    Assert-True ($errors.Count -eq 0) ("manifest is valid: " + $toolId)
    $info = Get-ManifestActionInfo $read.Data
    Assert-True $info.SupportsActions ("actions supported: " + $toolId)
    Assert-True ((@($info.ActionIds) -join ",") -ceq (@($expected[$toolId]) -join ",")) ("stable action IDs: " + $toolId)
}

$fx = Join-Path $env:TEMP ("toolrack_actions_" + [guid]::NewGuid().ToString("N"))
$tool = Join-Path $fx "action-tool"
New-Item -ItemType Directory -Force -Path $tool | Out-Null
try {
    [IO.File]::WriteAllText((Join-Path $tool "main.ps1"), @'
param([string]$Target, [string]$Value = "")
[IO.File]::WriteAllText($env:TOOLRACK_ACTION_MARKER, "T=" + $Target + " V=" + $Value)
exit 9
'@)

    $manifest = @'
{"schema":1,"id":"action-tool","name":"Action Tool","on":["background"],"run":{"type":"powershell","entry":"main.ps1","keep_open":false},"variants":[{"id":"alpha","label":"Alpha","args":["-Value","A"]},{"id":"beta","label":"Beta","args":["-Value","B"]}]}
'@
    [IO.File]::WriteAllText((Join-Path $tool "tool.json"), $manifest)

    $resolved = Resolve-LaunchPlan -ToolDir $tool -VariantIndex -1 -VariantSpecified $false -ActionId "beta" -ActionSpecified $true
    Assert-True $resolved.Ok "action resolves"
    Assert-True ((@($resolved.Args) -join ",") -ceq "-Value,B") "action resolves expected args"

    $reordered = @'
{"schema":1,"id":"action-tool","name":"Action Tool","on":["background"],"run":{"type":"powershell","entry":"main.ps1","keep_open":false},"variants":[{"id":"beta","label":"Beta","args":["-Value","B"]},{"id":"alpha","label":"Alpha","args":["-Value","A"]}]}
'@
    [IO.File]::WriteAllText((Join-Path $tool "tool.json"), $reordered)
    $resolved = Resolve-LaunchPlan -ToolDir $tool -VariantIndex -1 -VariantSpecified $false -ActionId "beta" -ActionSpecified $true
    Assert-True ($resolved.Ok -and (@($resolved.Args) -join ",") -ceq "-Value,B") "action survives variant reorder"

    $flatDir = Join-Path $fx "flat-tool"
    New-Item -ItemType Directory -Force -Path $flatDir | Out-Null
    [IO.File]::WriteAllText((Join-Path $flatDir "main.ps1"), "param([string]`$Target)`nexit 0`n")
    [IO.File]::WriteAllText((Join-Path $flatDir "tool.json"), '{"schema":1,"id":"flat-tool","name":"Flat","on":["background"],"run":{"type":"powershell","entry":"main.ps1","keep_open":false}}')
    $flatRead = Read-Manifest $flatDir
    $flatInfo = Get-ManifestActionInfo $flatRead.Data
    Assert-True ($flatInfo.SupportsActions -and (@($flatInfo.ActionIds) -join ",") -ceq "default") "flat tool has implicit default action"
    $flatDefault = Resolve-LaunchPlan -ToolDir $flatDir -VariantIndex -1 -VariantSpecified $false -ActionId "default" -ActionSpecified $true
    Assert-True $flatDefault.Ok "flat default action resolves"
    $flatBad = Resolve-LaunchPlan -ToolDir $flatDir -VariantIndex -1 -VariantSpecified $false -ActionId "other" -ActionSpecified $true
    Assert-True (-not $flatBad.Ok) "flat non-default action is rejected"

    $invalidCases = @(
        @('{"schema":1,"id":"action-tool","name":"X","on":["background"],"run":{"type":"powershell","entry":"main.ps1"},"variants":[{"id":"Bad_ID","label":"A","args":[]}]}', "invalid action ID"),
        @('{"schema":1,"id":"action-tool","name":"X","on":["background"],"run":{"type":"powershell","entry":"main.ps1"},"variants":[{"id":"same","label":"A","args":[]},{"id":"same","label":"B","args":[]}]}', "duplicate action ID"),
        @('{"schema":1,"id":"action-tool","name":"X","on":["background"],"run":{"type":"powershell","entry":"main.ps1"},"variants":[{"id":"one","label":"A","args":[]},{"label":"B","args":[]}]}', "mixed action IDs")
    )
    foreach ($case in $invalidCases) {
        [IO.File]::WriteAllText((Join-Path $tool "tool.json"), [string]$case[0])
        $read = Read-Manifest $tool
        $errors = @(Test-Manifest $read.Data "action-tool" $tool)
        Assert-True ($errors.Count -gt 0) ([string]$case[1] + " is rejected")
    }

    $legacy = '{"schema":1,"id":"action-tool","name":"Legacy","on":["background"],"run":{"type":"powershell","entry":"main.ps1","keep_open":false},"variants":[{"label":"A","args":["-Value","A"]}]}'
    [IO.File]::WriteAllText((Join-Path $tool "tool.json"), $legacy)
    $legacyRead = Read-Manifest $tool
    $legacyErrors = @(Test-Manifest $legacyRead.Data "action-tool" $tool)
    $legacyInfo = Get-ManifestActionInfo $legacyRead.Data
    Assert-True ($legacyErrors.Count -eq 0) "legacy manifest without IDs remains valid"
    Assert-True (-not $legacyInfo.SupportsActions -and @($legacyInfo.ActionIds).Count -eq 0) "legacy variant tool cannot be bound"
    $legacyAction = Resolve-LaunchPlan -ToolDir $tool -VariantIndex -1 -VariantSpecified $false -ActionId "alpha" -ActionSpecified $true
    Assert-True (-not $legacyAction.Ok) "legacy variant tool rejects action route"

    [IO.File]::WriteAllText((Join-Path $tool "tool.json"), $manifest)
    $marker = Join-Path $fx "marker.txt"
    $env:TOOLRACK_ACTION_MARKER = $marker
    $env:TOOLRACK_NOPAUSE = "1"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $tool -Action beta -Target "C:\action target"
    Assert-True ($LASTEXITCODE -eq 9) "launch -Action preserves tool exit code"
    Assert-True ([IO.File]::ReadAllText($marker) -eq "T=C:\action target V=B") "launch -Action passes target and args"

    Remove-Item -LiteralPath $marker -Force
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $tool -Action beta -Variant 0 -Target "C:\x"
    Assert-True ($LASTEXITCODE -eq 1 -and -not (Test-Path -LiteralPath $marker)) "Action and Variant together are rejected"

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $tool -Action missing -Target "C:\x"
    Assert-True ($LASTEXITCODE -eq 1 -and -not (Test-Path -LiteralPath $marker)) "unknown action does not run tool"

    $parseErrors = $null
    $tokens = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($core, [ref]$tokens, [ref]$parseErrors)
    $exitNodes = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.ExitStatementAst] }, $true))
    Assert-True (@($parseErrors).Count -eq 0) "launch core parses under PowerShell"
    Assert-True ($exitNodes.Count -eq 0) "launch core contains no exit statement"
    $coreText = [IO.File]::ReadAllText($core)
    Assert-True ($coreText -notmatch 'MessageBox|ReadKey') "launch core owns no UI or pause"
} finally {
    $env:TOOLRACK_ACTION_MARKER = $null
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test

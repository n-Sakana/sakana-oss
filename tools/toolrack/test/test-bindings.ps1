# test/test-bindings.ps1 -- strict global binding parser and resolver.
. (Join-Path $PSScriptRoot "_assert.ps1")

$root = Split-Path $PSScriptRoot -Parent
$library = Join-Path $root "common\bindings.ps1"
$resolverScript = Join-Path $root "common\resolve-host-config.ps1"
$defaultConfig = Join-Path $root "bindings.json"
foreach ($path in @($library, $resolverScript, $defaultConfig)) {
    Assert-True (Test-Path -LiteralPath $path -PathType Leaf) ("binding asset exists: " + (Split-Path $path -Leaf))
}
if (-not (Test-Path -LiteralPath $library -PathType Leaf) -or
    -not (Test-Path -LiteralPath $resolverScript -PathType Leaf) -or
    -not (Test-Path -LiteralPath $defaultConfig -PathType Leaf)) {
    Exit-Test
}

$dotOutput = @(. $library)
Assert-True ($dotOutput.Count -eq 0) "dot-sourcing bindings library has no output"

$fx = Join-Path $env:TEMP ("toolrack_bindings_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $fx | Out-Null
$preserveKey = "HKCU:\Software\Classes\Directory\Background\shell\ToolRack\shell\__binding-preserve-test"
try {
    function Test-ConfigText {
        param([string]$Text)
        $path = Join-Path $fx ("config_" + [guid]::NewGuid().ToString("N") + ".json")
        [IO.File]::WriteAllText($path, $Text)
        return Read-BindingsConfig -Path $path
    }

    $validText = @'
{"schema":1,"bindings":[{"id":"capture-hotkey","trigger":{"type":"hotkey","key":"C","modifiers":["alt","ctrl"]},"invoke":{"tool":"capture","action":"default"}},{"id":"capture-mouse","trigger":{"type":"mouse","button":"right","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}},{"id":"tree-hotkey","trigger":{"type":"hotkey","key":"F13","modifiers":["ctrl","shift"]},"invoke":{"tool":"tree","action":"depth-3"}}]}
'@
    $valid = Test-ConfigText $validText
    Assert-True $valid.Ok "valid global bindings parse"
    Assert-True (@($valid.Data.Bindings).Count -eq 3) "binding array is preserved"
    Assert-True ((@($valid.Data.Bindings[0].Trigger.Modifiers) -join ",") -ceq "ctrl,alt") "modifiers are normalized"
    Assert-True ((Format-Trigger $valid.Data.Bindings[0].Trigger) -eq "hotkey:ctrl+alt+C") "hotkey format is stable"
    Assert-True ((Format-Trigger $valid.Data.Bindings[1].Trigger) -eq "mouse:ctrl+right") "mouse format is stable"

    $invalidCases = [ordered]@{
        "malformed JSON" = '{'
        "root scalar" = '[]'
        "wrong schema type" = '{"schema":"1","bindings":[]}'
        "unsupported schema" = '{"schema":2,"bindings":[]}'
        "bindings scalar" = '{"schema":1,"bindings":{}}'
        "unknown root field" = '{"schema":1,"bindings":[],"extra":1}'
        "duplicate binding id" = '{"schema":1,"bindings":[{"id":"same","trigger":{"type":"hotkey","key":"A","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}},{"id":"same","trigger":{"type":"hotkey","key":"B","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "control character id" = "{`"schema`":1,`"bindings`":[{`"id`":`"bad``nid`",`"trigger`":{`"type`":`"hotkey`",`"key`":`"A`",`"modifiers`":[`"ctrl`"]},`"invoke`":{`"tool`":`"capture`",`"action`":`"default`"}}]}"
        "unknown binding field" = '{"schema":1,"bindings":[{"id":"a","extra":1,"trigger":{"type":"hotkey","key":"A","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "when field" = '{"schema":1,"bindings":[{"id":"a","when":{"process":["Code.exe"]},"trigger":{"type":"hotkey","key":"A","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "process field" = '{"schema":1,"bindings":[{"id":"a","process":["Code.exe"],"trigger":{"type":"hotkey","key":"A","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "arbitrary command" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"A","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default","command":"calc.exe"}}]}'
        "unknown trigger" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"key","key":"A","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "no modifier hotkey" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"A","modifiers":[]},"invoke":{"tool":"capture","action":"default"}}]}'
        "win modifier" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"A","modifiers":["win"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "duplicate modifier" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"A","modifiers":["ctrl","ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "lowercase key" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"a","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "F12" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"F12","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "unsupported key" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"ESC","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "plain right click" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"mouse","button":"right","modifiers":[]},"invoke":{"tool":"capture","action":"default"}}]}'
        "left mouse" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"mouse","button":"left","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}}]}'
        "empty tool" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"A","modifiers":["ctrl"]},"invoke":{"tool":"","action":"default"}}]}'
        "invalid action" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"A","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"Bad_ID"}}]}'
        "duplicate normalized trigger" = '{"schema":1,"bindings":[{"id":"a","trigger":{"type":"hotkey","key":"A","modifiers":["ctrl","alt"]},"invoke":{"tool":"capture","action":"default"}},{"id":"b","trigger":{"type":"hotkey","key":"A","modifiers":["alt","ctrl"]},"invoke":{"tool":"tree","action":"depth-3"}}]}'
    }
    foreach ($name in $invalidCases.Keys) {
        $result = Test-ConfigText ([string]$invalidCases[$name])
        Assert-True (-not $result.Ok -and @($result.Errors).Count -gt 0) ($name + " is rejected")
    }

    $tools = @(
        @{ Id = "capture"; SupportsActions = $true; ActionIds = @("default"); Dir = "C:\capture" },
        @{ Id = "tree"; SupportsActions = $true; ActionIds = @("depth-3", "custom"); Dir = "C:\tree" }
    )
    $resolved = Resolve-Bindings -Config $valid.Data -Tools $tools
    Assert-True $resolved.Ok "valid bindings resolve"
    Assert-True (@($resolved.Active).Count -eq 3 -and @($resolved.Rejected).Count -eq 0) "all valid targets are active"

    $isolationText = @'
{"schema":1,"bindings":[{"id":"good","trigger":{"type":"hotkey","key":"C","modifiers":["ctrl"]},"invoke":{"tool":"capture","action":"default"}},{"id":"missing-tool","trigger":{"type":"hotkey","key":"T","modifiers":["ctrl"]},"invoke":{"tool":"absent","action":"default"}},{"id":"missing-action","trigger":{"type":"hotkey","key":"F13","modifiers":["ctrl"]},"invoke":{"tool":"tree","action":"absent"}}]}
'@
    $isolation = Test-ConfigText $isolationText
    $resolved = Resolve-Bindings -Config $isolation.Data -Tools $tools
    Assert-True ($resolved.Ok -and @($resolved.Active).Count -eq 1 -and @($resolved.Rejected).Count -eq 2) "missing targets reject only their binding"
    Assert-True ($resolved.Active[0].Invoke.Action -ceq "default") "resolver does not guess another action"

    $default = Read-BindingsConfig -Path $defaultConfig
    Assert-True ($default.Ok -and @($default.Data.Bindings).Count -eq 3) "default bindings parse"

    $invalidInstallConfig = Join-Path $fx "invalid-install-bindings.json"
    [IO.File]::WriteAllText($invalidInstallConfig, '{"schema":1,"bindings":{}}')
    New-Item -Path $preserveKey -Force | Out-Null
    New-ItemProperty -Path $preserveKey -Name "Marker" -Value "keep" -PropertyType String -Force | Out-Null
    $installOutput = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "common\install.ps1") -BindingsConfigPath $invalidInstallConfig 2>&1)
    Assert-True ($LASTEXITCODE -eq 1) "invalid binding config stops install"
    Assert-Contains @($installOutput | ForEach-Object { [string]$_ }) "*bindings.json is invalid*" "install reports binding validation failure"
    $preserved = $null
    try { $preserved = (Get-ItemProperty -Path $preserveKey -Name "Marker" -ErrorAction Stop).Marker } catch {}
    Assert-True ($preserved -eq "keep") "invalid bindings preserve existing Explorer registry"

    $fixtureRoot = Join-Path $fx "resolver-root"
    New-Item -ItemType Directory -Force -Path (Join-Path $fixtureRoot "tool\capture") | Out-Null
    [IO.File]::WriteAllText((Join-Path $fixtureRoot "tool\capture\main.ps1"), "param([string]`$Target)`n")
    [IO.File]::WriteAllText((Join-Path $fixtureRoot "tool\capture\tool.json"), '{"schema":1,"id":"capture","name":"Capture","on":["background"],"run":{"type":"powershell","entry":"main.ps1","window":"gui"}}')
    [IO.File]::WriteAllText((Join-Path $fixtureRoot "bindings.json"), '{"schema":1,"bindings":[{"id":"capture-hotkey","trigger":{"type":"hotkey","key":"C","modifiers":["ctrl","alt"]},"invoke":{"tool":"capture","action":"default"}}]}')
    $outputPath = Join-Path $fx "resolved.json"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $resolverScript -Root $fixtureRoot -OutputPath $outputPath
    Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $outputPath -PathType Leaf)) "resolver writes caller output"
    if (Test-Path -LiteralPath $outputPath -PathType Leaf) {
        $output = ConvertFrom-Json ([IO.File]::ReadAllText($outputPath))
        Assert-True ($output.schema -eq 1 -and @($output.active).Count -eq 1) "resolved schema contains active binding"
        Assert-True ([string]$output.sourceConfigSha256 -match '^[A-Fa-f0-9]{64}$') "resolved config records source SHA-256"
        Assert-True ([IO.Path]::GetFullPath([string]$output.root) -eq [IO.Path]::GetFullPath($fixtureRoot)) "resolved config records root"
    }
} finally {
    Remove-Item -Path $preserveKey -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test

# test/test-menu-layout.ps1 -- menu.json validation and 16-slot page layout
. (Join-Path $PSScriptRoot "_assert.ps1")
. (Join-Path $PSScriptRoot "..\common\install.ps1")

foreach ($name in @("Read-MenuConfig", "Resolve-MenuLayout")) {
    if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
        Assert-True $false ("function exists: " + $name)
    }
}
if ($script:Failed -gt 0) { Exit-Test }

$tmp = Join-Path $env:TEMP ("toolrack_menu_" + [guid]::NewGuid().ToString("N") + ".json")
try {
    $json = @'
{
  "schema": 1,
  "default_category": "general",
  "categories": [
    { "id": "general", "label": "Tool Rack", "tools": ["gen-a", "gen-b", "missing"] },
    { "id": "ai", "label": "Tool Rack AI", "tools": ["ai-a", "too-big"] }
  ]
}
'@
    [System.IO.File]::WriteAllText($tmp, $json)
    $read = Read-MenuConfig $tmp
    Assert-True $read.Ok "valid menu.json accepted"
    Assert-True ($read.Data.DefaultCategory -eq "general") "default category parsed"
    Assert-True (@($read.Data.Categories).Count -eq 2) "category array preserved"

    $tools = @(
        @{ Id="gen-a"; Name="A"; On=@("background"); Window="console"; VariantLabels=@(1..10); Dir="C:\x\a" },
        @{ Id="gen-b"; Name="B"; On=@("background"); Window="console"; VariantLabels=@(1..5); Dir="C:\x\b" },
        @{ Id="ai-a"; Name="AI"; On=@("background"); Window="console"; VariantLabels=@(1..15); Dir="C:\x\ai" },
        @{ Id="too-big"; Name="Big"; On=@("background"); Window="console"; VariantLabels=@(1..16); Dir="C:\x\big" },
        @{ Id="unlisted"; Name="U"; On=@("background"); Window="console"; VariantLabels=@(); Dir="C:\x\u" }
    )
    $layout = Resolve-MenuLayout $tools $read.Data
    Assert-True $layout.Ok "layout succeeds"
    Assert-Contains @($layout.Warnings) "*missing*not found*" "missing configured tool warns"
    Assert-Contains @($layout.Warnings) "*unlisted*default category*" "unlisted tool falls back with warning"
    Assert-Contains @($layout.Rejected) "*too-big*17*" "single oversized tool rejected"

    $bg = @($layout.Pages | Where-Object { $_.Context -eq "background" })
    Assert-True ($bg.Count -eq 3) "background creates two general pages and one AI page"
    Assert-True ($bg[0].KeyName -eq "ToolRack") "default first page keeps legacy key"
    Assert-True ($bg[0].Label -eq "Tool Rack") "default first page label"
    Assert-True ($bg[0].Cost -eq 11 -and @($bg[0].Tools).Count -eq 1) "whole tool stays on first page"
    Assert-True ($bg[1].KeyName -eq "ToolRack.general.p2") "default overflow key"
    Assert-True ($bg[1].Label -eq "Tool Rack 2") "overflow page is visibly numbered"
    Assert-True ($bg[1].Cost -eq 7 -and @($bg[1].Tools).Count -eq 2) "remaining tools packed without splitting"
    Assert-True ($bg[2].KeyName -eq "ToolRack.ai") "named category key"
    Assert-True ($bg[2].Cost -eq 16) "exactly 16 slots is accepted"
    Assert-True (@($layout.Pages | Where-Object { $_.Cost -gt 16 }).Count -eq 0) "no page exceeds shell limit"

    Remove-Item -LiteralPath $tmp -Force
    [System.IO.File]::WriteAllText($tmp,
        '{"schema":1,"default_category":"a","categories":[{"id":"a","label":"A","tools":["x"]},{"id":"b","label":"B","tools":["x"]}]}')
    $bad = Read-MenuConfig $tmp
    Assert-True (-not $bad.Ok) "tool cannot be assigned to two categories"
    Assert-Contains @($bad.Errors) "*assigned more than once*" "duplicate assignment error is explicit"

    Remove-Item -LiteralPath $tmp -Force
    $fallback = Read-MenuConfig $tmp
    Assert-True ($fallback.Ok -and -not $fallback.Exists) "missing menu.json enables fallback"
    $fallbackLayout = Resolve-MenuLayout $tools $fallback.Data
    Assert-True (@($fallbackLayout.Pages | Where-Object { $_.Context -eq "background" }).Count -ge 2) "fallback also auto-pages"
} finally {
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
Exit-Test

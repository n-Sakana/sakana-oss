# test/test-clip.ps1 -- clip copies the configured snippet text to the clipboard
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$main = Join-Path $root "tool\clip\main.ps1"

# manifest is valid under the toolrack contract
. (Join-Path $PSScriptRoot "..\common\install.ps1")
$r = Read-Manifest (Join-Path $root "tool\clip")
Assert-True $r.Ok "clip tool.json parses"
$e = @(Test-Manifest $r.Data "clip" (Join-Path $root "tool\clip"))
Assert-True ($e.Count -eq 0) ("clip manifest valid (" + ($e -join '; ') + ")")

# each snippet key lands the exact configured text on the clipboard
$cases = @(
    @{ key = "codex";         expect = "codex --dangerously-bypass-approvals-and-sandbox" },
    @{ key = "codex-resume";  expect = "codex resume --dangerously-bypass-approvals-and-sandbox" },
    @{ key = "claude";        expect = "claude --dangerously-skip-permissions" },
    @{ key = "claude-resume"; expect = "claude --resume --dangerously-skip-permissions" }
)
foreach ($c in $cases) {
    Set-Clipboard -Value "SENTINEL-BEFORE"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Key $c.key | Out-Null
    $got = [string](Get-Clipboard -Raw)
    Assert-True ($got.Trim() -eq $c.expect) ("clip " + $c.key + " copied exact text")
}

Exit-Test

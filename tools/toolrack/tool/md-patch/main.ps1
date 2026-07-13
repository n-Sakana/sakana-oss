# MD Patch -- apply a strict AI patch to the selected folder.
param(
    [string]$Target = "$PWD",
    [ValidateSet("clipboard", "file")][string]$Source = "clipboard",
    [string]$InputFile = "",
    [switch]$NoOpen
)
. (Join-Path $PSScriptRoot "..\..\common\ui.ps1")
. (Join-Path $PSScriptRoot "format.ps1")
. (Join-Path $PSScriptRoot "engine.ps1")

function Get-PatchClipboardText {
    for ($attempt = 0; $attempt -lt 4; $attempt++) {
        try {
            $value = Get-Clipboard -Raw -Format Text -ErrorAction Stop
            if ([string]::IsNullOrEmpty($value)) { throw "Clipboard does not contain text." }
            return [string]$value
        } catch {
            if ($attempt -eq 3) { throw }
            Start-Sleep -Milliseconds 100
        }
    }
}

Start-UI -Title "MD Patch"
try {
    if (-not (Test-Path -LiteralPath $Target -PathType Container)) { throw "Target folder not found: $Target" }
    Write-Step "Reading patch"
    if ($Source -eq "clipboard") {
        $text = Get-PatchClipboardText
    } else {
        if (-not $InputFile) { $InputFile = Read-Value -Prompt "Patch MD file" }
        if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) { throw "Patch file not found: $InputFile" }
        $bytes = [IO.File]::ReadAllBytes($InputFile)
        $text = (New-PatchUtf8Encoding $false).GetString($bytes)
    }
    $patch = Read-PatchDocument $text
    Write-Step "Preflight"
    $prepared = Prepare-PatchApplication $patch $Target
    foreach ($action in $prepared.Actions) { Write-Dim ("{0,-17} {1}" -f $action.Type, $action.Path) }
    if ($env:TOOLRACK_NOPAUSE -ne "1") {
        if (-not (Confirm-Yes -Prompt ("Apply {0} operation(s)?" -f $prepared.Actions.Count) -Default $false)) {
            Write-Warn "Cancelled."
            Stop-UI
            exit 0
        }
    }
    Write-Step "Backing up and applying"
    $backupRoot = Join-Path (Get-ToolrackRoot) "output"
    $result = Apply-PreparedPatch $prepared $backupRoot
    Write-Ok ("Applied {0} operation(s)" -f $result.Applied)
    Stop-UI -OutPath $result.BackupPath
    if (-not $NoOpen) { Open-Folder $result.BackupPath }
    exit 0
} catch {
    Write-Err $_.Exception.Message
    Stop-UI
    exit 1
}

# MD Mirror -- create a mirror MD or restore one into a new output folder.
param(
    [string]$Target = "$PWD",
    [ValidateSet("create", "restore-file", "restore-clipboard", "primer")][string]$Mode = "create",
    [switch]$NoOpen
)
. (Join-Path $PSScriptRoot "..\..\common\ui.ps1")
. (Join-Path $PSScriptRoot "format.ps1")

function Get-MirrorClipboardText {
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

function Set-MirrorClipboardText {
    param([string]$Text)
    for ($attempt = 0; $attempt -lt 4; $attempt++) {
        try {
            Set-Clipboard -Value $Text -ErrorAction Stop
            return
        } catch {
            if ($attempt -eq 3) { throw }
            Start-Sleep -Milliseconds 100
        }
    }
}

Start-UI -Title "MD Mirror"
try {
    $root = Get-ToolrackRoot
    $outputRoot = Join-Path $root "output"
    if ($Mode -eq "create") {
        if (-not (Test-Path -LiteralPath $Target)) { throw "Path not found: $Target" }
        Write-Step "Reading source"
        $package = New-MirrorPackage -Target $Target
        $label = (Get-Item -LiteralPath $Target).Name
        $out = Get-OutputPath -Tool "md-mirror" -Label $label -Ext "md"
        [System.IO.File]::WriteAllText($out, $package.Document, (New-MirrorUtf8Encoding $false))
        Write-Ok ("{0} entries" -f $package.Entries.Count)
        Write-Dim ("{0:N0} characters" -f $package.Document.Length)
        foreach ($warning in $package.Warnings) { Write-Warn $warning }
        Stop-UI -OutPath $out
        if (-not $NoOpen) { Open-Folder (Split-Path $out -Parent) }
        exit 0
    }

    if ($Mode -eq "primer") {
        $primerPath = Join-Path $PSScriptRoot "primer.txt"
        $primerBytes = [System.IO.File]::ReadAllBytes($primerPath)
        $primer = (New-MirrorUtf8Encoding $false).GetString($primerBytes)
        Set-MirrorClipboardText $primer
        Write-Ok "Format instructions copied to the clipboard."
        Stop-UI
        exit 0
    }

    Write-Step "Reading mirror"
    if ($Mode -eq "restore-file") {
        if (-not (Test-Path -LiteralPath $Target -PathType Leaf)) { throw "MD file not found: $Target" }
        if ([System.IO.Path]::GetExtension($Target) -cne ".md") { throw "Target is not an .md file: $Target" }
        $inputBytes = [System.IO.File]::ReadAllBytes($Target)
        $text = (New-MirrorUtf8Encoding $false).GetString($inputBytes)
    } else {
        $text = Get-MirrorClipboardText
    }
    $mirror = Read-MirrorDocument -Text $text
    $destination = Get-MirrorRestoreDestination -OutputRoot $outputRoot
    Write-Step "Restoring into a new output folder"
    $result = Restore-MirrorDocument -Mirror $mirror -Destination $destination
    Write-Ok ("{0} files, {1} directories" -f $result.RestoredFiles, $result.RestoredDirectories)
    foreach ($path in $result.Skipped) { Write-Warn ("Not restored: " + $path) }
    Stop-UI -OutPath $result.Destination
    if (-not $NoOpen) { Open-Folder $result.Destination }
    exit 0
} catch {
    Write-Err $_.Exception.Message
    Stop-UI
    exit 1
}

param(
    [string]$Root = ".",
    [string]$OutFile = "",
    [int64]$MaxFileBytes = 1048576,

    [string[]]$ExcludeDirs = @(
        ".git", ".vs", ".idea",
        "bin", "obj",
        "node_modules", "packages",
        "dist", "build", "out", "target",
        ".venv", "venv", "__pycache__"
    ),

    [string[]]$ExcludeFilePatterns = @(
        ".env", ".env.*",
        "*.key", "*.pem", "*.pfx", "*.p12",
        "id_rsa*", "id_ed25519*",
        "*.user", "*.suo",
        "*-repo-dump-*.txt"
    )
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$RootFull = [System.IO.Path]::GetFullPath($Root)

if (-not (Test-Path -LiteralPath $RootFull -PathType Container)) {
    throw "Root directory not found: $RootFull"
}

if ([string]::IsNullOrWhiteSpace($OutFile)) {
    $repoName = Split-Path -Leaf $RootFull
    $parent = Split-Path -Parent $RootFull
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $OutFile = Join-Path $parent "$repoName-repo-dump-$stamp.txt"
}

$OutFull = [System.IO.Path]::GetFullPath($OutFile)

$BinaryExtensions = @(
    ".exe", ".dll", ".pdb", ".lib", ".obj",
    ".zip", ".7z", ".rar", ".tar", ".gz",
    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".ico",
    ".pdf", ".docx", ".xlsx", ".pptx",
    ".sqlite", ".db"
)

function Test-ExcludedDir {
    param([System.IO.DirectoryInfo]$Item)

    if (($Item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        return $true
    }

    foreach ($name in $ExcludeDirs) {
        if ($Item.Name -ieq $name) {
            return $true
        }
    }

    return $false
}

function Test-ExcludedFile {
    param([System.IO.FileInfo]$Item)

    if (($Item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        return $true
    }

    if ([string]::Equals(
        [System.IO.Path]::GetFullPath($Item.FullName),
        $OutFull,
        [System.StringComparison]::OrdinalIgnoreCase
    )) {
        return $true
    }

    foreach ($pattern in $ExcludeFilePatterns) {
        if ($Item.Name -like $pattern) {
            return $true
        }
    }

    return $false
}

function Get-RelativePathForOutput {
    param([string]$Path)

    $base = $RootFull
    if (-not $base.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $base = $base + [System.IO.Path]::DirectorySeparatorChar
    }

    $baseUri = New-Object System.Uri -ArgumentList $base
    $pathUri = New-Object System.Uri -ArgumentList ([System.IO.Path]::GetFullPath($Path))
    $rel = $baseUri.MakeRelativeUri($pathUri).ToString()

    return [System.Uri]::UnescapeDataString($rel)
}

function Test-BinaryByExtension {
    param([System.IO.FileInfo]$Item)

    $ext = $Item.Extension.ToLowerInvariant()
    return $BinaryExtensions -contains $ext
}

function Test-ContainsNullByte {
    param([string]$Path)

    $fs = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite
    )

    try {
        $len = [Math]::Min(8192, [int]$fs.Length)
        $buf = New-Object byte[] $len
        $read = $fs.Read($buf, 0, $len)

        for ($i = 0; $i -lt $read; $i++) {
            if ($buf[$i] -eq 0) {
                return $true
            }
        }

        return $false
    }
    finally {
        $fs.Dispose()
    }
}

function Read-TextFile {
    param([string]$Path)

    $utf8Strict = New-Object System.Text.UTF8Encoding -ArgumentList $false, $true

    try {
        return [System.IO.File]::ReadAllText($Path, $utf8Strict)
    }
    catch {
        return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
    }
}

function Get-IncludedChildren {
    param([string]$Dir)

    return @(
        Get-ChildItem -LiteralPath $Dir -Force |
        Where-Object {
            if ($_.PSIsContainer) {
                -not (Test-ExcludedDir $_)
            }
            else {
                -not (Test-ExcludedFile $_)
            }
        } |
        Sort-Object @{ Expression = { if ($_.PSIsContainer) { 0 } else { 1 } } }, Name
    )
}

$TreeSb = New-Object System.Text.StringBuilder
$Files = New-Object "System.Collections.Generic.List[System.IO.FileInfo]"
$Skipped = New-Object "System.Collections.Generic.List[string]"

function Append-Tree {
    param(
        [string]$Dir,
        [string]$Indent
    )

    foreach ($item in Get-IncludedChildren $Dir) {
        if ($item.PSIsContainer) {
            [void]$TreeSb.AppendLine("$Indent[D] $($item.Name)/")
            Append-Tree -Dir $item.FullName -Indent ($Indent + "  ")
        }
        else {
            [void]$TreeSb.AppendLine("$Indent[F] $($item.Name)")
        }
    }
}

function Collect-Files {
    param([string]$Dir)

    foreach ($item in Get-IncludedChildren $Dir) {
        if ($item.PSIsContainer) {
            Collect-Files $item.FullName
        }
        else {
            $Files.Add($item) | Out-Null
        }
    }
}

$rootName = Split-Path -Leaf $RootFull
[void]$TreeSb.AppendLine("[D] $rootName/")
Append-Tree -Dir $RootFull -Indent "  "
Collect-Files -Dir $RootFull

$FilesSb = New-Object System.Text.StringBuilder
$IncludedCount = 0

foreach ($file in $Files) {
    $rel = Get-RelativePathForOutput $file.FullName

    try {
        if (Test-BinaryByExtension $file) {
            $Skipped.Add("$rel :: binary extension") | Out-Null
            continue
        }

        if ($file.Length -gt $MaxFileBytes) {
            $Skipped.Add("$rel :: too large ($($file.Length) bytes)") | Out-Null
            continue
        }

        if (Test-ContainsNullByte $file.FullName) {
            $Skipped.Add("$rel :: contains NUL byte") | Out-Null
            continue
        }

        $text = Read-TextFile $file.FullName

        [void]$FilesSb.AppendLine("")
        [void]$FilesSb.AppendLine("===== FILE: $rel =====")
        [void]$FilesSb.AppendLine($text)
        [void]$FilesSb.AppendLine("===== END FILE: $rel =====")

        $IncludedCount++
    }
    catch {
        $Skipped.Add("$rel :: unreadable ($($_.Exception.Message))") | Out-Null
    }
}

$Final = New-Object System.Text.StringBuilder

[void]$Final.AppendLine("REPOSITORY EXPORT")
[void]$Final.AppendLine("ROOT: $RootFull")
[void]$Final.AppendLine("GENERATED_AT: $(Get-Date -Format o)")
[void]$Final.AppendLine("MAX_FILE_BYTES: $MaxFileBytes")
[void]$Final.AppendLine("EXCLUDED_DIRS: $($ExcludeDirs -join ', ')")
[void]$Final.AppendLine("EXCLUDED_FILE_PATTERNS: $($ExcludeFilePatterns -join ', ')")
[void]$Final.AppendLine("")
[void]$Final.AppendLine("===== TREE =====")
[void]$Final.Append($TreeSb.ToString())
[void]$Final.AppendLine("")
[void]$Final.AppendLine("===== SKIPPED FILES =====")

if ($Skipped.Count -eq 0) {
    [void]$Final.AppendLine("(none)")
}
else {
    foreach ($item in $Skipped) {
        [void]$Final.AppendLine("- $item")
    }
}

[void]$Final.AppendLine("")
[void]$Final.AppendLine("===== FILE CONTENTS =====")
[void]$Final.Append($FilesSb.ToString())

$outDir = Split-Path -Parent $OutFull
[System.IO.Directory]::CreateDirectory($outDir) | Out-Null

$utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
[System.IO.File]::WriteAllText($OutFull, $Final.ToString(), $utf8NoBom)

Write-Host "Wrote: $OutFull"
Write-Host "Files included: $IncludedCount"
Write-Host "Files skipped: $($Skipped.Count)"

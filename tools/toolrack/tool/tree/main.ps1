# Tree -- write the folder tree to a markdown file.
# -Depth: 0 = unlimited, N>0 = that depth. Omitted (-1) = interactive (Custom...).
param(
    [string]$Target = "$PWD",
    [int]$Depth = -1,
    [switch]$DirsOnly,
    [switch]$NoOpen
)
. (Join-Path $PSScriptRoot "..\..\common\ui.ps1")

$script:TreeExcludedDirectories = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($name in @("node_modules", ".git", '$RECYCLE.BIN', "AppData")) {
    [void]$script:TreeExcludedDirectories.Add($name)
}

function Get-TreeItems {
    param(
        [Parameter(Mandatory = $true)][System.IO.DirectoryInfo]$Root,
        [int]$MaxDepth
    )

    $results = New-Object 'System.Collections.Generic.List[System.IO.FileSystemInfo]'
    $walk = $null
    $walk = {
        param([System.IO.DirectoryInfo]$Directory, [int]$Level)

        $children = @(Get-ChildItem -LiteralPath $Directory.FullName -ErrorAction SilentlyContinue)
        foreach ($child in $children) {
            if ($child.PSIsContainer -and $script:TreeExcludedDirectories.Contains($child.Name)) {
                continue
            }

            [void]$results.Add($child)
            if (-not $child.PSIsContainer) { continue }
            if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
            if ($MaxDepth -gt 0 -and $Level -ge $MaxDepth) { continue }
            & $walk $child ($Level + 1)
        }
    }

    & $walk $Root 1
    return $results.ToArray()
}

if (-not (Test-Path -LiteralPath $Target -PathType Container)) { Write-Err "Folder not found: $Target"; exit 1 }
$targetItem = Get-Item -LiteralPath $Target
$targetFull = $targetItem.FullName

Start-UI -Title "Tree"

if ($Depth -lt 0) {
    $Depth = Read-Int -Prompt "Depth (0 = unlimited)" -Default 3
    if ($Depth -lt 0) { $Depth = 3 }
    $DirsOnly = -not (Confirm-Yes -Prompt "Include files?" -Default $true)
}
if ($Depth -eq 0 -and $targetFull -match "OneDrive") {
    Write-Warn "Unlimited scan under OneDrive may download cloud-only files."
    if (-not (Confirm-Yes -Prompt "Continue?" -Default $false)) { Stop-UI; exit 1 }
}

$label = Split-Path $targetFull -Leaf
$out = Get-OutputPath -Tool "tree" -Label $label -Ext "md"
$base = ($targetFull.TrimEnd("\") -split "\\").Count

$lines = New-Object "System.Collections.Generic.List[string]"
$lines.Add("# Tree: $Target")
$depthLabel = "unlimited"
if ($Depth -gt 0) { $depthLabel = "$Depth" }
$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')  Depth: $depthLabel  Files: $(-not $DirsOnly)")
$lines.Add("")
$lines.Add('```')

Write-Step "Scanning"
$items = @(Get-TreeItems -Root $targetItem -MaxDepth $Depth)
$count = 0
$items | Where-Object { -not $DirsOnly -or $_.PSIsContainer } |
    Sort-Object FullName | ForEach-Object {
        $indent = "  " * (($_.FullName -split "\\").Count - $base - 1)
        if ($_.PSIsContainer) { $lines.Add("$indent- $($_.Name)/") } else { $lines.Add("$indent- $($_.Name)") }
        $count++
    }
$lines.Add('```')
Write-Ok ("{0} entries" -f $count)

[System.IO.File]::WriteAllText($out, ($lines -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
Stop-UI -OutPath $out
if (-not $NoOpen) { Open-Folder (Split-Path $out -Parent) }

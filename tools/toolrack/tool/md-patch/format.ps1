# MD Patch v1 strict parser. PowerShell 5.1, ASCII source only.

$script:PatchHeader = "# MD Patch v1"
$script:PatchBinaryFileLimit = 5MB
$script:PatchBinaryTotalLimit = 10MB

function Test-PatchSingleLineValue {
    param([string]$Value, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw "$Name must not be empty." }
    foreach ($ch in $Value.ToCharArray()) {
        if ([int]$ch -lt 32 -or [int]$ch -eq 127) { throw "$Name contains a control character." }
    }
}

function Test-PatchRelativePath {
    param([string]$Path, [bool]$Directory = $false)
    Test-PatchSingleLineValue $Path "Path"
    if ($Path.Contains("\")) { throw "Path must use forward slashes: $Path" }
    if ($Path.StartsWith("/") -or [IO.Path]::IsPathRooted($Path)) { throw "Path must be relative: $Path" }
    if ($Path.Contains(":")) { throw "Path contains an invalid colon: $Path" }
    if ($Directory) {
        if (-not $Path.EndsWith("/")) { throw "Directory path must end with '/': $Path" }
        $check = $Path.Substring(0, $Path.Length - 1)
    } else {
        if ($Path.EndsWith("/")) { throw "File path must not end with '/': $Path" }
        $check = $Path
    }
    if ([string]::IsNullOrEmpty($check)) { throw "Path must not be empty." }
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    foreach ($segment in $check.Split([char]'/', [StringSplitOptions]::None)) {
        if ([string]::IsNullOrEmpty($segment) -or $segment -eq "." -or $segment -eq "..") { throw "Path contains an invalid segment: $Path" }
        if ($segment.EndsWith(".") -or $segment.EndsWith(" ")) { throw "Path segment ends with a dot or space: $Path" }
        if ($segment.IndexOfAny($invalid) -ge 0) { throw "Path contains an invalid character: $Path" }
        if ($segment -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$') { throw "Path uses a Windows reserved name: $Path" }
    }
    return $true
}

function Read-PatchRequiredLine {
    param([string[]]$Lines, [ref]$Index, [string]$Expected)
    if ($Index.Value -ge $Lines.Count) { throw "Line $($Index.Value + 1): expected '$Expected', reached end of input." }
    $actual = $Lines[$Index.Value]
    if ($actual -cne $Expected) { throw "Line $($Index.Value + 1): expected '$Expected', got '$actual'." }
    $Index.Value++
}

function Read-PatchValueLine {
    param([string[]]$Lines, [ref]$Index, [string]$Prefix)
    if ($Index.Value -ge $Lines.Count) { throw "Line $($Index.Value + 1): expected '$Prefix', reached end of input." }
    $actual = $Lines[$Index.Value]
    if (-not $actual.StartsWith($Prefix, [StringComparison]::Ordinal)) { throw "Line $($Index.Value + 1): expected '$Prefix'." }
    $value = $actual.Substring($Prefix.Length)
    $Index.Value++
    return $value
}

function ConvertFrom-PatchCount {
    param([string]$Value, [string]$Name, [bool]$AllowZero = $true)
    [int]$number = 0
    if (-not [int]::TryParse($Value, [Globalization.NumberStyles]::None, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        throw "Invalid ${Name}: $Value"
    }
    if ($number -lt 0 -or (-not $AllowZero -and $number -eq 0)) { throw "Invalid ${Name}: $Value" }
    return $number
}

function ConvertFrom-PatchByteCount {
    param([string]$Value, [bool]$AllowDash)
    if ($AllowDash -and $Value -ceq "-") { return [long]-1 }
    [long]$number = 0
    if (-not [long]::TryParse($Value, [Globalization.NumberStyles]::None, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -or $number -lt 0) {
        throw "Invalid byte count: $Value"
    }
    return $number
}

function Test-PatchShaValue {
    param([string]$Value, [bool]$AllowDash = $false)
    if ($AllowDash -and $Value -ceq "-") { return $true }
    if ($Value -cnotmatch '^[0-9a-f]{64}$') { throw "Invalid SHA-256 value: $Value" }
    return $true
}

function Read-PatchContentLines {
    param([string[]]$Lines, [ref]$Index, [int]$Count, [string]$Header)
    Read-PatchRequiredLine $Lines $Index $Header
    if (($Index.Value + $Count) -gt $Lines.Count) { throw "Line $($Index.Value + 1): content block is shorter than its declared count." }
    $content = New-Object System.Collections.Generic.List[string]
    for ($j = 0; $j -lt $Count; $j++) {
        $content.Add($Lines[$Index.Value])
        $Index.Value++
    }
    return [pscustomobject]@{ Lines = $content.ToArray() }
}

function Read-PatchTextFileFields {
    param([string[]]$Lines, [ref]$Index, [string]$EndMarker)
    $encoding = Read-PatchValueLine $Lines $Index "Encoding: "
    if (@("utf-8", "utf-8-bom", "utf-16le", "cp932") -cnotcontains $encoding) { throw "Invalid Encoding: $encoding" }
    $eol = Read-PatchValueLine $Lines $Index "EOL: "
    if (@("lf", "crlf", "none") -cnotcontains $eol) { throw "Invalid EOL: $eol" }
    $finalValue = Read-PatchValueLine $Lines $Index "Final Newline: "
    if (@("yes", "no") -cnotcontains $finalValue) { throw "Invalid Final Newline: $finalValue" }
    $final = ($finalValue -ceq "yes")
    $countValue = Read-PatchValueLine $Lines $Index "New Line Count: "
    $count = ConvertFrom-PatchCount $countValue "New Line Count"
    $block = Read-PatchContentLines $Lines $Index $count ("----- NEW CONTENT: $count LINES -----")
    Read-PatchRequiredLine $Lines $Index $EndMarker
    if ($count -eq 0 -and $final) { throw "An empty file cannot have Final Newline: yes." }
    if ($eol -eq "none" -and ($count -gt 1 -or $final)) { throw "EOL none requires at most one non-terminated line." }
    return [pscustomobject]@{
        Encoding = $encoding; Eol = $eol; FinalNewline = $final
        Lines = $block.Lines
    }
}

function Read-PatchBinaryFields {
    param([string[]]$Lines, [ref]$Index, [string]$EndMarker, [ref]$EmbeddedTotal)
    $byteValue = Read-PatchValueLine $Lines $Index "Bytes: "
    $byteCount = ConvertFrom-PatchByteCount $byteValue $true
    $sha = Read-PatchValueLine $Lines $Index "SHA-256: "
    [void](Test-PatchShaValue $sha $true)
    $lineCountValue = Read-PatchValueLine $Lines $Index "Base64 Line Count: "
    $lineCount = ConvertFrom-PatchCount $lineCountValue "Base64 Line Count"
    $block = Read-PatchContentLines $Lines $Index $lineCount ("----- BASE64 CONTENT: $lineCount LINES -----")
    Read-PatchRequiredLine $Lines $Index $EndMarker
    foreach ($line in $block.Lines) {
        if ($line -cnotmatch '^[A-Za-z0-9+/]*={0,2}$' -or $line.Length -gt 76) { throw "Invalid Base64 content line." }
    }
    try { $bytes = [Convert]::FromBase64String(($block.Lines -join "")) } catch { throw "Invalid Base64 content." }
    if ($bytes.Length -gt $script:PatchBinaryFileLimit) { throw "Binary patch content exceeds the 5 MiB file limit." }
    $EmbeddedTotal.Value += $bytes.Length
    if ($EmbeddedTotal.Value -gt $script:PatchBinaryTotalLimit) { throw "Binary patch content exceeds the 10 MiB total limit." }
    if ($byteCount -ge 0 -and $bytes.Length -ne $byteCount) { throw "Binary byte count mismatch." }
    $actualSha = Get-PatchSha256 $bytes
    if ($sha -cne "-" -and $sha -cne $actualSha) { throw "Binary SHA-256 mismatch." }
    return [pscustomobject]@{ Bytes = $bytes; Sha256 = $actualSha }
}

function Get-PatchSha256 {
    param([byte[]]$Bytes)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($Bytes)
        return (($hash | ForEach-Object { $_.ToString("x2") }) -join "")
    } finally { $sha.Dispose() }
}

function Read-PatchDocument {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)
    if ($Text.Contains([char]0)) { throw "Patch contains a NUL character." }
    $normalized = $Text.Replace("`r`n", "`n").Replace("`r", "`n")
    $lines = [regex]::Split($normalized, "`n")
    [int]$i = 0
    Read-PatchRequiredLine $lines ([ref]$i) $script:PatchHeader
    $countValue = Read-PatchValueLine $lines ([ref]$i) "Operation Count: "
    $operationCount = ConvertFrom-PatchCount $countValue "Operation Count" $false
    Read-PatchRequiredLine $lines ([ref]$i) ""
    $operations = New-Object System.Collections.Generic.List[object]
    $paths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    [long]$embeddedTotal = 0

    for ($operationNumber = 1; $operationNumber -le $operationCount; $operationNumber++) {
        if ($i -ge $lines.Count) { throw "Operation $operationNumber is missing." }
        $startLine = $lines[$i]
        $i++
        $type = ""
        $path = ""
        if ($startLine -cmatch '^===== MODIFY FILE: (.+) =====$') { $type = "modify"; $path = $Matches[1] }
        elseif ($startLine -cmatch '^===== CREATE TEXT FILE: (.+) =====$') { $type = "create-text"; $path = $Matches[1] }
        elseif ($startLine -cmatch '^===== REPLACE TEXT FILE: (.+) =====$') { $type = "replace-text"; $path = $Matches[1] }
        elseif ($startLine -cmatch '^===== DELETE FILE: (.+) =====$') { $type = "delete-file"; $path = $Matches[1] }
        elseif ($startLine -cmatch '^===== CREATE DIRECTORY: (.+) =====$') { $type = "create-directory"; $path = $Matches[1] }
        elseif ($startLine -cmatch '^===== DELETE DIRECTORY: (.+) =====$') { $type = "delete-directory"; $path = $Matches[1] }
        elseif ($startLine -cmatch '^===== CREATE BINARY FILE: (.+) =====$') { $type = "create-binary"; $path = $Matches[1] }
        elseif ($startLine -cmatch '^===== REPLACE BINARY FILE: (.+) =====$') { $type = "replace-binary"; $path = $Matches[1] }
        else { throw "Line ${i}: unknown operation header: $startLine" }

        $isDirectory = ($type -eq "create-directory" -or $type -eq "delete-directory")
        [void](Test-PatchRelativePath $path $isDirectory)
        $canonical = $path.TrimEnd('/')
        if (-not $paths.Add($canonical)) { throw "Duplicate operation path: $path" }

        if ($type -eq "modify") {
            $expectedSha = Read-PatchValueLine $lines ([ref]$i) "Expected SHA-256: "
            [void](Test-PatchShaValue $expectedSha)
            $changeCountValue = Read-PatchValueLine $lines ([ref]$i) "Change Count: "
            $changeCount = ConvertFrom-PatchCount $changeCountValue "Change Count" $false
            $changes = New-Object System.Collections.Generic.List[object]
            for ($changeNumber = 1; $changeNumber -le $changeCount; $changeNumber++) {
                Read-PatchRequiredLine $lines ([ref]$i) "----- CHANGE $changeNumber -----"
                $oldStart = ConvertFrom-PatchCount (Read-PatchValueLine $lines ([ref]$i) "Old Start Line: ") "Old Start Line" $false
                $oldCount = ConvertFrom-PatchCount (Read-PatchValueLine $lines ([ref]$i) "Old Line Count: ") "Old Line Count"
                $newStart = ConvertFrom-PatchCount (Read-PatchValueLine $lines ([ref]$i) "New Start Line: ") "New Start Line" $false
                $newCount = ConvertFrom-PatchCount (Read-PatchValueLine $lines ([ref]$i) "New Line Count: ") "New Line Count"
                $oldBlock = Read-PatchContentLines $lines ([ref]$i) $oldCount ("----- OLD CONTENT: $oldCount LINES -----")
                $newBlock = Read-PatchContentLines $lines ([ref]$i) $newCount ("----- NEW CONTENT: $newCount LINES -----")
                Read-PatchRequiredLine $lines ([ref]$i) "----- END CHANGE $changeNumber -----"
                $changes.Add([pscustomobject]@{
                    OldStartLine = $oldStart; OldLineCount = $oldCount; OldLines = $oldBlock.Lines
                    NewStartLine = $newStart; NewLineCount = $newCount; NewLines = $newBlock.Lines
                })
            }
            Read-PatchRequiredLine $lines ([ref]$i) "===== END MODIFY FILE: $path ====="
            $operations.Add([pscustomobject]@{ Type = $type; Path = $path; ExpectedSha256 = $expectedSha; Changes = $changes.ToArray() })
        } elseif ($type -eq "create-text" -or $type -eq "replace-text") {
            $expectedSha = ""
            if ($type -eq "replace-text") {
                $expectedSha = Read-PatchValueLine $lines ([ref]$i) "Expected SHA-256: "
                [void](Test-PatchShaValue $expectedSha)
            }
            $endMarker = if ($type -eq "create-text") { "===== END CREATE TEXT FILE: $path =====" } else { "===== END REPLACE TEXT FILE: $path =====" }
            $fields = Read-PatchTextFileFields $lines ([ref]$i) $endMarker
            $operations.Add([pscustomobject]@{
                Type = $type; Path = $path; ExpectedSha256 = $expectedSha
                Encoding = $fields.Encoding; Eol = $fields.Eol; FinalNewline = $fields.FinalNewline; Lines = $fields.Lines
            })
        } elseif ($type -eq "delete-file") {
            $expectedSha = Read-PatchValueLine $lines ([ref]$i) "Expected SHA-256: "
            [void](Test-PatchShaValue $expectedSha)
            Read-PatchRequiredLine $lines ([ref]$i) "===== END DELETE FILE: $path ====="
            $operations.Add([pscustomobject]@{ Type = $type; Path = $path; ExpectedSha256 = $expectedSha })
        } elseif ($type -eq "create-directory" -or $type -eq "delete-directory") {
            $endMarker = if ($type -eq "create-directory") { "===== END CREATE DIRECTORY: $path =====" } else { "===== END DELETE DIRECTORY: $path =====" }
            Read-PatchRequiredLine $lines ([ref]$i) $endMarker
            $operations.Add([pscustomobject]@{ Type = $type; Path = $path })
        } else {
            $expectedSha = ""
            if ($type -eq "replace-binary") {
                $expectedSha = Read-PatchValueLine $lines ([ref]$i) "Expected SHA-256: "
                [void](Test-PatchShaValue $expectedSha)
            }
            $endMarker = if ($type -eq "create-binary") { "===== END CREATE BINARY FILE: $path =====" } else { "===== END REPLACE BINARY FILE: $path =====" }
            $fields = Read-PatchBinaryFields $lines ([ref]$i) $endMarker ([ref]$embeddedTotal)
            $operations.Add([pscustomobject]@{
                Type = $type; Path = $path; ExpectedSha256 = $expectedSha
                Bytes = $fields.Bytes; Sha256 = $fields.Sha256
            })
        }
        Read-PatchRequiredLine $lines ([ref]$i) ""
    }
    Read-PatchRequiredLine $lines ([ref]$i) "===== END PATCH ====="
    if ($i -lt $lines.Count) {
        if ($i -ne ($lines.Count - 1) -or $lines[$i] -cne "") { throw "Line $($i + 1): unexpected content after the patch." }
        $i++
    }
    if ($i -ne $lines.Count) { throw "Unexpected trailing patch content." }
    return [pscustomobject]@{ Version = 1; Operations = $operations.ToArray() }
}

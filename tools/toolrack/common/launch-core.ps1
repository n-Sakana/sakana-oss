# common/launch-core.ps1 -- shared launch validation and process execution.
# This file is a library. It must not own UI, pauses, or process exit.

function Get-LaunchProp {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    $property = $Obj.PSObject.Properties[$Name]
    if ($property) { return ,$property.Value }
    return $null
}

function Test-LaunchProp {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $false }
    return ($null -ne $Obj.PSObject.Properties[$Name])
}

function New-LaunchPlanResult {
    return @{
        Ok = $false
        Errors = @()
        Name = "tool"
        Type = ""
        Entry = ""
        Window = "console"
        KeepOpen = $true
        Requires = @()
        Args = @()
        ToolDir = ""
        ActionId = ""
    }
}

function Resolve-LaunchPlan {
    param(
        [string]$ToolDir,
        [int]$VariantIndex = -1,
        [bool]$VariantSpecified = $false,
        [AllowNull()][string]$ActionId = $null,
        [bool]$ActionSpecified = $false
    )
    $result = New-LaunchPlanResult

    if ($VariantSpecified -and $ActionSpecified) {
        $result.Errors = @("Action and Variant cannot be specified together")
        return $result
    }
    if ($VariantSpecified -and $VariantIndex -lt -1) {
        $result.Errors = @("variant index must be -1 or greater")
        return $result
    }
    if ($ActionSpecified -and ($ActionId -isnot [string] -or $ActionId -cnotmatch '^[a-z0-9][a-z0-9-]*$')) {
        $result.Errors = @("action: invalid stable action ID")
        return $result
    }

    $toolFull = $null
    try { $toolFull = [IO.Path]::GetFullPath($ToolDir) }
    catch {
        $result.Errors = @("invalid tool folder: " + $ToolDir)
        return $result
    }
    $result.ToolDir = $toolFull
    $manifestPath = Join-Path $toolFull "tool.json"
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        $result.Errors = @("tool.json not found: " + $manifestPath)
        return $result
    }

    $manifest = $null
    try { $manifest = ConvertFrom-Json ([IO.File]::ReadAllText($manifestPath)) }
    catch {
        $result.Errors = @("tool.json parse error: " + $_.Exception.Message)
        return $result
    }
    if ($null -eq $manifest -or $manifest -isnot [pscustomobject]) {
        $result.Errors = @("tool.json root: must be an object")
        return $result
    }

    $schema = Get-LaunchProp $manifest "schema"
    if ($schema -isnot [int] -or $schema -ne 1) {
        $result.Errors = @("schema: unsupported or invalid")
        return $result
    }
    $name = Get-LaunchProp $manifest "name"
    if ($name -is [string] -and $name.Trim() -ne "") { $result.Name = $name }

    $run = Get-LaunchProp $manifest "run"
    if ($run -isnot [pscustomobject]) {
        $result.Errors = @("run: missing or invalid")
        return $result
    }
    $type = Get-LaunchProp $run "type"
    if ($type -isnot [string] -or $type -notin @("powershell", "python")) {
        $result.Errors = @("run.type: unknown or invalid")
        return $result
    }
    $result.Type = $type

    $entry = Get-LaunchProp $run "entry"
    if ($entry -isnot [string] -or $entry.Trim() -eq "" -or [IO.Path]::IsPathRooted($entry)) {
        $result.Errors = @("run.entry: must be a relative file path inside the tool folder")
        return $result
    }
    $entryFull = $null
    try { $entryFull = [IO.Path]::GetFullPath((Join-Path $toolFull $entry)) }
    catch {
        $result.Errors = @("run.entry: invalid path")
        return $result
    }
    $prefix = $toolFull.TrimEnd("\") + "\"
    if (-not $entryFull.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        $result.Errors = @("run.entry: must stay inside the tool folder")
        return $result
    }
    if (-not (Test-Path -LiteralPath $entryFull -PathType Leaf)) {
        $result.Errors = @("run.entry not found: " + $entryFull)
        return $result
    }
    $result.Entry = $entryFull

    if (Test-LaunchProp $run "window") {
        $window = Get-LaunchProp $run "window"
        if ($window -isnot [string] -or $window -notin @("console", "hidden", "gui")) {
            $result.Errors = @("run.window: invalid")
            return $result
        }
        $result.Window = $window
    }
    if ($result.Window -ne "console") { $result.KeepOpen = $false }

    if (Test-LaunchProp $run "keep_open") {
        $keepOpen = Get-LaunchProp $run "keep_open"
        if ($keepOpen -isnot [bool] -or $result.Window -ne "console") {
            $result.Errors = @("run.keep_open: invalid for this window mode")
            return $result
        }
        $result.KeepOpen = $keepOpen
    }

    if (Test-LaunchProp $run "requires") {
        $requires = Get-LaunchProp $run "requires"
        if ($requires -isnot [Array] -or $result.Type -ne "python") {
            $result.Errors = @("run.requires: must be an array for python type")
            return $result
        }
        $requireItems = @($requires)
        $badRequires = @($requireItems | Where-Object { $_ -isnot [string] -or $_ -notmatch '^[A-Za-z0-9_]+$' })
        if ($badRequires.Count -gt 0) {
            $result.Errors = @("run.requires: invalid import name")
            return $result
        }
        $result.Requires = $requireItems
    }

    $variantList = @()
    if (Test-LaunchProp $manifest "variants") {
        $variants = Get-LaunchProp $manifest "variants"
        if ($variants -isnot [Array]) {
            $result.Errors = @("variants: must be an array")
            return $result
        }
        $variantList = @($variants)
    }

    if ($ActionSpecified) {
        $result.ActionId = $ActionId
        if ($variantList.Count -eq 0) {
            if ($ActionId -cne "default") {
                $result.Errors = @("unknown action '" + $ActionId + "' for flat tool")
                return $result
            }
        } else {
            $ids = New-Object "System.Collections.Generic.HashSet[string]" ([StringComparer]::Ordinal)
            $idCount = 0
            $selected = $null
            for ($index = 0; $index -lt $variantList.Count; $index++) {
                $variant = $variantList[$index]
                if ($variant -isnot [pscustomobject]) {
                    $result.Errors = @("variant " + $index + ": invalid")
                    return $result
                }
                if (Test-LaunchProp $variant "id") {
                    $idCount++
                    $variantId = Get-LaunchProp $variant "id"
                    if ($variantId -isnot [string] -or $variantId -cnotmatch '^[a-z0-9][a-z0-9-]*$') {
                        $result.Errors = @("variant " + $index + " id: invalid")
                        return $result
                    }
                    if (-not $ids.Add($variantId)) {
                        $result.Errors = @("variant id: duplicate '" + $variantId + "'")
                        return $result
                    }
                    if ($variantId -ceq $ActionId) { $selected = $variant }
                }
            }
            if ($idCount -eq 0) {
                $result.Errors = @("tool does not expose stable action IDs")
                return $result
            }
            if ($idCount -ne $variantList.Count) {
                $result.Errors = @("variant action IDs must be present on every variant")
                return $result
            }
            if ($null -eq $selected) {
                $result.Errors = @("unknown action '" + $ActionId + "'")
                return $result
            }
            $selectedArgs = Get-LaunchProp $selected "args"
            if ($selectedArgs -isnot [Array]) {
                $result.Errors = @("action args: must be an array")
                return $result
            }
            $actionArgs = @($selectedArgs)
            $badActionArgs = @($actionArgs | Where-Object { $_ -isnot [string] -or $_.Length -eq 0 -or $_.Contains('"') })
            if ($badActionArgs.Count -gt 0) {
                $result.Errors = @("action args: invalid")
                return $result
            }
            $result.Args = $actionArgs
        }
    } else {
        $effectiveVariant = -1
        if ($VariantSpecified) { $effectiveVariant = $VariantIndex }
        if ($effectiveVariant -ge 0) {
            if ($effectiveVariant -ge $variantList.Count) {
                $result.Errors = @("variant index " + $effectiveVariant + " out of range (have " + $variantList.Count + ")")
                return $result
            }
            $selectedVariant = $variantList[$effectiveVariant]
            if ($selectedVariant -isnot [pscustomobject]) {
                $result.Errors = @("variant " + $effectiveVariant + ": invalid")
                return $result
            }
            $selectedArgs = Get-LaunchProp $selectedVariant "args"
            if ($selectedArgs -isnot [Array]) {
                $result.Errors = @("variant " + $effectiveVariant + " args: must be an array")
                return $result
            }
            $variantArgs = @($selectedArgs)
            $badVariantArgs = @($variantArgs | Where-Object { $_ -isnot [string] -or $_.Length -eq 0 -or $_.Contains('"') })
            if ($badVariantArgs.Count -gt 0) {
                $result.Errors = @("variant " + $effectiveVariant + " args: invalid")
                return $result
            }
            $result.Args = $variantArgs
        }
    }

    $result.Ok = $true
    return $result
}

function Resolve-Launch {
    param([string]$ToolDir, [int]$VariantIndex)
    return Resolve-LaunchPlan -ToolDir $ToolDir -VariantIndex $VariantIndex -VariantSpecified $true -ActionSpecified $false
}

function ConvertTo-WindowsArgument {
    param([string]$Item)
    if ($null -eq $Item -or $Item.Length -eq 0) { return '""' }
    if ($Item -notmatch '[\s"]') { return $Item }

    $builder = New-Object Text.StringBuilder
    [void]$builder.Append('"')
    $backslashes = 0
    foreach ($character in $Item.ToCharArray()) {
        if ($character -eq '\') {
            $backslashes++
        } elseif ($character -eq '"') {
            [void]$builder.Append('\' * (($backslashes * 2) + 1))
            [void]$builder.Append('"')
            $backslashes = 0
        } else {
            if ($backslashes -gt 0) {
                [void]$builder.Append('\' * $backslashes)
                $backslashes = 0
            }
            [void]$builder.Append($character)
        }
    }
    if ($backslashes -gt 0) { [void]$builder.Append('\' * ($backslashes * 2)) }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function ConvertTo-ArgString {
    param([string[]]$Items)
    return (@($Items | ForEach-Object { ConvertTo-WindowsArgument $_ }) -join " ")
}

function Get-MissingPythonModules {
    param([string[]]$Modules)
    if (@($Modules).Count -eq 0) { return @() }
    $code = "import importlib.util,sys;print(','.join([x for x in sys.argv[1:] if importlib.util.find_spec(x) is None]))"
    $output = $null
    $pythonExitCode = 1
    $oldErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $output = & py -3 -c $code @Modules 2>$null
        $pythonExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $oldErrorActionPreference
    }
    if ($pythonExitCode -ne 0) { throw "py -3 failed" }
    $text = ([string]$output).Trim()
    if ($text -eq "") { return @() }
    return @($text -split ",")
}

function Get-PythonwPath {
    $output = $null
    $pythonExitCode = 1
    $oldErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $output = & py -3 -c "import sys;print(sys.executable)" 2>$null
        $pythonExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $oldErrorActionPreference
    }
    $executable = ([string]$output).Trim()
    if ($pythonExitCode -ne 0 -or $executable -eq "") { throw "py -3 failed" }
    $pythonw = Join-Path (Split-Path $executable -Parent) "pythonw.exe"
    if (-not (Test-Path -LiteralPath $pythonw -PathType Leaf)) { throw "pythonw.exe not found: " + $pythonw }
    return $pythonw
}

function New-LaunchExecutionResult {
    return @{ Ok = $false; Errors = @(); Started = $false; ExitCode = 1; ProcessId = 0; Detached = $false }
}

function Invoke-HostAction {
    param(
        [Parameter(Mandatory = $true)][string]$ToolDir,
        [Parameter(Mandatory = $true)][string]$ActionId
    )
    $plan = Resolve-LaunchPlan -ToolDir $ToolDir -ActionId $ActionId -ActionSpecified $true
    if (-not $plan.Ok) {
        $result = New-LaunchExecutionResult
        $result.Errors = @($plan.Errors)
        return $result
    }
    return Invoke-LaunchPlan -Plan $plan -Target "" -Route Host
}

function Invoke-LaunchPlan {
    param(
        $Plan,
        [AllowNull()][string]$Target = "",
        [ValidateSet("Explorer", "Host")][string]$Route = "Explorer"
    )
    $execution = New-LaunchExecutionResult
    if ($null -eq $Plan -or -not $Plan.Ok) {
        $execution.Errors = @("launch plan is invalid")
        return $execution
    }

    $executable = ""
    $argumentItems = @()
    if ($Plan.Type -eq "powershell") {
        $executable = "powershell.exe"
        $argumentItems = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $Plan.Entry, "-Target", $Target) + @($Plan.Args)
        if ($Plan.Window -ne "console" -or $Route -eq "Host") {
            $argumentItems = @("-WindowStyle", "Hidden") + $argumentItems
        }
    } else {
        if ($null -eq (Get-Command py -ErrorAction SilentlyContinue)) {
            $execution.Errors = @("Python launcher (py) not found on this PC.`nInstall Python 3 to run this tool.")
            return $execution
        }
        $missing = @()
        try { $missing = @(Get-MissingPythonModules @($Plan.Requires)) }
        catch {
            $execution.Errors = @("Failed to run 'py -3'. Check the Python installation.`n" + $_.Exception.Message)
            return $execution
        }
        if ($missing.Count -gt 0) {
            $fix = ($missing | ForEach-Object { "pip install --user " + $_ }) -join "`n"
            $execution.Errors = @("Missing Python modules: " + ($missing -join ", ") + "`n`nTo fix, run:`n" + $fix)
            return $execution
        }
        if ($Plan.Window -ne "console") {
            try { $executable = Get-PythonwPath }
            catch {
                $execution.Errors = @([string]$_.Exception.Message)
                return $execution
            }
            $argumentItems = @($Plan.Entry, "--target", $Target) + @($Plan.Args)
        } else {
            $executable = "py"
            $argumentItems = @("-3", $Plan.Entry, "--target", $Target) + @($Plan.Args)
        }
    }

    $detached = ($Route -eq "Host" -or $Plan.Window -ne "console")
    try {
        $startInfo = New-Object Diagnostics.ProcessStartInfo
        $startInfo.FileName = $executable
        $startInfo.Arguments = ConvertTo-ArgString $argumentItems
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $detached
        if ($detached) { $startInfo.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden }
        $process = [Diagnostics.Process]::Start($startInfo)
        $execution.Started = $true
        $execution.ProcessId = $process.Id
        $execution.Detached = $detached
        if ($detached) {
            $execution.ExitCode = 0
        } else {
            $process.WaitForExit()
            $execution.ExitCode = $process.ExitCode
        }
        $process.Dispose()
        $execution.Ok = $true
    } catch {
        $execution.Errors = @("Failed to start tool: " + $_.Exception.Message)
    }
    return $execution
}

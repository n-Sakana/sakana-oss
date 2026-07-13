# common/install.ps1 -- scan tool/, validate manifests, build one .reg, import (HKCU).
# Dot-source loads functions only; direct run performs the install.
param([string]$BindingsConfigPath = "")
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:Root       = Split-Path $PSScriptRoot -Parent
$script:ToolRoot   = Join-Path $script:Root "tool"
$script:MenuPath   = Join-Path $script:Root "menu.json"
$script:BindingsPath = Join-Path $script:Root "bindings.json"
$script:LaunchPs1  = Join-Path $PSScriptRoot "launch.ps1"
$script:SilentVbs  = Join-Path $PSScriptRoot "silent.vbs"
$script:MenuName   = "Tool Rack"
$script:KeyName    = "ToolRack"
$script:MenuSchema = 1
$script:MenuLimit  = 16
$script:SupportedSchema = @(1)
$script:DefaultContexts = @(
    @{ on = "file";       base = "HKEY_CURRENT_USER\Software\Classes\*\shell" },
    @{ on = "folder";     base = "HKEY_CURRENT_USER\Software\Classes\Directory\shell" },
    @{ on = "background"; base = "HKEY_CURRENT_USER\Software\Classes\Directory\Background\shell" }
)
$script:Contexts = @($script:DefaultContexts)

. (Join-Path $PSScriptRoot "bindings.ps1")

function Get-Prop {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    $p = $Obj.PSObject.Properties[$Name]
    if ($p) { return ,$p.Value } else { return $null }
}

function Test-Prop {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $false }
    return ($null -ne $Obj.PSObject.Properties[$Name])
}

function Get-ManifestActionInfo {
    param($Manifest)
    $result = @{ SupportsActions = $false; ActionIds = @() }
    if (-not (Test-Prop $Manifest "variants")) {
        $result.SupportsActions = $true
        $result.ActionIds = @("default")
        return $result
    }
    $variants = Get-Prop $Manifest "variants"
    if ($variants -isnot [Array]) { return $result }
    $items = @($variants)
    if ($items.Count -eq 0) {
        $result.SupportsActions = $true
        $result.ActionIds = @("default")
        return $result
    }
    $ids = New-Object "System.Collections.Generic.HashSet[string]"
    $orderedIds = New-Object "System.Collections.Generic.List[string]"
    foreach ($variant in $items) {
        if ($variant -isnot [pscustomobject] -or -not (Test-Prop $variant "id")) { return $result }
        $id = Get-Prop $variant "id"
        if ($id -isnot [string] -or $id -cnotmatch '^[a-z0-9][a-z0-9-]*$' -or -not $ids.Add($id)) {
            return $result
        }
        [void]$orderedIds.Add($id)
    }
    $result.SupportsActions = $true
    $result.ActionIds = @($orderedIds)
    return $result
}

function Read-Manifest {
    param([string]$ToolDir)
    $result = @{ Ok = $false; Data = $null; Errors = @(); SupportsActions = $false; ActionIds = @() }
    $path = Join-Path $ToolDir "tool.json"
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $result.Errors = @("tool.json not found")
        return $result
    }
    try {
        $result.Data = ConvertFrom-Json ([System.IO.File]::ReadAllText($path))
        if ($null -eq $result.Data -or $result.Data -isnot [pscustomobject]) {
            $result.Errors = @("tool.json root: must be an object")
            return $result
        }
        $actionInfo = Get-ManifestActionInfo $result.Data
        $result.SupportsActions = $actionInfo.SupportsActions
        $result.ActionIds = @($actionInfo.ActionIds)
        $result.Ok = $true
    } catch {
        $result.Errors = @("tool.json parse error: " + $_.Exception.Message)
    }
    return $result
}

function Test-Manifest {
    # Returns violation messages; empty = valid. See design doc section 5.
    param($M, [string]$FolderName, [string]$ToolDir)
    $e = New-Object "System.Collections.Generic.List[string]"

    $schema = Get-Prop $M "schema"
    if ($schema -isnot [int] -or ($script:SupportedSchema -notcontains $schema)) {
        $e.Add("schema: must be an integer in [" + ($script:SupportedSchema -join ", ") + "]")
    }

    $id = Get-Prop $M "id"
    if ($id -isnot [string] -or $id -notmatch '^[A-Za-z0-9-]+$') {
        $e.Add("id: required; letters, digits, hyphen only")
    } elseif ($id -cne $FolderName) {
        $e.Add("id: must equal folder name '$FolderName' (got '$id')")
    }

    $name = Get-Prop $M "name"
    if ($name -isnot [string] -or $name.Trim() -eq "") {
        $e.Add("name: required, non-empty string")
    } elseif ($name -match '[\x00-\x1F]') {
        $e.Add("name: control characters are not allowed")
    }

    $on = Get-Prop $M "on"
    if ($on -isnot [System.Array]) {
        $e.Add("on: required array; non-empty subset of file/folder/background")
    } else {
        $onList = @($on)
        $bad = @($onList | Where-Object { $_ -isnot [string] -or $_ -notin @("file", "folder", "background") })
        if ($onList.Count -eq 0 -or $bad.Count -gt 0) {
            $e.Add("on: must be a non-empty subset of file/folder/background")
        }
    }

    $type = $null
    $run = Get-Prop $M "run"
    if ($run -isnot [pscustomobject]) {
        $e.Add("run: required object")
    } else {
        $type = Get-Prop $run "type"
        if ($type -isnot [string] -or $type -notin @("powershell", "python")) {
            $e.Add("run.type: must be 'powershell' or 'python'")
        }

        $entry = Get-Prop $run "entry"
        if ($entry -isnot [string] -or $entry.Trim() -eq "") {
            $e.Add("run.entry: required, non-empty string")
        } elseif ([System.IO.Path]::IsPathRooted($entry)) {
            $e.Add("run.entry: must be a relative path inside the tool folder")
        } else {
            $full = $null
            try { $full = [System.IO.Path]::GetFullPath((Join-Path $ToolDir $entry)) } catch { }
            $prefix = [System.IO.Path]::GetFullPath($ToolDir).TrimEnd("\") + "\"
            if ($null -eq $full -or -not $full.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $e.Add("run.entry: must stay inside the tool folder")
            } elseif (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
                $e.Add("run.entry: file not found: $entry")
            }
        }

        if (Test-Prop $run "window") {
            $window = Get-Prop $run "window"
            if ($window -isnot [string] -or $window -notin @("console", "hidden", "gui")) {
                $e.Add("run.window: must be 'console', 'hidden', or 'gui'")
            }
        } else {
            $window = "console"
        }

        if (Test-Prop $run "keep_open") {
            $keepOpen = Get-Prop $run "keep_open"
            if ($keepOpen -isnot [bool]) {
                $e.Add("run.keep_open: must be true or false")
            } elseif ($window -ne "console") {
                $e.Add("run.keep_open: only valid for console window")
            }
        }

        if (Test-Prop $run "requires") {
            $requires = Get-Prop $run "requires"
            if ($requires -isnot [System.Array]) {
                $e.Add("run.requires: must be an array")
            } else {
                $reqList = @($requires)
                $badReq = @($reqList | Where-Object { $_ -isnot [string] -or $_ -notmatch '^[A-Za-z0-9_]+$' })
                if ($badReq.Count -gt 0) {
                    $e.Add("run.requires: every element must be a top-level import name (letters, digits, underscore)")
                }
            }
            if ($type -ne "python") {
                $e.Add("run.requires: only valid for python type")
            }
        }
    }

    if (Test-Prop $M "variants") {
        $variants = Get-Prop $M "variants"
        if ($variants -isnot [System.Array]) {
            $e.Add("variants: must be an array")
        } else {
            $vList = @($variants)
            $labels = New-Object "System.Collections.Generic.HashSet[string]"
            $variantIds = New-Object "System.Collections.Generic.HashSet[string]"
            $variantIdCount = 0
            for ($i = 0; $i -lt $vList.Count; $i++) {
                if ($vList[$i] -isnot [pscustomobject]) {
                    $e.Add("variants[$i]: must be an object")
                    continue
                }
                $label = Get-Prop $vList[$i] "label"
                if ($label -isnot [string] -or $label.Trim() -eq "") {
                    $e.Add("variants[$i].label: required, non-empty string")
                } elseif ($label -match '[\x00-\x1F]') {
                    $e.Add("variants[$i].label: control characters are not allowed")
                } elseif (-not $labels.Add($label)) {
                    $e.Add("variants[$i].label: duplicate label '$label'")
                }
                if (Test-Prop $vList[$i] "id") {
                    $variantIdCount++
                    $variantId = Get-Prop $vList[$i] "id"
                    if ($variantId -isnot [string] -or $variantId -cnotmatch '^[a-z0-9][a-z0-9-]*$') {
                        $e.Add("variants[$i].id: must match ^[a-z0-9][a-z0-9-]*$")
                    } elseif (-not $variantIds.Add($variantId)) {
                        $e.Add("variants[$i].id: duplicate id '$variantId'")
                    }
                }
                if (-not (Test-Prop $vList[$i] "args")) {
                    $e.Add("variants[$i].args: required (use [] for none)")
                    continue
                }
                $vargs = Get-Prop $vList[$i] "args"
                if ($vargs -isnot [System.Array]) {
                    $e.Add("variants[$i].args: must be an array")
                    continue
                }
                foreach ($a in @($vargs)) {
                    if ($a -isnot [string]) {
                        $e.Add("variants[$i].args: every element must be a string")
                        break
                    }
                    if ($a.Length -eq 0) {
                        $e.Add("variants[$i].args: empty strings are not allowed")
                        break
                    }
                    if ($a.Contains('"')) {
                        $e.Add("variants[$i].args: double quotes are not allowed")
                        break
                    }
                }
            }
            if ($variantIdCount -gt 0 -and $variantIdCount -ne $vList.Count) {
                $e.Add("variants: id must be present on every variant or none")
            }
        }
    }
    return $e
}

function Read-MenuConfig {
    param([string]$Path)
    $result = @{ Ok = $false; Exists = $false; Data = $null; Errors = @() }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Ok = $true
        return $result
    }
    $result.Exists = $true

    try {
        $m = ConvertFrom-Json ([System.IO.File]::ReadAllText($Path))
    } catch {
        $result.Errors = @("menu.json parse error: " + $_.Exception.Message)
        return $result
    }
    if ($null -eq $m -or $m -isnot [pscustomobject]) {
        $result.Errors = @("menu.json root: must be an object")
        return $result
    }

    $e = New-Object "System.Collections.Generic.List[string]"
    $schema = Get-Prop $m "schema"
    if ($schema -isnot [int] -or $schema -ne $script:MenuSchema) {
        $e.Add("schema: must be integer " + $script:MenuSchema)
    }

    $defaultId = Get-Prop $m "default_category"
    if ($defaultId -isnot [string] -or $defaultId -notmatch '^[a-z0-9][a-z0-9-]*$') {
        $e.Add("default_category: required; lowercase letters, digits, hyphen only")
    }

    $normalized = New-Object "System.Collections.Generic.List[object]"
    $categoryIds = @{}
    $categoryLabels = @{}
    $toolOwners = @{}
    $categories = Get-Prop $m "categories"
    if ($categories -isnot [System.Array] -or @($categories).Count -eq 0) {
        $e.Add("categories: required, non-empty array")
    } else {
        for ($i = 0; $i -lt @($categories).Count; $i++) {
            $c = @($categories)[$i]
            if ($c -isnot [pscustomobject]) {
                $e.Add("categories[$i]: must be an object")
                continue
            }

            $id = Get-Prop $c "id"
            $idOk = ($id -is [string] -and $id -match '^[a-z0-9][a-z0-9-]*$')
            if (-not $idOk) {
                $e.Add("categories[$i].id: required; lowercase letters, digits, hyphen only")
            } elseif ($categoryIds.ContainsKey($id)) {
                $e.Add("categories[$i].id: duplicate id '$id'")
                $idOk = $false
            } else {
                $categoryIds[$id] = $true
            }

            $label = Get-Prop $c "label"
            $labelOk = ($label -is [string] -and $label.Trim() -ne "" -and $label -notmatch '[\x00-\x1F]')
            if (-not $labelOk) {
                $e.Add("categories[$i].label: required, non-empty string without control characters")
            } else {
                $labelKey = $label.ToLowerInvariant()
                if ($categoryLabels.ContainsKey($labelKey)) {
                    $e.Add("categories[$i].label: duplicate label '$label'")
                    $labelOk = $false
                } else {
                    $categoryLabels[$labelKey] = $true
                }
            }

            $toolIds = Get-Prop $c "tools"
            $toolsOk = ($toolIds -is [System.Array])
            $normalizedIds = New-Object "System.Collections.Generic.List[string]"
            if (-not $toolsOk) {
                $e.Add("categories[$i].tools: required array")
            } else {
                for ($j = 0; $j -lt @($toolIds).Count; $j++) {
                    $toolId = @($toolIds)[$j]
                    if ($toolId -isnot [string] -or $toolId -notmatch '^[A-Za-z0-9-]+$') {
                        $e.Add("categories[$i].tools[$j]: must be a tool id")
                        $toolsOk = $false
                        continue
                    }
                    if ($toolOwners.ContainsKey($toolId)) {
                        $e.Add("tool '$toolId' is assigned more than once")
                        $toolsOk = $false
                        continue
                    }
                    $toolOwners[$toolId] = $id
                    $normalizedIds.Add($toolId)
                }
            }

            if ($idOk -and $labelOk -and $toolsOk) {
                $normalized.Add(@{ Id = [string]$id; Label = [string]$label; ToolIds = $normalizedIds.ToArray() })
            }
        }
    }

    if ($defaultId -is [string] -and -not $categoryIds.ContainsKey($defaultId)) {
        $e.Add("default_category: category '$defaultId' not found")
    }
    if ($e.Count -gt 0) {
        $result.Errors = @($e)
        return $result
    }

    $result.Data = @{ DefaultCategory = [string]$defaultId; Categories = $normalized.ToArray() }
    $result.Ok = $true
    return $result
}

function Get-ToolMenuCost {
    param($Tool)
    return (1 + @($Tool.VariantLabels).Count)
}

function Resolve-MenuLayout {
    param($Tools, $MenuConfig)
    $result = @{
        Ok = $true
        Pages = @()
        Warnings = @()
        Rejected = @()
        RegisteredCount = 0
    }
    $toolList = @($Tools | ForEach-Object { $_ })

    if ($null -eq $MenuConfig) {
        $allIds = @($toolList | ForEach-Object { [string]$_.Id })
        $defaultId = "general"
        $categories = @(@{ Id = $defaultId; Label = $script:MenuName; ToolIds = $allIds })
    } else {
        $defaultId = [string]$MenuConfig.DefaultCategory
        $categories = @($MenuConfig.Categories)
    }

    $byId = @{}
    foreach ($t in $toolList) { $byId[[string]$t.Id] = $t }

    $resolvedCategories = New-Object "System.Collections.Generic.List[object]"
    $listed = @{}
    foreach ($c in $categories) {
        $resolvedTools = New-Object "System.Collections.Generic.List[object]"
        foreach ($id in @($c.ToolIds)) {
            $listed[[string]$id] = $true
            if ($byId.ContainsKey([string]$id)) {
                $resolvedTools.Add($byId[[string]$id])
            } else {
                $result.Warnings += ("configured tool '{0}' not found; ignored" -f $id)
            }
        }
        $resolvedCategories.Add(@{
            Id = [string]$c.Id
            Label = [string]$c.Label
            Tools = $resolvedTools
        })
    }

    $defaultCategory = @($resolvedCategories | Where-Object { $_.Id -eq $defaultId })[0]
    foreach ($t in $toolList) {
        if (-not $listed.ContainsKey([string]$t.Id)) {
            $defaultCategory.Tools.Add($t)
            $result.Warnings += ("tool '{0}' is not listed; added to default category '{1}'" -f $t.Id, $defaultId)
        }
    }

    $oversized = @{}
    foreach ($t in $toolList) {
        $cost = Get-ToolMenuCost $t
        if ($cost -gt $script:MenuLimit) {
            $oversized[[string]$t.Id] = $true
            $result.Rejected += ("tool '{0}' needs {1} menu slots; maximum is {2}; skipped" -f $t.Id, $cost, $script:MenuLimit)
        }
    }
    $result.RegisteredCount = $toolList.Count - $oversized.Count

    $pages = New-Object "System.Collections.Generic.List[object]"
    foreach ($ctx in $script:Contexts) {
        foreach ($c in $resolvedCategories) {
            $pageTools = @()
            $pageCost = 0
            $pageNumber = 1
            $contextTools = @($c.Tools | Where-Object {
                $_.On -contains $ctx.on -and -not $oversized.ContainsKey([string]$_.Id)
            })
            foreach ($t in $contextTools) {
                $cost = Get-ToolMenuCost $t
                if ($pageTools.Count -gt 0 -and ($pageCost + $cost) -gt $script:MenuLimit) {
                    if ($pageNumber -eq 1 -and $c.Id -eq $defaultId) {
                        $keyName = $script:KeyName
                    } elseif ($pageNumber -eq 1) {
                        $keyName = $script:KeyName + "." + $c.Id
                    } else {
                        $keyName = $script:KeyName + "." + $c.Id + ".p" + $pageNumber
                    }
                    if ($pageNumber -eq 1) { $label = $c.Label }
                    else { $label = $c.Label + " " + $pageNumber }
                    $pages.Add(@{
                        Context = $ctx.on; Base = $ctx.base; CategoryId = $c.Id
                        KeyName = $keyName; Label = $label; PageNumber = $pageNumber
                        Tools = @($pageTools); Cost = $pageCost
                    })
                    $pageNumber++
                    $pageTools = @()
                    $pageCost = 0
                }
                $pageTools += $t
                $pageCost += $cost
            }
            if ($pageTools.Count -gt 0) {
                if ($pageNumber -eq 1 -and $c.Id -eq $defaultId) {
                    $keyName = $script:KeyName
                } elseif ($pageNumber -eq 1) {
                    $keyName = $script:KeyName + "." + $c.Id
                } else {
                    $keyName = $script:KeyName + "." + $c.Id + ".p" + $pageNumber
                }
                if ($pageNumber -eq 1) { $label = $c.Label }
                else { $label = $c.Label + " " + $pageNumber }
                $pages.Add(@{
                    Context = $ctx.on; Base = $ctx.base; CategoryId = $c.Id
                    KeyName = $keyName; Label = $label; PageNumber = $pageNumber
                    Tools = @($pageTools); Cost = $pageCost
                })
            }
        }
    }
    $result.Pages = $pages.ToArray()
    return $result
}

function ConvertTo-RegString {
    param([string]$s)
    return ($s -replace "\\", "\\" -replace '"', '\"')
}

function Get-ToolCommand {
    param([string]$ToolDir, [int]$VariantIndex, [string]$Window)
    if ($Window -in @("hidden", "gui")) {
        return ('wscript.exe "{0}" "{1}" {2} "%V"' -f $script:SilentVbs, $ToolDir, $VariantIndex)
    }
    return ('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -Tool "{1}" -Variant {2} -Target "%V"' -f $script:LaunchPs1, $ToolDir, $VariantIndex)
}

function Test-OwnedMenuRootPath {
    param([string]$Path)
    foreach ($ctx in $script:Contexts) {
        $prefix = $ctx.base + "\"
        if ($Path.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $name = $Path.Substring($prefix.Length)
            if ($name -eq $script:KeyName -or
                $name.StartsWith($script:KeyName + ".", [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
    }
    return $false
}

function Get-OwnedMenuRootPaths {
    $roots = New-Object "System.Collections.Generic.List[string]"
    $hivePrefix = "HKEY_CURRENT_USER\"
    foreach ($ctx in $script:Contexts) {
        $key = $null
        try {
            $relative = $ctx.base.Substring($hivePrefix.Length)
            $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($relative)
            if ($null -eq $key) { continue }
            foreach ($name in $key.GetSubKeyNames()) {
                if ($name -eq $script:KeyName -or
                    $name.StartsWith($script:KeyName + ".", [System.StringComparison]::OrdinalIgnoreCase)) {
                    $roots.Add($ctx.base + "\" + $name)
                }
            }
        } catch {
            # A missing or unreadable context is equivalent to no owned keys there.
        } finally {
            if ($null -ne $key) { $key.Dispose() }
        }
    }
    return @($roots | Sort-Object -Unique)
}

function Build-RegLines {
    param($Pages, $OwnedRootPaths = @())
    $reg = New-Object "System.Collections.Generic.List[string]"
    $reg.Add("Windows Registry Editor Version 5.00")
    $reg.Add("")

    $deleteSeen = @{}
    $deletePaths = New-Object "System.Collections.Generic.List[string]"
    foreach ($ctx in $script:Contexts) {
        $path = $ctx.base + "\" + $script:KeyName
        $deleteSeen[$path] = $true
        $deletePaths.Add($path)
    }
    foreach ($path in @($OwnedRootPaths) + @($Pages | ForEach-Object { $_.Base + "\" + $_.KeyName })) {
        if ($null -ne $path -and (Test-OwnedMenuRootPath ([string]$path)) -and
            -not $deleteSeen.ContainsKey([string]$path)) {
            $deleteSeen[[string]$path] = $true
            $deletePaths.Add([string]$path)
        }
    }
    foreach ($path in $deletePaths) {
        $reg.Add("[-" + $path + "]")
        $reg.Add("")
    }

    foreach ($page in @($Pages)) {
        $base = $page.Base + "\" + $page.KeyName
        $reg.Add("[" + $base + "]")
        $reg.Add('"MUIVerb"="' + (ConvertTo-RegString $page.Label) + '"')
        $reg.Add('"subcommands"=""')
        $reg.Add("")
        foreach ($t in @($page.Tools)) {
            $labels = @($t.VariantLabels)
            if ($labels.Count -eq 0) {
                $reg.Add("[" + $base + "\shell\" + $t.Id + "]")
                $reg.Add('"MUIVerb"="' + (ConvertTo-RegString $t.Name) + '"')
                $reg.Add("")
                $reg.Add("[" + $base + "\shell\" + $t.Id + "\command]")
                $reg.Add('@="' + (ConvertTo-RegString (Get-ToolCommand $t.Dir -1 $t.Window)) + '"')
                $reg.Add("")
            } else {
                $reg.Add("[" + $base + "\shell\" + $t.Id + "]")
                $reg.Add('"MUIVerb"="' + (ConvertTo-RegString $t.Name) + '"')
                $reg.Add('"subcommands"=""')
                $reg.Add("")
                for ($i = 0; $i -lt $labels.Count; $i++) {
                    $reg.Add("[" + $base + "\shell\" + $t.Id + "\shell\v" + $i + "]")
                    $reg.Add('"MUIVerb"="' + (ConvertTo-RegString $labels[$i]) + '"')
                    $reg.Add("")
                    $reg.Add("[" + $base + "\shell\" + $t.Id + "\shell\v" + $i + "\command]")
                    $reg.Add('@="' + (ConvertTo-RegString (Get-ToolCommand $t.Dir $i $t.Window)) + '"')
                    $reg.Add("")
                }
            }
        }
    }
    return $reg
}

function Get-InstallContextValue {
    param($Context, [string]$Name, $Default)
    if ($null -eq $Context) { return ,$Default }
    if ($Context -is [Collections.IDictionary] -and $Context.Contains($Name)) { return ,$Context[$Name] }
    $property = $Context.PSObject.Properties[$Name]
    if ($property) { return ,$property.Value }
    return ,$Default
}

function New-InstallContext {
    [CmdletBinding()]
    param(
        [string]$Root = $script:Root,
        [string]$LocalRoot = (Join-Path $env:LOCALAPPDATA "ToolRack"),
        [string]$HostNamespace = "default",
        [string]$RunRegistryPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run",
        [string]$RunValueName = "ToolRackHost",
        $MenuContexts = $script:DefaultContexts
    )
    return @{
        Root = [IO.Path]::GetFullPath($Root).TrimEnd("\")
        LocalRoot = [IO.Path]::GetFullPath($LocalRoot)
        HostNamespace = $HostNamespace
        RunRegistryPath = $RunRegistryPath
        RunValueName = $RunValueName
        MenuContexts = @($MenuContexts)
        TestFailSelfTest = $false
        TestFailReady = $false
        TestFailRegistry = $false
    }
}

function Resolve-InstallContext {
    param($InstallContext)
    $context = New-InstallContext
    if ($null -ne $InstallContext) {
        foreach ($name in @("Root", "LocalRoot", "HostNamespace", "RunRegistryPath", "RunValueName", "MenuContexts",
                "TestFailSelfTest", "TestFailReady", "TestFailRegistry")) {
            $context[$name] = Get-InstallContextValue $InstallContext $name $context[$name]
        }
    }
    $context.Root = [IO.Path]::GetFullPath([string]$context.Root).TrimEnd("\")
    $context.LocalRoot = [IO.Path]::GetFullPath([string]$context.LocalRoot)
    $context.MenuContexts = @($context.MenuContexts)
    if ([string]$context.HostNamespace -cnotmatch '^[A-Za-z0-9-]{1,80}$') { throw "Host namespace is invalid" }
    if ([string]$context.RunValueName -cnotmatch '^[A-Za-z0-9._-]{1,80}$') { throw "Run value name is invalid" }
    if ([string]$context.RunRegistryPath -cnotmatch '^HKEY_CURRENT_USER\\[^\x00-\x1F]+$') { throw "Run registry path is invalid" }
    if ($context.MenuContexts.Count -ne 3) { throw "exactly three menu contexts are required" }
    foreach ($item in $context.MenuContexts) {
        if ($item.on -notin @("file", "folder", "background") -or
            [string]$item.base -cnotmatch '^HKEY_CURRENT_USER\\[^\x00-\x1F]+$') {
            throw "menu context is invalid"
        }
    }
    Assert-SafeLocalRoot $context.LocalRoot
    return $context
}

function Assert-SafeLocalRoot {
    param([string]$Path)
    $full = [IO.Path]::GetFullPath($Path).TrimEnd("\")
    if ((Split-Path $full -Leaf) -cne "ToolRack") { throw "local root must end in ToolRack" }
    if ($full -eq [IO.Path]::GetPathRoot($full).TrimEnd("\")) { throw "local root cannot be a drive root" }
}

function Set-InstallScriptContext {
    param($Context)
    $previous = @{
        Root = $script:Root; ToolRoot = $script:ToolRoot; MenuPath = $script:MenuPath; BindingsPath = $script:BindingsPath
        LaunchPs1 = $script:LaunchPs1; SilentVbs = $script:SilentVbs; Contexts = @($script:Contexts)
    }
    $script:Root = $Context.Root
    $script:ToolRoot = Join-Path $script:Root "tool"
    $script:MenuPath = Join-Path $script:Root "menu.json"
    $script:BindingsPath = Join-Path $script:Root "bindings.json"
    $script:LaunchPs1 = Join-Path $script:Root "common\launch.ps1"
    $script:SilentVbs = Join-Path $script:Root "common\silent.vbs"
    $script:Contexts = @($Context.MenuContexts)
    return $previous
}

function Restore-InstallScriptContext {
    param($Previous)
    $script:Root = $Previous.Root
    $script:ToolRoot = $Previous.ToolRoot
    $script:MenuPath = $Previous.MenuPath
    $script:BindingsPath = $Previous.BindingsPath
    $script:LaunchPs1 = $Previous.LaunchPs1
    $script:SilentVbs = $Previous.SilentVbs
    $script:Contexts = @($Previous.Contexts)
}

function ConvertTo-ProcessArgument {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value -or $Value.Length -eq 0) { return '""' }
    if ($Value -notmatch '[\s"]') { return $Value }
    $builder = New-Object Text.StringBuilder
    [void]$builder.Append('"')
    $slashes = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq '\') { $slashes++ }
        elseif ($character -eq '"') {
            [void]$builder.Append('\' * (($slashes * 2) + 1))
            [void]$builder.Append('"')
            $slashes = 0
        } else {
            if ($slashes -gt 0) { [void]$builder.Append('\' * $slashes); $slashes = 0 }
            [void]$builder.Append($character)
        }
    }
    if ($slashes -gt 0) { [void]$builder.Append('\' * ($slashes * 2)) }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Join-ProcessArguments {
    param([string[]]$Items)
    return (@($Items | ForEach-Object { ConvertTo-ProcessArgument $_ }) -join " ")
}

function Invoke-ChildProcess {
    param([string]$FileName, [string[]]$Arguments, [int]$TimeoutMilliseconds = 15000)
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $FileName
    $info.Arguments = Join-ProcessArguments $Arguments
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($info)
    try {
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutMilliseconds)) {
            try { $process.Kill() } catch { }
            return [pscustomobject]@{ ExitCode = -1; StdOut = ""; StdErr = "process timed out"; TimedOut = $true }
        }
        return [pscustomobject]@{
            ExitCode = $process.ExitCode
            StdOut = $stdoutTask.Result
            StdErr = $stderrTask.Result
            TimedOut = $false
        }
    } finally {
        $process.Dispose()
    }
}

function Invoke-HiddenPowerShell {
    param([string[]]$Arguments, [int]$TimeoutMilliseconds = 15000)
    return Invoke-ChildProcess -FileName "powershell.exe" -Arguments $Arguments -TimeoutMilliseconds $TimeoutMilliseconds
}

function Write-StrictUtf8Text {
    param([string]$Path, [string]$Text)
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    [IO.File]::WriteAllText($Path, $Text, (New-Object Text.UTF8Encoding($false, $true)))
}

function Write-StrictUtf8Json {
    param([string]$Path, $Value)
    Write-StrictUtf8Text $Path (ConvertTo-Json $Value -Depth 12)
}

function Write-AtomicStrictUtf8Text {
    param([string]$Path, [string]$Text)
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    $temporary = Join-Path $parent ((Split-Path $Path -Leaf) + ".tmp." + [guid]::NewGuid().ToString("N"))
    try {
        Write-StrictUtf8Text $temporary $Text
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            $backup = $Path + ".bak." + [guid]::NewGuid().ToString("N")
            try { [IO.File]::Replace($temporary, $Path, $backup) }
            finally { Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue }
        } else { [IO.File]::Move($temporary, $Path) }
    } finally {
        Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
    }
}

function Get-InstallPreflight {
    param([string]$BindingPath)
    $menu = Read-MenuConfig $script:MenuPath
    if (-not $menu.Ok) {
        Write-Host "  ERROR menu.json is invalid; registry was not changed." -ForegroundColor Red
        foreach ($msg in @($menu.Errors)) { Write-Host ("        - {0}" -f $msg) -ForegroundColor Red }
        return [pscustomobject]@{ Ok = $false }
    }

    $valid = New-Object "System.Collections.Generic.List[object]"
    $dirs = @()
    if (Test-Path -LiteralPath $script:ToolRoot) { $dirs = @(Get-ChildItem -LiteralPath $script:ToolRoot -Directory | Sort-Object Name) }
    foreach ($directory in $dirs) {
        $read = Read-Manifest $directory.FullName
        $errors = @($read.Errors)
        if ($read.Ok) { $errors = @(Test-Manifest $read.Data $directory.Name $directory.FullName) }
        if ($errors.Count -gt 0) {
            Write-Host ("  SKIP {0}" -f $directory.Name) -ForegroundColor Yellow
            foreach ($msg in $errors) { Write-Host ("       - {0}" -f $msg) -ForegroundColor Yellow }
            continue
        }
        $manifest = $read.Data
        $run = Get-Prop $manifest "run"
        $window = Get-Prop $run "window"
        if ($null -eq $window) { $window = "console" }
        $labels = @()
        if (Test-Prop $manifest "variants") {
            $variantValues = Get-Prop $manifest "variants"
            $labels = @(@($variantValues) | ForEach-Object { [string](Get-Prop $_ "label") })
        }
        $onValues = Get-Prop $manifest "on"
        $valid.Add(@{
            Id = [string](Get-Prop $manifest "id"); Name = [string](Get-Prop $manifest "name")
            On = @($onValues); Window = $window; VariantLabels = $labels; Dir = $directory.FullName
            SupportsActions = [bool]$read.SupportsActions; ActionIds = @($read.ActionIds)
        })
        Write-Host ("  OK   {0} ({1} variants)" -f $directory.Name, $labels.Count) -ForegroundColor Green
    }

    $bindingConfig = Read-BindingsConfig -Path $BindingPath
    if (-not $bindingConfig.Ok) {
        Write-Host "  ERROR bindings.json is invalid; registry was not changed." -ForegroundColor Red
        foreach ($msg in @($bindingConfig.Errors)) { Write-Host ("        - {0}" -f $msg) -ForegroundColor Red }
        return [pscustomobject]@{ Ok = $false }
    }
    $bindingResolution = Resolve-Bindings -Config $bindingConfig.Data -Tools $valid.ToArray()
    if (-not $bindingResolution.Ok) {
        Write-Host "  ERROR bindings could not be resolved; registry was not changed." -ForegroundColor Red
        foreach ($msg in @($bindingResolution.Errors)) { Write-Host ("        - {0}" -f $msg) -ForegroundColor Red }
        return [pscustomobject]@{ Ok = $false }
    }
    foreach ($msg in @($bindingResolution.Warnings)) { Write-Host ("  WARN binding {0}" -f $msg) -ForegroundColor Yellow }
    Write-Host ("  BIND {0} active, {1} rejected" -f @($bindingResolution.Active).Count,
        @($bindingResolution.Rejected).Count) -ForegroundColor DarkGray

    $layout = Resolve-MenuLayout $valid $menu.Data
    if (-not $menu.Exists) { Write-Host "  WARN menu.json not found; using one default category with automatic paging." -ForegroundColor Yellow }
    foreach ($msg in @($layout.Warnings)) { Write-Host ("  WARN {0}" -f $msg) -ForegroundColor Yellow }
    foreach ($msg in @($layout.Rejected)) { Write-Host ("  SKIP {0}" -f $msg) -ForegroundColor Yellow }
    foreach ($page in @($layout.Pages)) {
        Write-Host ("  MENU [{0}] {1} ({2} tools, {3}/{4} slots)" -f $page.Context, $page.Label,
            @($page.Tools).Count, $page.Cost, $script:MenuLimit) -ForegroundColor DarkGray
    }
    return [pscustomobject]@{ Ok = $true; Valid = $valid.ToArray(); Layout = $layout; BindingResolution = $bindingResolution }
}

function New-ResolvedConfigForInstall {
    param($Context, [string]$BindingPath, [string]$OutputPath)
    $resolver = Join-Path $Context.Root "common\resolve-host-config.ps1"
    if (-not (Test-Path -LiteralPath $resolver -PathType Leaf)) { throw "resolver not found: $resolver" }
    $result = Invoke-HiddenPowerShell @("-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File",
        $resolver, "-Root", $Context.Root, "-BindingsPath", $BindingPath, "-OutputPath", $OutputPath)
    if ($result.ExitCode -ne 0) { throw ("binding resolver failed: " + $result.StdErr.Trim()) }
    if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) { throw "binding resolver produced no output" }
}

function Invoke-HostSelfTestForInstall {
    param($Context, [string]$ResolvedPath, [string]$TemporaryRoot)
    $state = Join-Path $TemporaryRoot "selftest-state"
    New-Item -ItemType Directory -Force -Path $state | Out-Null
    Write-StrictUtf8Json (Join-Path $state "host.json") ([ordered]@{
        schema = 1; version = "1"; root = $Context.Root; namespace = $Context.HostNamespace
    })
    Copy-Item -LiteralPath $ResolvedPath -Destination (Join-Path $state "bindings.resolved.json")
    $hostScript = Join-Path $Context.Root "common\host.ps1"
    $arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $hostScript, "-SelfTest", "-StateRoot", $state)
    if ([bool]$Context.TestFailSelfTest) { $arguments += "-TestBadSource" }
    $result = Invoke-HiddenPowerShell $arguments 20000
    if ($result.ExitCode -ne 0) { throw ("Host SelfTest failed: " + $result.StdErr.Trim()) }
    try { $status = ConvertFrom-Json $result.StdOut.Trim() } catch { throw "Host SelfTest returned invalid JSON" }
    if (-not [bool]$status.ok) { throw "Host SelfTest did not report success" }
}

function New-LocalSnapshot {
    param([string]$LocalRoot, [string]$TemporaryRoot)
    $exists = Test-Path -LiteralPath $LocalRoot -PathType Container
    $backup = Join-Path $TemporaryRoot "local-backup"
    if ($exists) { Copy-Item -LiteralPath $LocalRoot -Destination $backup -Recurse -Force }
    return [pscustomobject]@{ Existed = $exists; Backup = $backup }
}

function Remove-OwnedLocalRoot {
    param([string]$LocalRoot)
    Assert-SafeLocalRoot $LocalRoot
    if (Test-Path -LiteralPath $LocalRoot) { Remove-Item -LiteralPath $LocalRoot -Recurse -Force }
}

function Restore-LocalSnapshot {
    param([string]$LocalRoot, $Snapshot)
    Remove-OwnedLocalRoot $LocalRoot
    if ($Snapshot.Existed) { Copy-Item -LiteralPath $Snapshot.Backup -Destination $LocalRoot -Recurse -Force }
}

function Test-FileBytesEqual {
    param([string]$First, [string]$Second)
    if (-not (Test-Path -LiteralPath $First -PathType Leaf) -or -not (Test-Path -LiteralPath $Second -PathType Leaf)) { return $false }
    return [Convert]::ToBase64String([IO.File]::ReadAllBytes($First)) -ceq [Convert]::ToBase64String([IO.File]::ReadAllBytes($Second))
}

function Ensure-VersionedBootstrap {
    param($Context)
    $bootstrapParent = Join-Path $Context.LocalRoot "bootstrap"
    $destination = Join-Path $bootstrapParent "v1"
    $sourceBootstrap = Join-Path $Context.Root "common\host-bootstrap.ps1"
    $sourceVbs = Join-Path $Context.Root "common\host-start.vbs"
    if (Test-Path -LiteralPath $destination -PathType Container) {
        if (-not (Test-FileBytesEqual $sourceBootstrap (Join-Path $destination "host-bootstrap.ps1")) -or
            -not (Test-FileBytesEqual $sourceVbs (Join-Path $destination "start-host.vbs"))) {
            throw "bootstrap v1 already exists with different content"
        }
        return $destination
    }
    New-Item -ItemType Directory -Force -Path $bootstrapParent | Out-Null
    $stage = Join-Path $bootstrapParent ("v1.stage." + [guid]::NewGuid().ToString("N"))
    try {
        New-Item -ItemType Directory -Path $stage | Out-Null
        Copy-Item -LiteralPath $sourceBootstrap -Destination (Join-Path $stage "host-bootstrap.ps1")
        Copy-Item -LiteralPath $sourceVbs -Destination (Join-Path $stage "start-host.vbs")
        Move-Item -LiteralPath $stage -Destination $destination
    } finally {
        Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
    }
    return $destination
}

function New-HostGeneration {
    param($Context, [string]$ResolvedPath)
    $generation = [guid]::NewGuid().ToString("N")
    $stateParent = Join-Path $Context.LocalRoot "state"
    New-Item -ItemType Directory -Force -Path $stateParent | Out-Null
    $stage = Join-Path $stateParent ($generation + ".stage")
    $destination = Join-Path $stateParent $generation
    try {
        New-Item -ItemType Directory -Path $stage | Out-Null
        Write-StrictUtf8Text (Join-Path $stage "root.txt") $Context.Root
        Write-StrictUtf8Json (Join-Path $stage "host.json") ([ordered]@{
            schema = 1; version = "1"; root = $Context.Root; namespace = $Context.HostNamespace
        })
        Copy-Item -LiteralPath $ResolvedPath -Destination (Join-Path $stage "bindings.resolved.json")
        Move-Item -LiteralPath $stage -Destination $destination
    } finally {
        Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
    }
    return [pscustomobject]@{ Generation = $generation; StateRoot = $destination }
}

function Read-StrictUtf8InstallText {
    param([string]$Path)
    return [IO.File]::ReadAllText($Path, (New-Object Text.UTF8Encoding($false, $true)))
}

function Get-ActiveStateRoot {
    param([string]$LocalRoot)
    $activePath = Join-Path $LocalRoot "active.txt"
    if (-not (Test-Path -LiteralPath $activePath -PathType Leaf)) { return $null }
    try { $generation = (Read-StrictUtf8InstallText $activePath).Trim() } catch { return $null }
    if ($generation -cnotmatch '^[a-f0-9]{32}$') { return $null }
    $state = [IO.Path]::GetFullPath((Join-Path $LocalRoot ("state\" + $generation)))
    if (-not (Test-Path -LiteralPath $state -PathType Container)) { return $null }
    return $state
}

function Get-StateRepositoryRoot {
    param([string]$StateRoot, [string]$FallbackRoot)
    $rootPath = Join-Path $StateRoot "root.txt"
    if (-not (Test-Path -LiteralPath $rootPath -PathType Leaf)) { return $FallbackRoot }
    try { return [IO.Path]::GetFullPath((Read-StrictUtf8InstallText $rootPath).Trim()) } catch { return $FallbackRoot }
}

function Get-HostStatusInternal {
    param([string]$StateRoot, [string]$FallbackRoot, [int]$PipeTimeout = 300)
    if ([string]::IsNullOrWhiteSpace($StateRoot)) { return $null }
    $repositoryRoot = Get-StateRepositoryRoot $StateRoot $FallbackRoot
    $control = Join-Path $repositoryRoot "common\host-control.ps1"
    if (-not (Test-Path -LiteralPath $control -PathType Leaf)) { return $null }
    $result = Invoke-HiddenPowerShell @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $control, "-Status",
        "-StateRoot", $StateRoot, "-TimeoutMilliseconds", ([string]$PipeTimeout)) 4000
    if ($result.ExitCode -ne 0) { return $null }
    $line = @($result.StdOut -split "`r?`n" | Where-Object { $_.Trim().StartsWith("{") } | Select-Object -Last 1)
    if ($line.Count -ne 1) { return $null }
    try { return ConvertFrom-Json $line[0] } catch { return $null }
}

function Wait-HostReadyInternal {
    param([string]$StateRoot, [string]$FallbackRoot, [int]$TimeoutMilliseconds = 5000)
    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
    do {
        $status = Get-HostStatusInternal $StateRoot $FallbackRoot 200
        if ($null -ne $status -and [bool]$status.ready) { return $status }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    return $null
}

function Test-HostProcessIdentity {
    param([int]$ProcessId, [string]$StateRoot)
    try {
        $item = Get-CimInstance Win32_Process -Filter ("ProcessId=" + $ProcessId)
        if ($null -eq $item) { return $false }
        $command = [string]$item.CommandLine
        return ($command.IndexOf($StateRoot, [StringComparison]::OrdinalIgnoreCase) -ge 0 -and
            ($command.IndexOf("host.ps1", [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
             $command.IndexOf("host-bootstrap.ps1", [StringComparison]::OrdinalIgnoreCase) -ge 0))
    } catch { return $false }
}

function Stop-HostInternal {
    param([string]$StateRoot, [string]$FallbackRoot)
    $status = Get-HostStatusInternal $StateRoot $FallbackRoot 300
    if ($null -eq $status) { return [pscustomobject]@{ WasRunning = $false; Stopped = $true; ProcessId = 0 } }
    $processId = [int]$status.pid
    $repositoryRoot = Get-StateRepositoryRoot $StateRoot $FallbackRoot
    $control = Join-Path $repositoryRoot "common\host-control.ps1"
    [void](Invoke-HiddenPowerShell @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $control, "-Shutdown",
        "-StateRoot", $StateRoot, "-TimeoutMilliseconds", "1000") 3000)
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($null -ne $process) {
        [void]$process.WaitForExit(2000)
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    }
    if ($null -ne $process -and (Test-HostProcessIdentity $processId $StateRoot)) {
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 100
    }
    $stopped = $null -eq (Get-Process -Id $processId -ErrorAction SilentlyContinue)
    return [pscustomobject]@{ WasRunning = $true; Stopped = $stopped; ProcessId = $processId }
}

function Start-HostInternal {
    param($Context, [string]$StateRoot)
    $bootstrap = Join-Path $Context.LocalRoot "bootstrap\v1\host-bootstrap.ps1"
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = "powershell.exe"
    $info.Arguments = Join-ProcessArguments @("-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
        "-File", $bootstrap, "-StateDirectory", $StateRoot)
    $info.UseShellExecute = $true
    $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    return [Diagnostics.Process]::Start($info)
}

function ConvertTo-RegistryShortPath {
    param([string]$Path)
    $prefix = "HKEY_CURRENT_USER\"
    if (-not $Path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw "only HKCU registry paths are allowed" }
    return "HKCU\" + $Path.Substring($prefix.Length)
}

function Invoke-RegCommand {
    param([string[]]$Arguments)
    $oldEap = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        & reg.exe @Arguments 2>$null | Out-Null
        return $LASTEXITCODE
    } finally { $ErrorActionPreference = $oldEap }
}

function Import-RegistryLines {
    param([string[]]$Lines, [string]$TemporaryRoot)
    $path = Join-Path $TemporaryRoot ("registry." + [guid]::NewGuid().ToString("N") + ".reg")
    try {
        [IO.File]::WriteAllLines($path, $Lines, [Text.Encoding]::Unicode)
        return (Invoke-RegCommand @("import", $path)) -eq 0
    } finally { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
}

function Get-RegistryRunBackup {
    param($Context)
    $prefix = "HKEY_CURRENT_USER\"
    $subKey = $Context.RunRegistryPath.Substring($prefix.Length)
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($subKey, $false)
    try {
        if ($null -eq $key -or ($key.GetValueNames() -notcontains $Context.RunValueName)) {
            return [pscustomobject]@{ Exists = $false; Value = $null; Kind = $null }
        }
        return [pscustomobject]@{
            Exists = $true
            Value = $key.GetValue($Context.RunValueName, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            Kind = $key.GetValueKind($Context.RunValueName)
        }
    } finally { if ($null -ne $key) { $key.Dispose() } }
}

function Backup-RegistryState {
    param($Context, [string[]]$MenuTargets, [string]$TemporaryRoot)
    $exports = New-Object "System.Collections.Generic.List[string]"
    $index = 0
    foreach ($path in @($MenuTargets | Sort-Object -Unique)) {
        $short = ConvertTo-RegistryShortPath $path
        if ((Invoke-RegCommand @("query", $short)) -ne 0) { continue }
        $file = Join-Path $TemporaryRoot ("menu-backup-" + $index + ".reg")
        if ((Invoke-RegCommand @("export", $short, $file, "/y")) -ne 0) { throw "failed to back up registry key: $short" }
        [void]$exports.Add($file)
        $index++
    }
    return [pscustomobject]@{ MenuTargets = @($MenuTargets | Sort-Object -Unique); Exports = $exports.ToArray(); Run = Get-RegistryRunBackup $Context }
}

function Restore-RegistryState {
    param($Context, $Backup, [string]$TemporaryRoot)
    $deleteLines = New-Object "System.Collections.Generic.List[string]"
    $deleteLines.Add("Windows Registry Editor Version 5.00")
    $deleteLines.Add("")
    foreach ($path in @($Backup.MenuTargets)) { $deleteLines.Add("[-" + $path + "]"); $deleteLines.Add("") }
    if (-not (Import-RegistryLines $deleteLines.ToArray() $TemporaryRoot)) { throw "failed to clear registry during rollback" }
    foreach ($file in @($Backup.Exports)) {
        if ((Invoke-RegCommand @("import", $file)) -ne 0) { throw "failed to restore registry menu backup" }
    }
    $prefix = "HKEY_CURRENT_USER\"
    $subKey = $Context.RunRegistryPath.Substring($prefix.Length)
    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($subKey)
    try {
        if ($Backup.Run.Exists) { $key.SetValue($Context.RunValueName, $Backup.Run.Value, $Backup.Run.Kind) }
        else { $key.DeleteValue($Context.RunValueName, $false) }
    } finally { $key.Dispose() }
}

function Get-MenuRegistryTargets {
    param($Pages, [string[]]$Owned)
    $targets = New-Object "System.Collections.Generic.List[string]"
    foreach ($context in $script:Contexts) { [void]$targets.Add($context.base + "\" + $script:KeyName) }
    foreach ($path in @($Owned)) { if ($null -ne $path) { [void]$targets.Add([string]$path) } }
    foreach ($page in @($Pages)) { [void]$targets.Add($page.Base + "\" + $page.KeyName) }
    return @($targets.ToArray() | Sort-Object -Unique)
}

function Get-RunCommand {
    param($Context)
    $wscript = Join-Path $env:SystemRoot "System32\wscript.exe"
    $startVbs = Join-Path $Context.LocalRoot "bootstrap\v1\start-host.vbs"
    $command = '"' + $wscript + '" "' + $startVbs + '"'
    if ($command.Length -ge 260) { throw "Run command exceeds 259 characters" }
    return $command
}

function Build-InstallRegistryLines {
    param($Pages, [string[]]$Owned, $Context, [string]$RunCommand)
    $lines = New-Object "System.Collections.Generic.List[string]"
    foreach ($line in @(Build-RegLines $Pages $Owned)) { $lines.Add([string]$line) }
    $lines.Add("[" + $Context.RunRegistryPath + "]")
    $lines.Add('"' + (ConvertTo-RegString $Context.RunValueName) + '"="' + (ConvertTo-RegString $RunCommand) + '"')
    $lines.Add("")
    return $lines.ToArray()
}

function Remove-OldHostGenerations {
    param($Context, [string]$CurrentGeneration, [AllowNull()][string]$PreviousGeneration)
    $stateParent = Join-Path $Context.LocalRoot "state"
    if (-not (Test-Path -LiteralPath $stateParent -PathType Container)) { return }
    foreach ($directory in @(Get-ChildItem -LiteralPath $stateParent -Directory)) {
        if ($directory.Name -ceq $CurrentGeneration -or
            (-not [string]::IsNullOrWhiteSpace($PreviousGeneration) -and $directory.Name -ceq $PreviousGeneration)) { continue }
        if ($directory.Name -cmatch '^[a-f0-9]{32}$') { Remove-Item -LiteralPath $directory.FullName -Recurse -Force }
    }
}

function Invoke-Install {
    [CmdletBinding()]
    param($InstallContext = $null, [string]$BindingPath = "")
    $ErrorActionPreference = "Stop"
    Write-Host ""
    Write-Host "=== toolrack install ===" -ForegroundColor Cyan
    Write-Host ""

    $context = $null
    $previous = $null
    $temporaryRoot = Join-Path $env:TEMP ("toolrack_install_" + [guid]::NewGuid().ToString("N"))
    $localSnapshot = $null
    $registryBackup = $null
    $registryTouched = $false
    $oldState = $null
    $oldGeneration = $null
    $oldStopped = $false
    $oldWasRunning = $false
    $newState = $null
    $newProcessId = 0
    try {
        $context = Resolve-InstallContext $InstallContext
        $previous = Set-InstallScriptContext $context
        if ([string]::IsNullOrWhiteSpace($BindingPath)) { $BindingPath = $script:BindingsPath }
        else { $BindingPath = [IO.Path]::GetFullPath($BindingPath) }

        $preflight = Get-InstallPreflight $BindingPath
        if (-not $preflight.Ok) { return [pscustomobject]@{ Ok = $false; Phase = "preflight"; Error = "validation failed" } }

        New-Item -ItemType Directory -Force -Path $temporaryRoot | Out-Null
        $resolvedPath = Join-Path $temporaryRoot "bindings.resolved.json"
        New-ResolvedConfigForInstall $context $BindingPath $resolvedPath
        Invoke-HostSelfTestForInstall $context $resolvedPath $temporaryRoot

        $owned = @(Get-OwnedMenuRootPaths)
        $menuTargets = @(Get-MenuRegistryTargets $preflight.Layout.Pages $owned)
        $registryBackup = Backup-RegistryState $context $menuTargets $temporaryRoot
        $localSnapshot = New-LocalSnapshot $context.LocalRoot $temporaryRoot
        $oldState = Get-ActiveStateRoot $context.LocalRoot
        if ($null -ne $oldState) { $oldGeneration = Split-Path $oldState -Leaf }

        [void](Ensure-VersionedBootstrap $context)
        $newState = New-HostGeneration $context $resolvedPath

        if ($null -ne $oldState) {
            $stopResult = Stop-HostInternal $oldState $context.Root
            $oldWasRunning = [bool]$stopResult.WasRunning
            $oldStopped = [bool]$stopResult.Stopped -and $oldWasRunning
            if ($oldWasRunning -and -not $stopResult.Stopped) { throw "old Host could not be stopped" }
        }

        $newProcess = Start-HostInternal $context $newState.StateRoot
        $newProcessId = $newProcess.Id
        $newProcess.Dispose()
        if ([bool]$context.TestFailReady) { throw "Host ready timeout (test seam)" }
        $ready = Wait-HostReadyInternal $newState.StateRoot $context.Root 5000
        if ($null -eq $ready) { throw "new Host did not become ready within 5 seconds" }

        Write-AtomicStrictUtf8Text (Join-Path $context.LocalRoot "active.txt") $newState.Generation
        $runCommand = Get-RunCommand $context
        $registryLines = Build-InstallRegistryLines $preflight.Layout.Pages $owned $context $runCommand
        $registryTouched = $true
        if (-not (Import-RegistryLines $registryLines $temporaryRoot)) { throw "registry import failed" }
        if ([bool]$context.TestFailRegistry) { throw "registry import failed (test seam)" }

        Remove-OldHostGenerations $context $newState.Generation $oldGeneration
        Write-Host ""
        if ($preflight.Layout.RegisteredCount -gt 0) {
            Write-Host ("Registered {0} tool(s) in {1} menu page(s)." -f $preflight.Layout.RegisteredCount,
                @($preflight.Layout.Pages).Count) -ForegroundColor Cyan
        } else { Write-Host "No valid tools. Cleared all Tool Rack menu entries." -ForegroundColor Yellow }
        Write-Host ("Host ready (PID {0}, generation {1})." -f $ready.pid, $newState.Generation) -ForegroundColor Green
        return [pscustomobject]@{ Ok = $true; Phase = "complete"; StateRoot = $newState.StateRoot; ProcessId = [int]$ready.pid; Generation = $newState.Generation }
    } catch {
        $failure = $_.Exception.Message
        Write-Host ("  ERROR {0}" -f $failure) -ForegroundColor Red
        Write-Host "        registry and Host state will be restored." -ForegroundColor Red
        $restoreErrors = New-Object "System.Collections.Generic.List[string]"
        if ($null -ne $context -and $null -ne $newState) {
            try {
                [void](Stop-HostInternal $newState.StateRoot $context.Root)
                if ($newProcessId -gt 0 -and $null -ne (Get-Process -Id $newProcessId -ErrorAction SilentlyContinue) -and
                    (Test-HostProcessIdentity $newProcessId $newState.StateRoot)) {
                    Stop-Process -Id $newProcessId -Force -ErrorAction SilentlyContinue
                }
            } catch { $restoreErrors.Add("stop new Host: " + $_.Exception.Message) }
        }
        if ($null -ne $context -and $null -ne $localSnapshot) {
            try { Restore-LocalSnapshot $context.LocalRoot $localSnapshot } catch { $restoreErrors.Add("local state: " + $_.Exception.Message) }
        }
        if ($registryTouched -and $null -ne $registryBackup) {
            try { Restore-RegistryState $context $registryBackup $temporaryRoot } catch { $restoreErrors.Add("registry: " + $_.Exception.Message) }
        }
        if ($oldStopped -and $oldWasRunning -and $null -ne $oldState -and (Test-Path -LiteralPath $oldState -PathType Container)) {
            try {
                $oldProcess = Start-HostInternal $context $oldState
                $oldProcess.Dispose()
                if ($null -eq (Wait-HostReadyInternal $oldState $context.Root 5000)) { throw "old Host did not become ready" }
            } catch { $restoreErrors.Add("old Host: " + $_.Exception.Message) }
        }
        if ($restoreErrors.Count -eq 0) { Write-Host "  OK   previous state restored" -ForegroundColor Green }
        else { foreach ($message in $restoreErrors) { Write-Host ("  WARN restore {0}" -f $message) -ForegroundColor Yellow } }
        return [pscustomobject]@{ Ok = $false; Phase = "transaction"; Error = $failure; RestoreOk = ($restoreErrors.Count -eq 0) }
    } finally {
        if ($null -ne $previous) { Restore-InstallScriptContext $previous }
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if ($MyInvocation.InvocationName -ne ".") {
    if ([string]::IsNullOrWhiteSpace($BindingsConfigPath)) { $installResult = Invoke-Install }
    else { $installResult = Invoke-Install -BindingPath $BindingsConfigPath }
    if (-not $installResult.Ok) { exit 1 }
    exit 0
}

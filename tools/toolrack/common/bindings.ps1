# common/bindings.ps1 -- strict global binding parser and target resolver.

function Get-BindingProp {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return ,$property.Value }
    return $null
}

function Test-BindingProp {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $false }
    return ($null -ne $Object.PSObject.Properties[$Name])
}

function Add-UnknownPropertyErrors {
    param($Object, [string[]]$Allowed, [string]$Path, $Errors)
    if ($Object -isnot [pscustomobject]) { return }
    foreach ($property in @($Object.PSObject.Properties)) {
        if ($Allowed -cnotcontains $property.Name) {
            [void]$Errors.Add($Path + "." + $property.Name + ": unknown property")
        }
    }
}

function Get-NormalizedModifiers {
    param($Value, [string]$Path, $Errors)
    if ($Value -isnot [Array]) {
        [void]$Errors.Add($Path + ": must be an array")
        return @()
    }
    $items = @($Value)
    if ($items.Count -eq 0) {
        [void]$Errors.Add($Path + ": at least one modifier is required")
        return @()
    }
    $seen = @{}
    foreach ($item in $items) {
        if ($item -isnot [string] -or $item -notin @("ctrl", "alt", "shift")) {
            [void]$Errors.Add($Path + ": allowed values are ctrl, alt, shift")
            continue
        }
        if ($seen.ContainsKey($item)) {
            [void]$Errors.Add($Path + ": duplicate modifier '" + $item + "'")
            continue
        }
        $seen[$item] = $true
    }
    $normalized = New-Object "System.Collections.Generic.List[string]"
    foreach ($name in @("ctrl", "alt", "shift")) {
        if ($seen.ContainsKey($name)) { [void]$normalized.Add($name) }
    }
    return $normalized.ToArray()
}

function Test-HotkeyKey {
    param([string]$Key)
    if ($Key -cmatch '^[A-Z0-9]$') { return $true }
    if ($Key -cnotmatch '^F([0-9]{1,2})$') { return $false }
    $number = [int]$Matches[1]
    return (($number -ge 1 -and $number -le 11) -or ($number -ge 13 -and $number -le 24))
}

function Format-Trigger {
    param($Trigger)
    $modifiers = @($Trigger.Modifiers) -join "+"
    if ([string]$Trigger.Type -eq "hotkey") {
        return "hotkey:" + $modifiers + "+" + [string]$Trigger.Key
    }
    return "mouse:" + $modifiers + "+" + [string]$Trigger.Button
}

function Read-BindingsConfig {
    param([string]$Path)
    $result = @{ Ok = $false; Data = $null; Errors = @(); Warnings = @() }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Errors = @("bindings.json not found: " + $Path)
        return $result
    }
    $root = $null
    try { $root = ConvertFrom-Json ([IO.File]::ReadAllText($Path)) }
    catch {
        $result.Errors = @("bindings.json parse error: " + $_.Exception.Message)
        return $result
    }
    if ($root -isnot [pscustomobject]) {
        $result.Errors = @("bindings.json root: must be an object")
        return $result
    }

    $errors = New-Object "System.Collections.Generic.List[string]"
    Add-UnknownPropertyErrors $root @("schema", "bindings") "root" $errors
    $schema = Get-BindingProp $root "schema"
    if ($schema -isnot [int] -or $schema -ne 1) {
        [void]$errors.Add("schema: must be integer 1")
    }
    $rawBindings = Get-BindingProp $root "bindings"
    if ($rawBindings -isnot [Array]) {
        [void]$errors.Add("bindings: must be an array")
        $rawBindings = @()
    }

    $normalizedBindings = New-Object "System.Collections.Generic.List[object]"
    $bindingIds = New-Object "System.Collections.Generic.HashSet[string]" ([StringComparer]::Ordinal)
    $triggers = New-Object "System.Collections.Generic.HashSet[string]" ([StringComparer]::Ordinal)
    $bindingItems = @($rawBindings)
    for ($index = 0; $index -lt $bindingItems.Count; $index++) {
        $binding = $bindingItems[$index]
        $pathPrefix = "bindings[" + $index + "]"
        $before = $errors.Count
        if ($binding -isnot [pscustomobject]) {
            [void]$errors.Add($pathPrefix + ": must be an object")
            continue
        }
        Add-UnknownPropertyErrors $binding @("id", "trigger", "invoke") $pathPrefix $errors

        $id = Get-BindingProp $binding "id"
        if ($id -isnot [string] -or $id -cnotmatch '^[a-z0-9][a-z0-9-]*$') {
            [void]$errors.Add($pathPrefix + ".id: invalid")
        } elseif (-not $bindingIds.Add($id)) {
            [void]$errors.Add($pathPrefix + ".id: duplicate '" + $id + "'")
        }

        $normalizedTrigger = $null
        $trigger = Get-BindingProp $binding "trigger"
        if ($trigger -isnot [pscustomobject]) {
            [void]$errors.Add($pathPrefix + ".trigger: must be an object")
        } else {
            Add-UnknownPropertyErrors $trigger @("type", "key", "button", "modifiers") ($pathPrefix + ".trigger") $errors
            $type = Get-BindingProp $trigger "type"
            if ($type -isnot [string] -or $type -notin @("hotkey", "mouse")) {
                [void]$errors.Add($pathPrefix + ".trigger.type: must be hotkey or mouse")
            } else {
                $modifiers = Get-NormalizedModifiers (Get-BindingProp $trigger "modifiers") ($pathPrefix + ".trigger.modifiers") $errors
                if ($type -eq "hotkey") {
                    if (Test-BindingProp $trigger "button") {
                        [void]$errors.Add($pathPrefix + ".trigger.button: not valid for hotkey")
                    }
                    $key = Get-BindingProp $trigger "key"
                    if ($key -isnot [string] -or -not (Test-HotkeyKey $key)) {
                        [void]$errors.Add($pathPrefix + ".trigger.key: unsupported key")
                    } else {
                        $normalizedTrigger = @{ Type = "hotkey"; Key = $key; Modifiers = @($modifiers) }
                    }
                } else {
                    if (Test-BindingProp $trigger "key") {
                        [void]$errors.Add($pathPrefix + ".trigger.key: not valid for mouse")
                    }
                    $button = Get-BindingProp $trigger "button"
                    if ($button -isnot [string] -or $button -cne "right") {
                        [void]$errors.Add($pathPrefix + ".trigger.button: v1 supports right only")
                    } else {
                        $normalizedTrigger = @{ Type = "mouse"; Button = "right"; Modifiers = @($modifiers) }
                    }
                }
            }
        }

        $normalizedInvoke = $null
        $invoke = Get-BindingProp $binding "invoke"
        if ($invoke -isnot [pscustomobject]) {
            [void]$errors.Add($pathPrefix + ".invoke: must be an object")
        } else {
            Add-UnknownPropertyErrors $invoke @("tool", "action") ($pathPrefix + ".invoke") $errors
            $tool = Get-BindingProp $invoke "tool"
            $action = Get-BindingProp $invoke "action"
            if ($tool -isnot [string] -or $tool -notmatch '^[A-Za-z0-9-]+$') {
                [void]$errors.Add($pathPrefix + ".invoke.tool: invalid tool id")
            }
            if ($action -isnot [string] -or $action -cnotmatch '^[a-z0-9][a-z0-9-]*$') {
                [void]$errors.Add($pathPrefix + ".invoke.action: invalid action id")
            }
            if ($tool -is [string] -and $action -is [string]) {
                $normalizedInvoke = @{ Tool = $tool; Action = $action }
            }
        }

        if ($errors.Count -eq $before -and $null -ne $normalizedTrigger -and $null -ne $normalizedInvoke) {
            $triggerKey = Format-Trigger $normalizedTrigger
            if (-not $triggers.Add($triggerKey)) {
                [void]$errors.Add($pathPrefix + ".trigger: duplicate normalized trigger '" + $triggerKey + "'")
            } else {
                [void]$normalizedBindings.Add(@{ Id = $id; Trigger = $normalizedTrigger; Invoke = $normalizedInvoke })
            }
        }
    }

    if ($errors.Count -gt 0) {
        $result.Errors = $errors.ToArray()
        return $result
    }
    $result.Data = @{ Schema = 1; Bindings = $normalizedBindings.ToArray(); SourcePath = [IO.Path]::GetFullPath($Path) }
    $result.Ok = $true
    return $result
}

function Resolve-Bindings {
    param($Config, [object[]]$Tools)
    $result = @{ Ok = $false; Errors = @(); Warnings = @(); Active = @(); Rejected = @() }
    if ($null -eq $Config -or $Config.Schema -ne 1) {
        $result.Errors = @("normalized binding config is invalid")
        return $result
    }
    $active = New-Object "System.Collections.Generic.List[object]"
    $rejected = New-Object "System.Collections.Generic.List[object]"
    foreach ($binding in @($Config.Bindings)) {
        $tool = @($Tools | Where-Object { [string]$_.Id -ceq [string]$binding.Invoke.Tool } | Select-Object -First 1)
        if ($tool.Count -eq 0) {
            [void]$rejected.Add(@{ Id = $binding.Id; Reason = "tool not found: " + $binding.Invoke.Tool })
            continue
        }
        $toolItem = $tool[0]
        if (-not [bool]$toolItem.SupportsActions) {
            [void]$rejected.Add(@{ Id = $binding.Id; Reason = "tool has no stable action IDs: " + $binding.Invoke.Tool })
            continue
        }
        if (@($toolItem.ActionIds) -cnotcontains [string]$binding.Invoke.Action) {
            [void]$rejected.Add(@{ Id = $binding.Id; Reason = "action not found: " + $binding.Invoke.Tool + "/" + $binding.Invoke.Action })
            continue
        }
        [void]$active.Add(@{
            Id = $binding.Id
            Trigger = $binding.Trigger
            Invoke = $binding.Invoke
            ToolDir = [string]$toolItem.Dir
        })
    }
    $result.Active = $active.ToArray()
    $result.Rejected = $rejected.ToArray()
    $result.Warnings = @($rejected.ToArray() | ForEach-Object { $_.Id + ": " + $_.Reason })
    $result.Ok = $true
    return $result
}

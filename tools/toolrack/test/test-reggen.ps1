# test/test-reggen.ps1 -- unit tests for .reg text generation
. (Join-Path $PSScriptRoot "_assert.ps1")
. (Join-Path $PSScriptRoot "..\common\install.ps1")

$tools = @(
    @{ Id="flat"; Name="Flat Tool"; On=@("folder","background"); Window="console"; VariantLabels=@(); Dir="C:\x\tool\flat" },
    @{ Id="fly";  Name="Fly"; On=@("file"); Window="gui"; VariantLabels=@("A","B"); Dir="C:\x\tool\fly" }
)
$menu = @{
    DefaultCategory = "general"
    Categories = @(
        @{ Id="general"; Label="Tool Rack"; ToolIds=@("flat") },
        @{ Id="ai"; Label="Tool Rack AI"; ToolIds=@("fly") }
    )
}
$layout = Resolve-MenuLayout $tools $menu
$stale = @('HKEY_CURRENT_USER\Software\Classes\*\shell\ToolRack.old')
$lines = @(Build-RegLines $layout.Pages $stale)

# key lines: exact match (-contains) -- '[' breaks -like wildcards
Assert-True ($lines -contains "Windows Registry Editor Version 5.00") "reg header"
Assert-True ($lines -contains '[-HKEY_CURRENT_USER\Software\Classes\*\shell\ToolRack]') "delete: file ctx"
Assert-True ($lines -contains '[-HKEY_CURRENT_USER\Software\Classes\Directory\shell\ToolRack]') "delete: folder ctx"
Assert-True ($lines -contains '[-HKEY_CURRENT_USER\Software\Classes\Directory\Background\shell\ToolRack]') "delete: background ctx"
Assert-True ($lines -contains '[-HKEY_CURRENT_USER\Software\Classes\*\shell\ToolRack.old]') "delete: stale owned page"
Assert-True ($lines -contains '"MUIVerb"="Tool Rack"') "parent MUIVerb"
Assert-True ($lines -contains '[HKEY_CURRENT_USER\Software\Classes\Directory\shell\ToolRack\shell\flat\command]') "flat command key"
Assert-True ($lines -contains '[HKEY_CURRENT_USER\Software\Classes\*\shell\ToolRack.ai\shell\fly\shell\v1\command]') "categorized variant v1 key"
# command contents: -like is fine (no brackets in patterns)
Assert-Contains $lines '*-Variant -1 -Target \"%V\"*' "flat uses -Variant -1"
Assert-Contains $lines '*wscript.exe*silent.vbs*' "gui goes through silent.vbs"
$hiddenCommand = Get-ToolCommand "C:\x\tool\hidden" -1 "hidden"
Assert-True ($hiddenCommand -like "wscript.exe*silent.vbs*") "hidden goes through silent.vbs"
# a context with no tools must not create the parent key
$only = @(@{ Id="flat"; Name="F"; On=@("folder"); Window="console"; VariantLabels=@(); Dir="C:\x\t\flat" })
$onlyLayout = Resolve-MenuLayout $only $null
$l2 = @(Build-RegLines $onlyLayout.Pages @())
$fileParent = @($l2 | Where-Object { $_ -eq "[HKEY_CURRENT_USER\Software\Classes\*\shell\ToolRack]" })
Assert-True ($fileParent.Count -eq 0) "no parent key for empty context"
Exit-Test

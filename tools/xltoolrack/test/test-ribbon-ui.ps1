. (Join-Path $PSScriptRoot '_harness.ps1')

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$dir = New-TestDirectory
$excel = $null
$target = $null
$addin = $null
try {
    $addinPath = Build-Addin $dir 'xlam'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $target = $excel.Workbooks.Add()
    $addin = $excel.Workbooks.Open($addinPath, 0, $false)

    Start-Sleep -Seconds 1
    $root = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$excel.Hwnd)
    $all = $root.FindAll(
        [Windows.Automation.TreeScope]::Descendants,
        [Windows.Automation.Condition]::TrueCondition)
    $tab = $null
    for ($index = 0; $index -lt $all.Count; $index++) {
        if ($all.Item($index).Current.Name -eq 'xltoolrack') {
            $tab = $all.Item($index)
            break
        }
    }
    Assert-True ($null -ne $tab) 'live Excel UI contains the xltoolrack Ribbon tab'

    if ($null -ne $tab) {
        $tabParent = [Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($tab)
        $tabCondition = New-Object Windows.Automation.PropertyCondition(
            [Windows.Automation.AutomationElement]::ControlTypeProperty,
            [Windows.Automation.ControlType]::TabItem)
        $ribbonTabs = $tabParent.FindAll([Windows.Automation.TreeScope]::Children, $tabCondition)
        $lastTabName = if ($ribbonTabs.Count -gt 0) { $ribbonTabs.Item($ribbonTabs.Count - 1).Current.Name } else { '' }
        Assert-Equal 'xltoolrack' $lastTabName 'xltoolrack is the rightmost visible Ribbon tab'

        $selection = $tab.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern)
        $selection.Select()
        Start-Sleep -Milliseconds 500
        $all = $root.FindAll(
            [Windows.Automation.TreeScope]::Descendants,
            [Windows.Automation.Condition]::TrueCondition)
        $expected = @('Conway Life', 'Parallel Pi Race', 'Multi Stopwatch')
        $found = @()
        for ($index = 0; $index -lt $all.Count; $index++) {
            $name = $all.Item($index).Current.Name
            if ($expected -contains $name) { $found += $name }
        }
        foreach ($name in $expected) {
            Assert-True ($found -contains $name) "live Ribbon contains $name"
        }
    }
} finally {
    if ($null -ne $addin) {
        try { $addin.Close($false) } catch {}
        Release-Com $addin
    }
    if ($null -ne $target) {
        try { $target.Close($false) } catch {}
        Release-Com $target
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-Com $excel
    }
    Remove-TestDirectory $dir
}

Exit-Test

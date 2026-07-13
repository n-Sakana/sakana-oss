# Timer -- WPF countdown timer. Sample gui-type tool: no console, ignores -Target.
param(
    [string]$Target = "",
    [int]$Minutes = 0
)
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Timer" SizeToContent="WidthAndHeight" ResizeMode="CanMinimize"
        WindowStartupLocation="CenterScreen" Topmost="True">
  <StackPanel Margin="20">
    <TextBlock x:Name="Display" Text="00:00" FontSize="64" FontFamily="Consolas"
               HorizontalAlignment="Center"/>
    <StackPanel x:Name="SetupRow" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,8,0,0">
      <TextBox x:Name="MinBox" Width="64" FontSize="20" Text="5"
               VerticalContentAlignment="Center" HorizontalContentAlignment="Center"/>
      <TextBlock Text=" min" FontSize="20" VerticalAlignment="Center"/>
    </StackPanel>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,12,0,0">
      <Button x:Name="StartBtn" Content="Start" Width="84" Height="34" Margin="4,0"/>
      <Button x:Name="ResetBtn" Content="Reset" Width="84" Height="34" Margin="4,0"/>
    </StackPanel>
    <CheckBox x:Name="TopChk" Content="Always on top" IsChecked="True"
              HorizontalAlignment="Center" Margin="0,12,0,0"/>
  </StackPanel>
</Window>
"@

$win = [System.Windows.Markup.XamlReader]::Parse($xaml)
$display  = $win.FindName("Display")
$setupRow = $win.FindName("SetupRow")
$minBox   = $win.FindName("MinBox")
$startBtn = $win.FindName("StartBtn")
$resetBtn = $win.FindName("ResetBtn")
$topChk   = $win.FindName("TopChk")

$script:Total = 0
$script:Remaining = 0
$script:Running = $false
$script:SetupMode = ($Minutes -le 0)

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)

function Update-Display {
    $m = [math]::Floor($script:Remaining / 60)
    $s = $script:Remaining % 60
    $display.Text = ("{0:00}:{1:00}" -f $m, $s)
}
function Stop-Countdown {
    $timer.Stop()
    $script:Running = $false
    $startBtn.Content = "Start"
}
function Start-Countdown {
    if ($script:Remaining -le 0) { return }
    $display.Foreground = [System.Windows.Media.Brushes]::Black
    $timer.Start()
    $script:Running = $true
    $startBtn.Content = "Pause"
}

$timer.Add_Tick({
    $script:Remaining--
    Update-Display
    if ($script:Remaining -le 0) {
        Stop-Countdown
        $display.Foreground = [System.Windows.Media.Brushes]::Red
        [System.Media.SystemSounds]::Exclamation.Play()
        $win.Topmost = $true
        $win.Activate() | Out-Null
    }
})

$startBtn.Add_Click({
    if ($script:Running) { Stop-Countdown; return }
    if ($script:SetupMode -and $script:Remaining -le 0) {
        $mins = 0
        if (-not [int]::TryParse($minBox.Text, [ref]$mins) -or $mins -lt 1 -or $mins -gt 999) {
            $minBox.Text = "5"; return
        }
        $script:Total = $mins * 60
        $script:Remaining = $script:Total
        $setupRow.Visibility = "Collapsed"
        Update-Display
    }
    Start-Countdown
})

$resetBtn.Add_Click({
    Stop-Countdown
    $display.Foreground = [System.Windows.Media.Brushes]::Black
    if ($script:SetupMode) {
        $script:Remaining = 0
        $setupRow.Visibility = "Visible"
        $display.Text = "00:00"
    } else {
        $script:Remaining = $script:Total
        Update-Display
    }
})

$topChk.Add_Checked({ $win.Topmost = $true })
$topChk.Add_Unchecked({ $win.Topmost = $false })

if (-not $script:SetupMode) {
    if ($Minutes -gt 999) { $Minutes = 999 }
    $script:Total = $Minutes * 60
    $script:Remaining = $script:Total
    $setupRow.Visibility = "Collapsed"
    Update-Display
    Start-Countdown
}

if ($env:TOOLRACK_SMOKE -eq "1") { exit 0 }
[void]$win.ShowDialog()

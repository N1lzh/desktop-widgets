<#
    WorkClock - how long you have worked today, how much is left of your daily
    target and when you are done.

    Run:  powershell -NoProfile -ExecutionPolicy Bypass -File widgets\workclock.ps1
    Silent launcher: Widget.vbs workclock
#>

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'lib\Widget.psm1') -Force

# --------------------------------------------------------------------------
# state
# --------------------------------------------------------------------------
$script:S = Initialize-Widget -Id 'workclock' -Name 'WorkClock' -StateFile 'state.json' `
    -ScriptPath $PSCommandPath -BeforeSave { $script:S.workedSeconds = [math]::Round((Get-Worked), 1) } `
    -Defaults ([ordered]@{
        date          = (Get-Date).ToString('yyyy-MM-dd')
        workedSeconds = 0.0
        running       = $true
        targetMinutes = 480
    })

# accumulated time = closed segments + the segment currently running
$script:base     = [double]$script:S.workedSeconds
$script:segStart = if ($script:S.running) { Get-Date } else { $null }
$script:lastTick = Get-Date

function Get-Worked {
    if ($null -ne $script:segStart) { $script:base + ((Get-Date) - $script:segStart).TotalSeconds }
    else { $script:base }
}

function Stop-Clock {
    if ($null -ne $script:segStart) {
        $script:base += ((Get-Date) - $script:segStart).TotalSeconds
        $script:segStart = $null
    }
    $script:S.running = $false
}

function Start-Clock {
    if ($null -eq $script:segStart) { $script:segStart = Get-Date }
    $script:S.running = $true
}

function Reset-Day {
    $script:base = 0.0
    $script:S.date = (Get-Date).ToString('yyyy-MM-dd')
    if ($script:S.running) { $script:segStart = Get-Date }
}

# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
$menu = @'
        <MenuItem x:Name="MiToggle"  Header="Pause"/>
        <Separator/>
        <MenuItem x:Name="MiAdjust"  Header="Adjust worked time..."/>
        <MenuItem x:Name="MiTarget"  Header="Set daily target..."/>
        <MenuItem x:Name="MiReset"   Header="Reset today"/>
        <Separator/>
'@

$body = @'
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Grid Grid.Row="0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse x:Name="Dot" Width="7" Height="7" Fill="#4CD07D" Margin="1,0,8,0"/>
          <TextBlock x:Name="StatusText" Text="WORKING" FontSize="9" FontWeight="SemiBold" Foreground="#8DA0B3"/>
        </StackPanel>
        <StackPanel x:Name="Controls" Orientation="Horizontal" HorizontalAlignment="Right" Opacity="0.18">
          <Button x:Name="BtnToggle" Style="{StaticResource Glyph}" Content="&#xE769;" ToolTip="Pause / resume"/>
          <Button x:Name="BtnReset"  Style="{StaticResource Glyph}" Content="&#xE72C;" ToolTip="Reset today"/>
          <Button x:Name="BtnQuit"   Style="{StaticResource Glyph}" Content="&#xE8BB;" ToolTip="Quit"/>
        </StackPanel>
      </Grid>

      <Grid Grid.Row="1" Margin="0,1,0,7">
        <TextBlock x:Name="TimeText" Text="0:00" FontSize="34" FontWeight="Light" Foreground="#F1F6FB"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,1,6">
          <TextBlock x:Name="RemainText" Text="8:00" FontSize="16" Foreground="#79B4FF"/>
          <TextBlock x:Name="RemainLabel" Text=" left" FontSize="11" Foreground="#7B8A99" VerticalAlignment="Bottom" Margin="2,0,0,2"/>
        </StackPanel>
      </Grid>

      <Border Grid.Row="2" x:Name="Track" Height="5" CornerRadius="3" Background="#24FFFFFF">
        <Border x:Name="Fill" HorizontalAlignment="Left" Width="0" CornerRadius="3" Background="#5AA9FF"/>
      </Border>

      <Grid Grid.Row="3" Margin="0,8,0,0">
        <TextBlock x:Name="FootLeft"  Text="of 8:00" FontSize="10.5" Foreground="#74838F"/>
        <TextBlock x:Name="FootRight" HorizontalAlignment="Right" Text="" FontSize="10.5" Foreground="#74838F"/>
      </Grid>
'@

$win = New-WidgetWindow -Body $body -MenuItems $menu
$script:UI = Get-WidgetElements $win @(
    'Dot','StatusText','Controls','BtnToggle','BtnReset','BtnQuit','TimeText','RemainText',
    'RemainLabel','Track','Fill','FootLeft','FootRight','MiToggle','MiAdjust','MiTarget','MiReset')

# --------------------------------------------------------------------------
# rendering
# --------------------------------------------------------------------------
$P = Get-WidgetPalette
$brushAccent = Get-WidgetBrush $P.Accent
$brushDone   = Get-WidgetBrush $P.Good
$brushPaused = Get-WidgetBrush $P.Warn
$brushRemain = Get-WidgetBrush $P.Soft

function Update-Ui {
    $ui        = $script:UI
    $worked    = Get-Worked
    $targetSec = [double]$script:S.targetMinutes * 60
    $left      = $targetSec - $worked
    $done      = $left -le 0

    $ui.TimeText.Text = Format-HM $worked

    if ($done) {
        $ui.RemainText.Text       = Format-HM ([math]::Abs($left))
        $ui.RemainText.Foreground = $brushDone
        $ui.RemainLabel.Text      = ' over'
    } else {
        # round up so a fresh day reads 8:00 left, not 7:59
        $ui.RemainText.Text       = Format-HM ([math]::Ceiling($left / 60) * 60)
        $ui.RemainText.Foreground = $brushRemain
        $ui.RemainLabel.Text      = ' left'
    }

    $p = if ($targetSec -gt 0) { [math]::Min(1.0, [math]::Max(0.0, $worked / $targetSec)) } else { 0 }
    $w = $ui.Track.ActualWidth
    if ($w -gt 0) { $ui.Fill.Width = $w * $p }
    $ui.Fill.Background = if ($done) { $brushDone } else { $brushAccent }

    if ($script:S.running) {
        $ui.Dot.Fill          = if ($done) { $brushDone } else { $brushAccent }
        $ui.StatusText.Text   = if ($done) { 'TARGET REACHED' } else { 'WORKING' }
        $ui.BtnToggle.Content = [char]0xE769   # pause
        $ui.MiToggle.Header   = 'Pause'
        $ui.FootRight.Text    = if ($done) { 'you can go' } else { 'done ' + (Get-Date).AddSeconds($left).ToString('HH:mm') }
    } else {
        $ui.Dot.Fill          = $brushPaused
        $ui.StatusText.Text   = 'PAUSED'
        $ui.BtnToggle.Content = [char]0xE768   # play
        $ui.MiToggle.Header   = 'Resume'
        $ui.FootRight.Text    = 'paused'
    }

    $ui.FootLeft.Text = 'of ' + (Format-HM ([double]$script:S.targetMinutes * 60))
}

# --------------------------------------------------------------------------
# behaviour
# --------------------------------------------------------------------------
$script:saveCounter = 0

$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [timespan]::FromSeconds(1)
$timer.Add_Tick({
    $now = Get-Date

    # machine was asleep / lid closed: do not count the gap as work
    $gap = ($now - $script:lastTick).TotalSeconds
    if ($gap -gt 90 -and $null -ne $script:segStart) {
        $script:segStart = $script:segStart.AddSeconds($gap)
    }
    $script:lastTick = $now

    # new day -> start over
    if ($now.ToString('yyyy-MM-dd') -ne $script:S.date) { Reset-Day; Save-WidgetState }

    Update-Ui

    $script:saveCounter++
    if ($script:saveCounter % 30 -eq 0) { Save-WidgetState }
})

$toggle = {
    if ($script:S.running) { Stop-Clock } else { Start-Clock }
    Save-WidgetState; Update-Ui
}
$UI.BtnToggle.Add_Click($toggle)
$UI.MiToggle.Add_Click($toggle)

$reset = {
    if (Show-WidgetConfirm 'Reset today''s worked time to 0:00?') { Reset-Day; Save-WidgetState; Update-Ui }
}
$UI.BtnReset.Add_Click($reset)
$UI.MiReset.Add_Click($reset)

$UI.BtnQuit.Add_Click({ Save-WidgetState; $win.Close() })

$UI.MiAdjust.Add_Click({
    $ans = Show-WidgetInput 'Worked today so far (h:mm, e.g. 3:20):' (Format-HM (Get-Worked))
    if ([string]::IsNullOrWhiteSpace($ans)) { return }
    $min = ConvertFrom-Duration $ans
    if ($null -eq $min) { Show-WidgetMessage "Could not read '$ans'. Try 3:20, 200m or 3.5."; return }
    $script:base = [double]$min * 60
    if ($null -ne $script:segStart) { $script:segStart = Get-Date }
    Save-WidgetState; Update-Ui
})

$UI.MiTarget.Add_Click({
    $ans = Show-WidgetInput 'Daily target (h:mm, e.g. 7:48):' (Format-HM ([double]$script:S.targetMinutes * 60))
    if ([string]::IsNullOrWhiteSpace($ans)) { return }
    $min = ConvertFrom-Duration $ans
    if ($null -eq $min -or $min -le 0) { Show-WidgetMessage "Could not read '$ans'. Try 7:48, 468m or 7.8."; return }
    $script:S.targetMinutes = $min
    Save-WidgetState; Update-Ui
})

Register-WidgetChrome $win
$win.Add_ContentRendered({ Update-Ui })
$win.Add_Closing({ $timer.Stop() })

$timer.Start()
Update-Ui
Write-WidgetLog 'show'
[void]$win.ShowDialog()
Write-WidgetLog 'exited'

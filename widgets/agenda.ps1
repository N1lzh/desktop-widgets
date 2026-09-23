<#
    Agenda - today's meetings from a published Outlook (or any) ICS feed.

    Same card as WorkClock, one size, always: a countdown to the next meeting
    on top and the rest of the day underneath.

    Getting the feed URL out of Outlook:
      Outlook on the web -> Settings -> Calendar -> Shared calendars ->
      "Publish a calendar", pick the calendar, "Can view all details",
      Publish, then copy the ICS link (not the HTML one).

    Run:  powershell -NoProfile -ExecutionPolicy Bypass -File widgets\agenda.ps1
    Silent launcher: Widget.vbs agenda
#>

$ErrorActionPreference = 'Stop'
$libDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'lib'
Import-Module (Join-Path $libDir 'Widget.psm1') -Force
Import-Module (Join-Path $libDir 'Ics.psm1') -Force

# the card is a fixed size - meetings come and go, the widget does not resize
$ROWS   = 5      # meeting lines that fit; anything beyond becomes "+n more"
$ROW_H  = 21
$CHROME = 127    # header + countdown + rule + footer + padding

$script:S = Initialize-Widget -Id 'agenda' -Name 'Agenda' -ScriptPath $PSCommandPath `
    -Defaults ([ordered]@{
        icsUrl         = ''
        refreshMinutes = 10
        showAllDay     = $true
    })

$script:events    = $null          # today's meetings, $null until we have any
$script:lastError = $null
$script:lastSync  = $null
$script:loading   = $false
$script:job       = $null
$script:day       = (Get-Date).Date

$CachePath  = Get-WidgetDataPath 'agenda.ics'
$ExportPath = Get-WidgetDataPath 'agenda-today.json'
$IcsModule  = Join-Path $libDir 'Ics.psm1'

# the meetings the card shows, in starting order
function Get-VisibleEvents {
    $evts = @($script:events | Where-Object { $_ })
    # out of office is a blocker, not a meeting: a daily 16:00-23:30 one would
    # own the countdown all afternoon and book 7:30h before the day has started.
    # All-day leave stays - it is worth seeing and never counted as booked.
    $evts = @($evts | Where-Object { $_.AllDay -or -not $_.Oof })
    if (-not [bool]$script:S.showAllDay) { $evts = @($evts | Where-Object { -not $_.AllDay }) }
    $evts
}

# today's meetings as JSON for other tools (e.g. a terminal greeting from WSL),
# with the same rules as the card, so nobody has to parse the ICS again
function Export-Agenda {
    try {
        $fmt  = 'yyyy-MM-ddTHH:mm:sszzz'
        $data = [ordered]@{
            date   = $script:day.ToString('yyyy-MM-dd')
            loaded = $null -ne $script:events
            synced = if ($script:lastSync) { $script:lastSync.ToString($fmt) } else { $null }
            error  = $script:lastError
            events = @(Get-VisibleEvents | ForEach-Object {
                [ordered]@{
                    start    = ([datetimeoffset]$_.Start).ToString($fmt)
                    end      = ([datetimeoffset]$_.End).ToString($fmt)
                    subject  = [string]$_.Subject
                    location = [string]$_.Location
                    allDay   = [bool]$_.AllDay
                    busy     = [bool]$_.Busy
                }
            })
        }
        # write next to it and swap, so a reader never sees half a file
        $tmp = "$ExportPath.tmp"
        ConvertTo-Json -InputObject $data -Depth 4 | Set-Content -LiteralPath $tmp -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $ExportPath -Force
    } catch {
        Write-WidgetLog "export failed: $($_.Exception.Message)"
    }
}

# --------------------------------------------------------------------------
# fetching - off the UI thread, so a slow feed never freezes the widget
# --------------------------------------------------------------------------
$script:pool = [runspacefactory]::CreateRunspacePool(1, 1)
$script:pool.Open()

$fetchWork = {
    param($url, $cachePath, $icsModule, $day)

    $out = @{ Ok = $false; Events = @(); Error = $null; FromCache = $false }
    $text = $null

    try {
        $u = $url.Trim()
        if ($u -match '^(?i)webcal://') { $u = 'https://' + $u.Substring(9) }

        if ($u -match '^(?i)https?://') {
            [Net.ServicePointManager]::SecurityProtocol =
                [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11
            $wc = New-Object Net.WebClient
            $wc.Encoding = [Text.Encoding]::UTF8
            $wc.Headers.Add('User-Agent', 'Mozilla/5.0 (Agenda widget)')
            try {
                $proxy = [Net.WebRequest]::GetSystemWebProxy()
                $proxy.Credentials = [Net.CredentialCache]::DefaultCredentials
                $wc.Proxy = $proxy
            } catch { }
            $text = $wc.DownloadString($u)
        } elseif (Test-Path -LiteralPath $u) {
            $text = Get-Content -LiteralPath $u -Raw -Encoding UTF8
        } else {
            throw "Not a URL or a file: $u"
        }

        if ($text -notmatch 'BEGIN:VCALENDAR') { throw 'that link is not an ICS calendar' }
        try {
            $dir = Split-Path $cachePath -Parent
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Set-Content -LiteralPath $cachePath -Value $text -Encoding UTF8
        } catch { }
    } catch {
        # .NET's own wording is long and localised - say the useful part instead
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        $msg = $ex.Message
        if ($ex -is [Net.WebException]) {
            if ($ex.Response) {
                $msg = 'the calendar server said ' + [int]$ex.Response.StatusCode
            } else {
                $msg = switch ([string]$ex.Status) {
                    'NameResolutionFailure' { 'no connection' }
                    'ConnectFailure'        { 'no connection' }
                    'Timeout'               { 'the server did not answer' }
                    'TrustFailure'          { 'the server certificate was refused' }
                    default                 { 'could not reach the calendar' }
                }
            }
        }
        $out.Error = $msg
        $text = $null
    }

    # no feed today? then at least show what we saw last time
    if (-not $text -and (Test-Path -LiteralPath $cachePath)) {
        try { $text = Get-Content -LiteralPath $cachePath -Raw -Encoding UTF8; $out.FromCache = $true } catch { }
    }

    if ($text) {
        try {
            Import-Module $icsModule -Force
            $out.Events = @(Get-IcsEvents -Text $text -Date $day)
            $out.Ok = $true
        } catch {
            if (-not $out.Error) { $out.Error = "Could not read the calendar: $($_.Exception.Message)" }
        }
    }
    $out
}

function Start-Fetch {
    if ($script:job) { return }
    if ([string]::IsNullOrWhiteSpace([string]$script:S.icsUrl)) { return }
    $ps = [powershell]::Create()
    $ps.RunspacePool = $script:pool
    [void]$ps.AddScript($fetchWork).AddArgument([string]$script:S.icsUrl).
        AddArgument($CachePath).AddArgument($IcsModule).AddArgument((Get-Date).Date)
    $script:job = @{ Ps = $ps; Handle = $ps.BeginInvoke() }
    $script:loading = $true
    Write-WidgetLog 'fetch started'
}

function Complete-Fetch {
    if (-not $script:job -or -not $script:job.Handle.IsCompleted) { return }
    $res = $null
    try { $res = @($script:job.Ps.EndInvoke($script:job.Handle))[0] }
    catch { $script:lastError = $_.Exception.Message }
    try { $script:job.Ps.Dispose() } catch { }
    $script:job     = $null
    $script:loading = $false

    if ($res) {
        $script:lastError = $res.Error
        if ($res.Ok) {
            $script:events   = @($res.Events)
            $script:lastSync = Get-Date
            $script:day      = (Get-Date).Date
        }
        Write-WidgetLog ("fetch done: {0} events, cache={1}, err={2}" -f @($res.Events).Count, $res.FromCache, $res.Error)
    }
    Export-Agenda
    Update-Ui
}

# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
$menu = @'
        <MenuItem x:Name="MiRefresh"  Header="Refresh now"/>
        <Separator/>
        <MenuItem x:Name="MiUrl"      Header="Calendar URL..."/>
        <MenuItem x:Name="MiInterval" Header="Refresh every..."/>
        <MenuItem x:Name="MiAllDay"   Header="Show all-day events" IsCheckable="True"/>
        <Separator/>
'@

$body = @"
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Grid Grid.Row="0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse x:Name="Dot" Width="7" Height="7" Fill="#5AA9FF" Margin="1,0,8,0"/>
          <TextBlock x:Name="StatusText" Text="TODAY" FontSize="9" FontWeight="SemiBold" Foreground="#8DA0B3"/>
        </StackPanel>
        <StackPanel x:Name="Controls" Orientation="Horizontal" HorizontalAlignment="Right" Opacity="0.18">
          <Button x:Name="BtnRefresh" Style="{StaticResource Glyph}" Content="&#xE72C;" ToolTip="Refresh now"/>
          <Button x:Name="BtnUrl"     Style="{StaticResource Glyph}" Content="&#xE713;" ToolTip="Calendar URL"/>
          <Button x:Name="BtnQuit"    Style="{StaticResource Glyph}" Content="&#xE8BB;" ToolTip="Quit"/>
        </StackPanel>
      </Grid>

      <Grid Grid.Row="1" Margin="0,1,0,8">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="HeroText" Grid.Column="0" Text="--" FontSize="34" FontWeight="Light"
                   Foreground="#F1F6FB" VerticalAlignment="Center"/>
        <StackPanel x:Name="HeroBlock" Grid.Column="1" Margin="11,0,0,0" VerticalAlignment="Center">
          <TextBlock x:Name="HeroSubject" Text="" FontSize="12.5" Foreground="#E2EAF2"
                     TextTrimming="CharacterEllipsis"/>
          <TextBlock x:Name="HeroWhen" Text="" FontSize="10.5" Foreground="#79B4FF" Margin="0,2,0,0"
                     TextTrimming="CharacterEllipsis"/>
        </StackPanel>
      </Grid>

      <Border Grid.Row="2" Height="1" Background="#20FFFFFF" Margin="0,0,0,7"/>

      <Grid Grid.Row="3" x:Name="ListHost" ClipToBounds="True">
        <StackPanel x:Name="List"/>
      </Grid>

      <Grid Grid.Row="4" Margin="0,6,0,0">
        <TextBlock x:Name="FootLeft"  Text="" FontSize="10.5" Foreground="#74838F"/>
        <TextBlock x:Name="FootRight" HorizontalAlignment="Right" Text="" FontSize="10.5" Foreground="#74838F"/>
      </Grid>
"@

$win = New-WidgetWindow -Body $body -MenuItems $menu -Height ($CHROME + $ROWS * $ROW_H)
$script:UI = Get-WidgetElements $win @(
    'Card','Dot','StatusText','Controls','BtnRefresh','BtnUrl','BtnQuit','HeroText','HeroSubject',
    'HeroWhen','HeroBlock','ListHost','List','FootLeft','FootRight','MiRefresh','MiUrl','MiInterval','MiAllDay')

$P = Get-WidgetPalette
$brushAccent = Get-WidgetBrush $P.Accent
$brushSoft   = Get-WidgetBrush $P.Soft
$brushGood   = Get-WidgetBrush $P.Good
$brushWarn   = Get-WidgetBrush $P.Warn
$brushTitle  = Get-WidgetBrush $P.Title
$brushText   = Get-WidgetBrush $P.Text
$brushFaint  = Get-WidgetBrush $P.Faint

# --------------------------------------------------------------------------
# rendering
# --------------------------------------------------------------------------
function New-Row([string]$When, [string]$What, $WhenBrush, $WhatBrush, [double]$Opacity = 1.0, [string]$Tip) {
    $g = New-Object Windows.Controls.Grid
    $g.Height  = $ROW_H
    $g.Opacity = $Opacity
    foreach ($w in @(44, 0)) {
        $c = New-Object Windows.Controls.ColumnDefinition
        $c.Width = if ($w -gt 0) { [Windows.GridLength]::new($w) } else { [Windows.GridLength]::new(1, 'Star') }
        $g.ColumnDefinitions.Add($c)
    }

    $a = New-Object Windows.Controls.TextBlock
    $a.Text = $When; $a.FontSize = 11; $a.Foreground = $WhenBrush; $a.VerticalAlignment = 'Center'
    [Windows.Controls.Grid]::SetColumn($a, 0)

    $b = New-Object Windows.Controls.TextBlock
    $b.Text = $What; $b.FontSize = 11.5; $b.Foreground = $WhatBrush; $b.VerticalAlignment = 'Center'
    $b.TextTrimming = 'CharacterEllipsis'
    [Windows.Controls.Grid]::SetColumn($b, 1)

    if ($Tip) { $g.ToolTip = $Tip }
    [void]$g.Children.Add($a)
    [void]$g.Children.Add($b)
    $g
}

function Get-MeetingTip($m) {
    $lines = @($m.Subject)
    $lines += if ($m.AllDay) { 'all day' }
              else { '{0} - {1}' -f $m.Start.ToString('HH:mm'), $m.End.ToString('HH:mm') }
    if ($m.Location) { $lines += $m.Location }
    ($lines -join "`n")
}

function Set-Hero([string]$Big, $Brush, [string]$Subject, [string]$When) {
    $ui = $script:UI
    if ($Big) {
        $ui.HeroText.Text       = $Big
        $ui.HeroText.Foreground = $Brush
        $ui.HeroText.Visibility = 'Visible'
        $ui.HeroBlock.Margin    = '11,0,0,0'
    } else {
        $ui.HeroText.Visibility = 'Collapsed'
        $ui.HeroBlock.Margin    = '0'
    }
    $ui.HeroSubject.Text = $Subject
    $ui.HeroWhen.Text    = $When
}

function Update-Ui {
    $ui   = $script:UI
    $now  = Get-Date
    $have = $null -ne $script:events
    $evts = @(Get-VisibleEvents)

    $timed   = @($evts | Where-Object { -not $_.AllDay })
    $current = @($timed | Where-Object { $_.Start -le $now -and $_.End -gt $now })[0]
    $next    = @($timed | Where-Object { $_.Start -gt $now })[0]

    # ---- the countdown ----
    # states without a number give the whole line to the text instead of putting
    # a stunted placeholder where the big digits go
    if ([string]::IsNullOrWhiteSpace([string]$script:S.icsUrl)) {
        Set-Hero '' $brushTitle 'No calendar yet' 'right-click the card > Calendar URL...'
        $ui.StatusText.Text = 'NO CALENDAR'
        $ui.Dot.Fill        = $brushFaint
    } elseif (-not $have) {
        Set-Hero '' $brushTitle `
            $(if ($script:lastError) { 'Could not load the calendar' } else { 'Loading your calendar...' }) `
            $(if ($script:lastError) { $script:lastError } else { '' })
        $ui.StatusText.Text = if ($script:lastError) { 'SYNC FAILED' } else { 'LOADING' }
        $ui.Dot.Fill        = if ($script:lastError) { $brushWarn } else { $brushFaint }
    } elseif ($current) {
        Set-Hero (Format-HM ([math]::Ceiling(($current.End - $now).TotalSeconds / 60) * 60)) $brushWarn `
            $current.Subject ('left  -  until ' + $current.End.ToString('HH:mm'))
        $ui.StatusText.Text = 'IN A MEETING'
        $ui.Dot.Fill        = $brushWarn
    } elseif ($next) {
        $mins = ($next.Start - $now).TotalMinutes
        $soon = $mins -le 5
        Set-Hero (Format-HM ([math]::Ceiling($mins) * 60)) $(if ($soon) { $brushWarn } else { $brushTitle }) `
            $next.Subject ('{0} - {1}' -f $next.Start.ToString('HH:mm'), $next.End.ToString('HH:mm'))
        $ui.StatusText.Text = 'NEXT UP'
        $ui.Dot.Fill        = if ($soon) { $brushWarn } else { $brushAccent }
    } else {
        Set-Hero 'free' $brushGood `
            $(if ($timed.Count -eq 0) { 'No meetings today' } else { 'Nothing left today' }) ''
        $ui.StatusText.Text = 'TODAY'
        $ui.Dot.Fill        = $brushGood
    }
    $ui.HeroSubject.ToolTip = if ($current) { Get-MeetingTip $current } elseif ($next) { Get-MeetingTip $next } else { $null }

    # ---- the day ----
    $cap = [int][math]::Floor($ui.ListHost.ActualHeight / $ROW_H)
    if ($cap -lt 1) { $cap = $ROWS }

    $past   = @($evts | Where-Object { -not $_.AllDay -and $_.End -le $now })
    $ahead  = @($evts | Where-Object { $_.AllDay -or $_.End -gt $now })
    $hidden = 0
    if ($ahead.Count -gt $cap) {
        $keep   = [math]::Max(1, $cap - 1)     # the last line becomes "+n more"
        $hidden = $ahead.Count - $keep
        $show   = @($ahead[0..($keep - 1)])
    } else {
        $show = $ahead
        $room = $cap - $ahead.Count
        if ($room -gt 0 -and $past.Count -gt 0) {
            $take = [math]::Min($room, $past.Count)
            $show = @($past[($past.Count - $take)..($past.Count - 1)]) + $show
        }
    }

    $ui.List.Children.Clear()
    foreach ($m in $show) {
        $isPast = (-not $m.AllDay) -and $m.End -le $now
        $isNow  = $current -and $m.Start -eq $current.Start -and $m.Subject -eq $current.Subject
        $when   = if ($m.AllDay) { 'all' } elseif ($isNow) { 'now' } else { $m.Start.ToString('HH:mm') }
        $wb     = if ($isNow) { $brushWarn } elseif ($isPast) { $brushFaint } else { $brushSoft }
        $tb     = if ($isNow) { $brushTitle } elseif ($isPast) { $brushFaint } else { $brushText }
        $op     = if ($isPast) { 0.5 } else { 1.0 }
        [void]$ui.List.Children.Add((New-Row $when $m.Subject $wb $tb $op (Get-MeetingTip $m)))
    }
    if ($hidden -gt 0) {
        [void]$ui.List.Children.Add((New-Row '' ("+$hidden more later") $brushFaint $brushFaint 1.0 $null))
    }

    # ---- the footer ----
    $n = $evts.Count
    $ui.FootLeft.Text = if (-not $have -or $n -eq 0) { '' }
                        elseif ($n -eq 1) { '1 meeting' }
                        else { "$n meetings" }

    if ($script:lastError -and $have) {
        $ui.FootRight.Text       = 'sync failed'
        $ui.FootRight.Foreground = $brushWarn
    } else {
        $busy = 0.0
        foreach ($m in $timed) { if ($m.Busy) { $busy += ($m.End - $m.Start).TotalSeconds } }
        $ui.FootRight.Text       = if ($busy -gt 0) { (Format-HM $busy) + ' booked' } else { '' }
        $ui.FootRight.Foreground = $brushFaint
    }

    $tip = @()
    if ($script:lastSync)  { $tip += 'updated ' + $script:lastSync.ToString('HH:mm') }
    if ($script:loading)   { $tip += 'refreshing...' }
    if ($script:lastError) { $tip += 'last refresh failed: ' + $script:lastError }
    $ui.Card.ToolTip = if ($tip.Count) { $tip -join "`n" } else { $null }
}

# --------------------------------------------------------------------------
# behaviour
# --------------------------------------------------------------------------
$script:ticks = 0

$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [timespan]::FromSeconds(15)
$timer.Add_Tick({
    Complete-Fetch

    # midnight, or a wake-up after the day rolled over: pull the new day
    if ((Get-Date).Date -ne $script:day) { $script:day = (Get-Date).Date; $script:events = $null; Start-Fetch }

    $script:ticks++
    $every = [math]::Max(1, [int](([int]$script:S.refreshMinutes * 60) / 15))
    if ($script:ticks % $every -eq 0) { Start-Fetch }

    Update-Ui
})

$refresh = { Start-Fetch; Update-Ui }
$UI.BtnRefresh.Add_Click($refresh)
$UI.MiRefresh.Add_Click($refresh)

$askUrl = {
    $ans = Show-WidgetInput @'
Paste the ICS link of your calendar.

Outlook on the web: Settings > Calendar > Shared calendars >
Publish a calendar > "Can view all details" > Publish,
then copy the ICS link.
'@ ([string]$script:S.icsUrl)
    if ($null -eq $ans -or $ans -eq '') { return }
    $script:S.icsUrl  = $ans.Trim()
    $script:events    = $null
    $script:lastError = $null
    Save-WidgetState
    Start-Fetch; Update-Ui
}
$UI.BtnUrl.Add_Click($askUrl)
$UI.MiUrl.Add_Click($askUrl)

$UI.MiInterval.Add_Click({
    $ans = Show-WidgetInput 'Refresh the calendar every (e.g. 10m, 1h):' ("$($script:S.refreshMinutes)m")
    if ([string]::IsNullOrWhiteSpace($ans)) { return }
    $min = ConvertFrom-Duration $ans
    if ($null -eq $min -or $min -lt 1) { Show-WidgetMessage "Could not read '$ans'. Try 10m, 30m or 1h."; return }
    $script:S.refreshMinutes = $min
    $script:ticks = 0
    Save-WidgetState
})

$UI.MiAllDay.IsChecked = [bool]$script:S.showAllDay
$UI.MiAllDay.Add_Click({
    $script:S.showAllDay = [bool]$script:UI.MiAllDay.IsChecked
    Save-WidgetState; Export-Agenda; Update-Ui
})

$UI.BtnQuit.Add_Click({ Save-WidgetState; $win.Close() })

Register-WidgetChrome $win

$win.Add_ContentRendered({
    Update-Ui
    if ([string]::IsNullOrWhiteSpace([string]$script:S.icsUrl)) {
        # first start: ask for the feed right away rather than sit there empty.
        # priority first - the other overload takes the delegate's arguments and
        # would hand 'ApplicationIdle' to a scriptblock that accepts none
        [void]$win.Dispatcher.BeginInvoke(
            [Windows.Threading.DispatcherPriority]::ApplicationIdle, [action]$askUrl)
    }
})

$win.Add_Closing({
    $timer.Stop()
    try { if ($script:job) { $script:job.Ps.Dispose() } } catch { }
    try { $script:pool.Close(); $script:pool.Dispose() } catch { }
})

Start-Fetch
$timer.Start()
Update-Ui
Write-WidgetLog 'show'
[void]$win.ShowDialog()
Write-WidgetLog 'exited'

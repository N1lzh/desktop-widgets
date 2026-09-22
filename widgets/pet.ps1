<#
    Pet - a small creature that sits on the desktop and reacts to things.

    It has no needs and it cannot die. There is nothing to feed, nothing that
    runs out, nothing that blames you for not looking at it for two days. It
    only reacts: to the machine going quiet, to a meeting coming up, to a todo
    you ticked off, to the day getting long, and to being petted.

    What it watches, all of it already on disk next to its own state file:
      state.json   the WorkClock  - paused, and how far into the day you are
      todos.json   the Todos list - how many are ticked off
      agenda.ics   the Agenda     - what starts in the next few minutes
      the machine  - how long since you last touched keyboard or mouse

    Click the face to pet it. Everywhere else on the card still drags.

    Run:  powershell -NoProfile -ExecutionPolicy Bypass -File widgets\pet.ps1
    Silent launcher: Widget.vbs pet
#>

$ErrorActionPreference = 'Stop'
$libDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'lib'
Import-Module (Join-Path $libDir 'Widget.psm1') -Force
Import-Module (Join-Path $libDir 'Ics.psm1') -Force

$TICK_MS = 200    # the heartbeat - blinking, bobbing and mood expiry run on it
$POLL_S  = 5      # how often it looks at the other widgets' state files
$ICS_MIN = 5      # how often the cached calendar is re-read
$IDLE_S  = 300    # untouched keyboard and mouse for this long and it dozes off
$EGG_S   = 40     # how long the egg wobbles on the very first start

# --------------------------------------------------------------------------
# how long since the last keystroke or mouse move - the one thing about the
# machine it needs that PowerShell cannot ask for on its own
# --------------------------------------------------------------------------
if (-not ('WC.Idle' -as [type])) {
Add-Type -Namespace WC -Name Idle -MemberDefinition @'
    [StructLayout(LayoutKind.Sequential)]
    private struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
    [DllImport("user32.dll")] private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    public static double Seconds() {
        LASTINPUTINFO lii = new LASTINPUTINFO();
        lii.cbSize = (uint)Marshal.SizeOf(lii);
        if (!GetLastInputInfo(ref lii)) return 0;
        return (double)(unchecked((uint)Environment.TickCount - lii.dwTime)) / 1000.0;
    }
'@
}

# --------------------------------------------------------------------------
# state
# --------------------------------------------------------------------------
$script:S = Initialize-Widget -Id 'pet' -Name 'Pet' -ScriptPath $PSCommandPath `
    -Defaults ([ordered]@{
        petName    = ''      # what you called it, if you called it anything
        style      = 'classic'   # which of the three creatures it is
        hatched    = ''      # when it first appeared - its age is measured from here
        pets       = 0       # how often it has been petted, ever
        treatDate  = ''      # the day it last got a treat
        lastSeen   = ''      # the last day it was running, for the morning hello
        doneCount  = -1      # ticked-off todos last time we looked, -1 = never looked
        celebrated = ''      # the day it cheered the daily target
    })

if (-not $script:S.hatched) { $script:S.hatched = (Get-Date).ToString('s'); Save-WidgetState }

$script:hatchedAt = try { [datetime]::Parse($script:S.hatched) } catch { Get-Date }
$script:wasEgg    = ((Get-Date) - $script:hatchedAt).TotalSeconds -lt $EGG_S

# --------------------------------------------------------------------------
# the faces. Plain ASCII on purpose: it is what reads as a face in a monospace
# card, and it keeps this file in the same encoding as every other one here.
# --------------------------------------------------------------------------
$P = Get-WidgetPalette

# what each mood is called and what colour the dot goes. The face itself comes
# from the chosen style below, so every style says the same things.
$MOODS = @{
    egg      = @{ Word = 'SOMETHING IN THERE'; Color = $P.Warn   }
    content  = @{ Word = 'PEACEFUL';           Color = $P.Good   }
    noticed  = @{ Word = 'HELLO';              Color = $P.Accent }
    happy    = @{ Word = 'HAPPY';              Color = $P.Good   }
    petted   = @{ Word = 'PETTED';             Color = $P.Good   }
    dizzy    = @{ Word = 'DIZZY';              Color = $P.Warn   }
    eating   = @{ Word = 'SNACKING';           Color = $P.Good   }
    sleeping = @{ Word = 'SNOOZING';           Color = $P.Muted  }
    waking   = @{ Word = 'OH, HI';             Color = $P.Accent }
    alert    = @{ Word = 'HEADS UP';           Color = $P.Warn   }
    meeting  = @{ Word = 'IN A MEETING';       Color = $P.Accent }
    tired    = @{ Word = 'TIRED';              Color = $P.Warn   }
    hungry   = @{ Word = 'PECKISH';            Color = $P.Warn   }
    wheee    = @{ Word = 'WHEEE';              Color = $P.Accent }
    curious  = @{ Word = 'CURIOUS';            Color = $P.Accent }
    sneeze   = @{ Word = 'ACHOO';              Color = $P.Warn   }
    stretch  = @{ Word = 'STRETCHING';         Color = $P.Good   }
    bored    = @{ Word = 'BORED';              Color = $P.Muted  }
    singing  = @{ Word = 'SINGING';            Color = $P.Good   }
    worried  = @{ Word = 'WORRIED';            Color = $P.Warn   }
}

# Three creatures to pick from in the menu, @(eyes open, eyes shut) per mood.
# Plain ASCII on purpose: it is what reads as a face in a monospace card, and it
# keeps this file in the same encoding as every other one here.
$FACES = @{
    classic = @{
        egg      = @('(  o  )',       '(  o  )')
        content  = @('( ^ . ^ )',     '( - . - )')
        noticed  = @('( o . o )',     '( - . - )')
        happy    = @('\ ( ^ o ^ ) /', '\ ( - o - ) /')
        petted   = @('( ^ 3 ^ )',     '( ^ 3 ^ )')
        dizzy    = @('( @ _ @ )',     '( x _ x )')
        eating   = @('( ^ 0 ^ )',     '( ^ o ^ )')
        sleeping = @('( - . - )',     '( - o - )')
        waking   = @('( o . o )',     '( - . - )')
        alert    = @('( O _ O )',     '( - _ - )')
        meeting  = @('( o _ o )',     '( - _ - )')
        tired    = @('( u _ u )',     '( u _ u )')
        hungry   = @('( > _ < )',     '( - _ - )')
        wheee    = @('\ ( o o ) /',   '\ ( o o ) /')
        curious  = @('( o . O )',     '( - . - )')
        sneeze   = @('( > o < )',     '( - o - )')
        stretch  = @('( ~ o ~ )',     '( ~ - ~ )')
        bored    = @('( . _ . )',     '( - _ - )')
        singing  = @('( ^ o ^ ) ~',   '( ^ - ^ ) ~')
        worried  = @('( o _ o ) ;',   '( - _ - ) ;')
    }
    cat = @{
        egg      = @('(  o  )',        '(  .  )')
        content  = @('( ^ w ^ )~',     '( - w - )~')
        noticed  = @('( o w o )~',     '( - w - )~')
        happy    = @('\ ( ^ w ^ ) /',  '\ ( - w - ) /')
        petted   = @('( * w * )~',     '( ^ w ^ )~')
        dizzy    = @('( @ w @ )',      '( x w x )')
        eating   = @('( o w o ) ~',    '( - w - ) ~')
        sleeping = @('( - w - ) z',    '( - w - ) Z')
        waking   = @('( o w - )~',     '( o w o )~')
        alert    = @('( O w O )',      '( - w - )')
        meeting  = @('( = w = )~',     '( - w - )~')
        tired    = @('( u w u )~',     '( u w u )~')
        hungry   = @('( > w < )',      '( - w - )')
        wheee    = @('\ ( o w o ) /',  '\ ( o w o ) /')
        curious  = @('( o w O )~',     '( - w - )~')
        sneeze   = @('( > w < ) !',    '( - w - ) !')
        stretch  = @('( ~ w ~ )~',     '( ~ w ~ )~')
        bored    = @('( . w . )~',     '( - w - )~')
        singing  = @('( ^ w ^ ) ~~',   '( - w - ) ~~')
        worried  = @('( ; w ; )',      '( - w - )')
    }
    robot = @{
        egg      = @('[  .  ]',       '[  o  ]')
        content  = @('[ o _ o ]',     '[ - _ - ]')
        noticed  = @('[ O _ o ]',     '[ - _ - ]')
        happy    = @('\ [ ^ _ ^ ] /', '\ [ - _ - ] /')
        petted   = @('[ * _ * ]',     '[ * _ * ]')
        dizzy    = @('[ x _ x ]',     '[ @ _ @ ]')
        eating   = @('[ o _ 0 ]',     '[ - _ 0 ]')
        sleeping = @('[ _ _ _ ]',     '[ . _ . ]')
        waking   = @('[ O _ O ]',     '[ - _ - ]')
        alert    = @('[ ! _ ! ]',     '[ - _ - ]')
        meeting  = @('[ = _ = ]',     '[ - _ - ]')
        tired    = @('[ u _ u ]',     '[ u _ u ]')
        hungry   = @('[ > _ < ]',     '[ - _ - ]')
        wheee    = @('\ [ o _ o ] /', '\ [ o _ o ] /')
        curious  = @('[ ? _ ? ]',     '[ - _ - ]')
        sneeze   = @('[ x _ o ]',     '[ o _ x ]')
        stretch  = @('[ ~ _ ~ ]',     '[ - _ - ]')
        bored    = @('[ . _ . ]',     '[ - _ - ]')
        singing  = @('[ o _ o ] ~',   '[ - _ - ] ~')
        worried  = @('[ ! _ o ]',     '[ - _ - ]')
    }
}

# Everything it says is drawn from a pool, so it rarely says the same thing
# twice in a row. What it says to itself when nothing at all is going on:
$CHATTER = @(
    '...', '*hums*', 'nice desktop', 'watching the cursor', 'still here',
    'the fan is loud today', 'thinking about nothing', 'this is a good corner',
    'zzz? no. not yet.', '*blinks slowly*', 'the cursor went that way',
    'i like it here', 'what were we doing again?', 'nothing to report',
    '*looks out of the window*', 'the wallpaper is nice today')

# the moods that last a while pick a line here and change it every so often
$AMB_LINES = @{
    sleeping = @('z z z', 'z Z z', '*dreaming*', '*snores quietly*', 'out cold')
    tired    = @('long one today', 'that is a lot of hours', '*yawns*',
                 'the day is winning', 'we have done enough')
    hungry   = @('is it lunch yet?', 'lunch? lunch.', 'i smell something',
                 'what is in the fridge?', '*stomach noises*')
    bored    = @('*taps foot*', 'nothing is happening', '*counts pixels*')
}

# and every few minutes, when nothing else is going on, it does one of these
# by itself - the Hop ones jump on the spot
$ANTICS = @(
    @{ Mood = 'stretch';  Line = '*stretches*' }
    @{ Mood = 'stretch';  Line = '*rolls over*' }
    @{ Mood = 'sneeze';   Line = '*achoo*';                Hop = $true }
    @{ Mood = 'curious';  Line = 'what was that?' }
    @{ Mood = 'curious';  Line = '*follows the cursor*' }
    @{ Mood = 'curious';  Line = '*sniffs the taskbar*' }
    @{ Mood = 'bored';    Line = 'is there anything to do?' }
    @{ Mood = 'bored';    Line = '*stares into the middle distance*' }
    @{ Mood = 'singing';  Line = '*la la la*' }
    @{ Mood = 'singing';  Line = '*whistles*' }
    @{ Mood = 'happy';    Line = 'found a crumb';          Hop = $true }
    @{ Mood = 'happy';    Line = '*does a little dance*';  Hop = $true }
    @{ Mood = 'dizzy';    Line = '*chases its tail*';      Hop = $true }
    @{ Mood = 'sleeping'; Line = '*micro nap*' }
    @{ Mood = 'noticed';  Line = 'oh, you are still here' }
)

$PET_LINES     = @('*purrs*', 'again! again!', 'that is the spot', '<3', '*wiggles*',
                   '*leans in*', 'more of that', '*happy noises*')
$TREAT_LINES   = @('nom nom nom', 'the good stuff', '*crunch*', 'you remembered',
                   '*inhales it*', 'best day')
$WAKE_LINES    = @('oh, you are back', 'where have you been?', '*wakes up*',
                   'i kept your seat warm', 'you missed nothing')
$MORNING_LINES = @('morning!', 'a new day', 'good morning', 'here we go again')
$DONE_LINES    = @('one down', 'nice one', '*claps*', 'that is one less', 'good job')
$CLEAR_LINES   = @('the list is empty!', 'all done!', 'nothing left!')
$TARGET_LINES  = @('target reached - go home', 'that is a full day',
                   'you can stop now', 'done is done')
$NEWTODO_LINES = @('ooh, a new one', 'one more for the pile',
                   '*reads over your shoulder*')
$OVER_LINES    = @('that is over', 'free again', '*exhales*')
$GAP_LINES     = @('did i miss something?', 'how long was i out?', 'what happened?')

# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
$menu = @'
        <MenuItem x:Name="MiTreat" Header="Give a treat"/>
        <MenuItem x:Name="MiName"  Header="Name it..."/>
        <MenuItem Header="Style">
          <MenuItem x:Name="MiStyleClassic" Header="Classic&#9;( ^ . ^ )"  IsCheckable="True"/>
          <MenuItem x:Name="MiStyleCat"     Header="Cat&#9;( ^ w ^ )~"     IsCheckable="True"/>
          <MenuItem x:Name="MiStyleRobot"   Header="Robot&#9;[ o _ o ]"    IsCheckable="True"/>
        </MenuItem>
        <Separator/>
        <MenuItem x:Name="MiNew"   Header="Start over as an egg..."/>
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
          <Ellipse x:Name="Dot" Width="7" Height="7" Fill="#48D28A" Margin="1,0,8,0"/>
          <TextBlock x:Name="StatusText" Text="PEACEFUL" FontSize="9" FontWeight="SemiBold" Foreground="#8DA0B3"/>
        </StackPanel>
        <StackPanel x:Name="Controls" Orientation="Horizontal" HorizontalAlignment="Right" Opacity="0.18">
          <Button x:Name="BtnTreat" Style="{StaticResource Glyph}" Content="&#xE735;" ToolTip="Give a treat"/>
          <Button x:Name="BtnQuit"  Style="{StaticResource Glyph}" Content="&#xE8BB;" ToolTip="Quit"/>
        </StackPanel>
      </Grid>

      <Grid Grid.Row="1" Height="46">
        <TextBlock x:Name="Face" Text="( ^ . ^ )" FontFamily="Consolas" FontSize="25"
                   Foreground="#F1F6FB" HorizontalAlignment="Center" VerticalAlignment="Center"
                   Cursor="Hand" ToolTip="pet me">
          <TextBlock.RenderTransform>
            <TranslateTransform x:Name="FaceMove"/>
          </TextBlock.RenderTransform>
        </TextBlock>
      </Grid>

      <TextBlock Grid.Row="2" x:Name="LineText" Text="" FontSize="11" Foreground="#8DA0B3"
                 HorizontalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="0,0,0,2"/>

      <Grid Grid.Row="3" Margin="0,8,0,0">
        <TextBlock x:Name="FootLeft"  Text="" FontSize="10.5" Foreground="#74838F"/>
        <TextBlock x:Name="FootRight" HorizontalAlignment="Right" Text="" FontSize="10.5" Foreground="#74838F"/>
      </Grid>
'@

$win = New-WidgetWindow -Body $body -MenuItems $menu
$script:UI = Get-WidgetElements $win @(
    'Dot','StatusText','Controls','BtnTreat','BtnQuit','Face','FaceMove','LineText',
    'FootLeft','FootRight','MiTreat','MiName','MiNew',
    'MiStyleClassic','MiStyleCat','MiStyleRobot')

# --------------------------------------------------------------------------
# moods. Two layers: whatever is going on in general (ambient), and a short
# reaction on top of it (temporary) that expires by itself.
# --------------------------------------------------------------------------
$script:ambMood   = 'content'
$script:ambLine   = ''
$script:tmpMood   = $null
$script:tmpLine   = ''
$script:tmpUntil  = [datetime]::MinValue
$script:lastMood  = ''

function Set-TempMood([string]$Name, [string]$Line, [double]$Seconds) {
    $script:tmpMood  = $Name
    $script:tmpLine  = $Line
    $script:tmpUntil = (Get-Date).AddSeconds($Seconds)
}

function Get-CurrentMood {
    if ($script:tmpMood -and (Get-Date) -lt $script:tmpUntil) { return $script:tmpMood }
    $script:tmpMood = $null
    $script:ambMood
}

function Get-CurrentLine {
    if ($script:tmpMood -and (Get-Date) -lt $script:tmpUntil) { return $script:tmpLine }
    $script:ambLine
}

# --------------------------------------------------------------------------
# reading the other widgets - they all keep their state next to ours
# --------------------------------------------------------------------------
function Read-Json([string]$Path) {
    try {
        if (Test-Path -LiteralPath $Path) {
            return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json)
        }
    } catch { }   # half-written file: we simply look again in five seconds
    $null
}

$script:icsPath    = Get-WidgetDataPath 'agenda.ics'
$script:icsRead    = [datetime]::MinValue
$script:icsStamp   = [datetime]::MinValue
$script:meetings   = @()

function Update-Meetings {
    $now = Get-Date
    if (($now - $script:icsRead).TotalMinutes -lt $ICS_MIN) { return }
    $script:icsRead = $now
    try {
        if (-not (Test-Path -LiteralPath $script:icsPath)) { $script:meetings = @(); return }
        # the Agenda widget refreshes the cache every ten minutes; no point
        # parsing the same text again
        $stamp = (Get-Item -LiteralPath $script:icsPath).LastWriteTime
        if ($stamp -eq $script:icsStamp -and $script:meetings.Count) { return }
        $script:icsStamp = $stamp
        $text = Get-Content -LiteralPath $script:icsPath -Raw -Encoding UTF8
        # same rule the Agenda card uses: out-of-office is a blocker, not a meeting
        $script:meetings = @(Get-IcsEvents -Text $text -Date $now |
            Where-Object { -not $_.AllDay -and -not $_.Oof })
    } catch { $script:meetings = @() }
}

$script:wasIdle     = $false
$script:firstPoll   = $true
$script:wheeeUntil  = [datetime]::MinValue
$script:hopUntil    = [datetime]::MinValue
$script:lastItems   = -1
$script:wasMeeting  = $false
$script:lastDay     = ''
$script:ambLineAt   = [datetime]::MinValue

# --------------------------------------------------------------------------
# the battery, if there is one. Looked at once a minute; a machine without a
# battery is asked exactly once and then left alone.
# --------------------------------------------------------------------------
$script:hasBattery = $null
$script:batAt      = [datetime]::MinValue
$script:onBattery  = $null
$script:batPct     = 100

function Update-Battery {
    if ($script:hasBattery -eq $false) { return }
    $now = Get-Date
    if (($now - $script:batAt).TotalSeconds -lt 60) { return }
    $script:batAt = $now
    try {
        $b = @(Get-CimInstance -ClassName Win32_Battery -ErrorAction Stop)[0]
        if (-not $b) { $script:hasBattery = $false; return }
        $script:hasBattery = $true
        $was = $script:onBattery
        $script:onBattery = ([int]$b.BatteryStatus -eq 1)    # 1 = on battery, 2 = on AC
        $script:batPct    = [int]$b.EstimatedChargeRemaining
        if ($null -ne $was -and $was -ne $script:onBattery) {
            if ($script:onBattery) { Set-TempMood 'noticed' 'unplugged - we are on battery' 8 }
            else { Set-TempMood 'happy' (@('ah, power', 'plugged in again') | Get-Random) 8 }
        }
    } catch { $script:hasBattery = $false }
}

function Set-Ambient([string]$Mood, [string]$Line) {
    $changed = $script:ambMood -ne $Mood
    $script:ambMood = $Mood
    if ($Line) { $script:ambLine = $Line; return }
    $pool = $AMB_LINES[$Mood]
    # nothing in particular to say - drop the last mood's line and let the
    # idle chatter fill the silence
    if (-not $pool) { if ($changed) { $script:ambLine = '' }; return }
    if ($changed -or (Get-Date) -gt $script:ambLineAt) {
        $script:ambLine   = $pool | Get-Random
        $script:ambLineAt = (Get-Date).AddSeconds((Get-Random -Minimum 60 -Maximum 240))
    }
}

function Update-Signals {
    $now   = Get-Date
    $today = $now.ToString('yyyy-MM-dd')
    $idle  = [WC.Idle]::Seconds()

    Update-Meetings
    Update-Battery

    $soon = $null; $current = $null
    foreach ($m in $script:meetings) {
        if ($m.Start -le $now -and $m.End -gt $now) { $current = $m }
        elseif ($m.Start -gt $now -and -not $soon)  { $soon = $m }
    }

    # --- what just changed: these fire a short reaction ---------------------
    $todos = Read-Json (Get-WidgetDataPath 'todos.json')
    if ($todos) {
        $items = @(@($todos.items) | Where-Object { $_ })
        $done  = @($items | Where-Object { $_.done }).Count
        $open  = $items.Count - $done
        if (-not $script:firstPoll) {
            if ([int]$script:S.doneCount -ge 0 -and $done -gt [int]$script:S.doneCount) {
                $line = if ($open -eq 0) { $CLEAR_LINES | Get-Random }
                        else { '{0}, {1} to go' -f ($DONE_LINES | Get-Random), $open }
                Set-TempMood 'happy' $line 10
            } elseif ($script:lastItems -ge 0 -and $items.Count -gt $script:lastItems) {
                Set-TempMood 'curious' ($NEWTODO_LINES | Get-Random) 8
            }
        }
        $script:S.doneCount = $done
        $script:lastItems   = $items.Count
    }

    $clock = Read-Json (Get-WidgetDataPath 'state.json')
    $paused = $false; $overtime = $false
    if ($clock -and $clock.date -eq $today) {
        $paused   = -not [bool]$clock.running
        $worked   = [double]$clock.workedSeconds
        $target   = [double]$clock.targetMinutes * 60
        $overtime = $target -gt 0 -and $worked -gt $target + 1800
        if ($target -gt 0 -and $worked -ge $target -and $script:S.celebrated -ne $today) {
            $script:S.celebrated = $today
            if (-not $script:firstPoll) { Set-TempMood 'happy' ($TARGET_LINES | Get-Random) 15 }
            Save-WidgetState
        }
    }

    # the meeting you were in has just ended
    if ($script:wasMeeting -and -not $current -and -not $script:firstPoll) {
        Set-TempMood 'happy' ($OVER_LINES | Get-Random) 8
    }
    $script:wasMeeting = [bool]$current

    # you came back to the machine
    if ($script:wasIdle -and $idle -lt 5 -and -not $script:firstPoll) {
        Set-TempMood 'waking' ($WAKE_LINES | Get-Random) 8
    }
    $script:wasIdle = $idle -ge $IDLE_S

    # midnight passed while it was running - it is a day older
    if ($script:lastDay -and $script:lastDay -ne $today -and -not $script:firstPoll) {
        Set-TempMood 'stretch' (@('a new day', 'midnight already', '*stretches*') | Get-Random) 12
    }
    $script:lastDay = $today

    # --- and what is going on in general, most pressing first ---------------
    $hour = $now.Hour
    if ($current) {
        Set-Ambient 'meeting' ('shh... ' + $current.Subject)
    } elseif ($soon -and ($soon.Start - $now).TotalMinutes -le 5) {
        $mins = [math]::Max(1, [int][math]::Ceiling(($soon.Start - $now).TotalMinutes))
        Set-Ambient 'alert' ('{0} in {1} min' -f $soon.Subject, $mins)
    } elseif ($script:hasBattery -and $script:onBattery -and $script:batPct -le 20) {
        Set-Ambient 'worried' ('battery at {0}%' -f $script:batPct)
    } elseif ($idle -ge $IDLE_S -or $hour -ge 16 -or $hour -lt 6) {
        Set-Ambient 'sleeping' $null
    } elseif ($paused) {
        Set-Ambient 'sleeping' 'break time'
    } elseif ($overtime) {
        Set-Ambient 'tired' $null
    } elseif ($now.TimeOfDay -gt ([timespan]'11:45') -and $now.TimeOfDay -lt ([timespan]'13:15') -and
              $script:S.treatDate -ne $today) {
        Set-Ambient 'hungry' $null
    } else {
        Set-Ambient 'content' $null
    }

    $script:firstPoll = $false
}

# --------------------------------------------------------------------------
# drawing
# --------------------------------------------------------------------------
$script:brushes = @{}
function Get-Brush([string]$Color) {
    if (-not $script:brushes.ContainsKey($Color)) { $script:brushes[$Color] = Get-WidgetBrush $Color }
    $script:brushes[$Color]
}

$HEART = [char]0x2665
$HOLLOW = [char]0x2661

function Get-Bond {
    $days = [math]::Floor(((Get-Date) - $script:hatchedAt).TotalDays)
    [math]::Min(5, 1 + [math]::Floor([int]$script:S.pets / 10) + [math]::Floor($days / 7))
}

function Update-Footer {
    $days = [math]::Floor(((Get-Date) - $script:hatchedAt).TotalDays) + 1
    $age  = if ($days -le 1) { 'day one' } else { "day $days" }
    $script:UI.FootLeft.Text = if ($script:S.petName) { '{0} - {1}' -f $script:S.petName, $age } else { $age }

    $bond = Get-Bond
    $script:UI.FootRight.Text = ($HEART.ToString() * $bond) + ($HOLLOW.ToString() * (5 - $bond))
}

$script:blink    = $false
$script:reblink  = $false
$script:nextBlink = 12
$script:chatterAt = (Get-Date).AddMinutes(2)

function Get-Face([string]$Mood) {
    $set = $FACES[[string]$script:S.style]
    if (-not $set) { $set = $FACES['classic'] }
    $pair = $set[$Mood]
    if (-not $pair) { $pair = $FACES['classic'][$Mood] }
    if (-not $pair) { $pair = $FACES['classic']['content'] }
    $pair
}

function Update-Face {
    $name = Get-CurrentMood
    $m    = $MOODS[$name]
    if (-not $m) { $m = $MOODS['content']; $name = 'content' }
    $ui   = $script:UI
    $pair = Get-Face $name

    $ui.Face.Text       = if ($script:blink) { $pair[1] } else { $pair[0] }
    $ui.StatusText.Text = $m.Word
    $ui.Dot.Fill        = Get-Brush $m.Color
    $ui.LineText.Text   = Get-CurrentLine
    $script:lastMood    = $name
}

# --------------------------------------------------------------------------
# being petted
# --------------------------------------------------------------------------
$script:petBurst = 0
$script:petAt    = [datetime]::MinValue

function Test-Egg {
    ((Get-Date) - $script:hatchedAt).TotalSeconds -lt $EGG_S
}

function Invoke-Pet {
    $now = Get-Date
    # hammering on it makes it giddy, then dizzy - it recovers on its own
    if (($now - $script:petAt).TotalSeconds -lt 2) { $script:petBurst++ } else { $script:petBurst = 1 }
    $script:petAt = $now
    $script:S.pets = [int]$script:S.pets + 1

    if (Test-Egg) {
        # there is nothing to pet yet - it just rocks a little
        Set-TempMood 'egg' '*wobble*' 3
    } elseif ($script:ambMood -eq 'sleeping' -and $script:petBurst -eq 1) {
        Set-TempMood 'waking' 'mmh... hello' 6
    } elseif ($script:petBurst -ge 5) {
        Set-TempMood 'dizzy' 'too much! too much!' 6
    } else {
        Set-TempMood 'petted' ($PET_LINES | Get-Random) 4
        if ((Get-Random -Minimum 0 -Maximum 3) -eq 0) { $script:hopUntil = $now.AddMilliseconds(700) }
    }
    Save-WidgetState
    Update-Face
    Update-Footer
}

function Add-Treat {
    if (Test-Egg) { Set-TempMood 'egg' 'eggs do not eat' 4; Update-Face; return }
    $script:S.treatDate = (Get-Date).ToString('yyyy-MM-dd')
    Set-TempMood 'eating' ($TREAT_LINES | Get-Random) 6
    Save-WidgetState
    Update-Face
}

$UI.Face.Add_MouseLeftButtonDown({
    # the face pets, the rest of the card still drags
    $args[1].Handled = $true
    Invoke-Pet
})

$UI.BtnTreat.Add_Click({ Add-Treat })
$UI.MiTreat.Add_Click({ Add-Treat })
$UI.BtnQuit.Add_Click({ Save-WidgetState; $win.Close() })

$UI.MiName.Add_Click({
    $ans = Show-WidgetInput 'What is it called?' $script:S.petName
    if ($null -eq $ans) { return }
    $script:S.petName = $ans.Trim()
    Save-WidgetState
    if ($script:S.petName) { Set-TempMood 'happy' ('{0}. i like it.' -f $script:S.petName) 8 }
    Update-Face; Update-Footer
})

# --------------------------------------------------------------------------
# which creature it is. Switching is live - the moods and everything it
# reacts to stay exactly the same, only the face changes.
# --------------------------------------------------------------------------
$script:styleItems = @{
    classic = $UI.MiStyleClassic
    cat     = $UI.MiStyleCat
    robot   = $UI.MiStyleRobot
}

function Set-Style([string]$Name) {
    if (-not $FACES.ContainsKey($Name)) { $Name = 'classic' }
    $script:S.style = $Name
    foreach ($k in $script:styleItems.Keys) { $script:styleItems[$k].IsChecked = ($k -eq $Name) }
    Save-WidgetState
    Update-Face
}

$UI.MiStyleClassic.Add_Click({ Set-TempMood 'noticed' 'how do i look?' 5; Set-Style 'classic' })
$UI.MiStyleCat.Add_Click({     Set-TempMood 'noticed' 'how do i look?' 5; Set-Style 'cat' })
$UI.MiStyleRobot.Add_Click({   Set-TempMood 'noticed' 'how do i look?' 5; Set-Style 'robot' })

$UI.MiNew.Add_Click({
    if (-not (Show-WidgetConfirm 'Start over? It goes back to being an egg and forgets its name.')) { return }
    $script:hatchedAt   = Get-Date
    $script:S.hatched   = $script:hatchedAt.ToString('s')
    $script:S.petName   = ''
    $script:S.pets      = 0
    $script:wasEgg      = $true
    Save-WidgetState
    Update-Face; Update-Footer
})

# dragged across the desktop. Not while the card is still being put back where
# it was last time - that is a move too, and it is not a ride.
$script:placed = $false
$win.Add_LocationChanged({ if ($script:placed) { $script:wheeeUntil = (Get-Date).AddSeconds(1.5) } })

# noticed you reaching for it
$UI.Face.Add_MouseEnter({
    if ($script:ambMood -eq 'content' -and -not $script:tmpMood) { Set-TempMood 'noticed' '' 2.5 }
})

# --------------------------------------------------------------------------
# the heartbeat
# --------------------------------------------------------------------------
$script:ticks    = 0
$script:saveAt   = (Get-Date).AddMinutes(1)
$script:lastTick = Get-Date
$script:anticAt  = (Get-Date).AddMinutes((Get-Random -Minimum 2 -Maximum 6))
$HOP_MS          = 700

$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [timespan]::FromMilliseconds($TICK_MS)
$timer.Add_Tick({
    $script:ticks++
    $now = Get-Date

    # --- still an egg? -----------------------------------------------------
    $eggAge = ($now - $script:hatchedAt).TotalSeconds
    if ($eggAge -lt $EGG_S) {
        $script:ambMood = 'egg'
        $script:ambLine = if ($eggAge -gt $EGG_S - 6) { '*crack*' }
                          elseif ($eggAge -gt $EGG_S / 2) { 'it moved!' }
                          else { 'something is in there...' }
        # the egg wobbles instead of breathing, harder as it gets closer
        $amp = 1.0 + 2.0 * ($eggAge / $EGG_S)
        $script:UI.FaceMove.X = [math]::Sin($script:ticks * 0.6) * $amp
        $script:UI.FaceMove.Y = 0
        Update-Face
        return
    }
    if ($script:wasEgg) {
        $script:wasEgg = $false
        $script:UI.FaceMove.X = 0
        Set-TempMood 'happy' 'hi! i live here now' 12
    }

    # --- the machine was away: asleep, or the lid was shut -----------------
    if (($now - $script:lastTick).TotalSeconds -gt 120) {
        Set-TempMood 'waking' ($GAP_LINES | Get-Random) 10
    }
    $script:lastTick = $now

    # --- a hop, or otherwise slow breathing --------------------------------
    if ($now -lt $script:hopUntil) {
        $left = ($script:hopUntil - $now).TotalMilliseconds
        $script:UI.FaceMove.Y = -[math]::Abs([math]::Sin((1 - $left / $HOP_MS) * [math]::PI * 2)) * 7
    } else {
        $script:UI.FaceMove.Y = [math]::Sin($script:ticks * 0.11) * 1.6
    }

    # --- blinking ----------------------------------------------------------
    if ($script:blink) {
        $script:blink = $false
        # now and then it blinks twice in a row
        if ($script:reblink) { $script:reblink = $false; $script:nextBlink = 2 }
        Update-Face
    } elseif (--$script:nextBlink -le 0) {
        $script:nextBlink = Get-Random -Minimum 12 -Maximum 45   # every 2.5 - 9 s
        $script:reblink   = (Get-Random -Minimum 0 -Maximum 4) -eq 0
        $script:blink     = $true
        Update-Face
    }

    # --- being thrown around the desktop -----------------------------------
    if ($now -lt $script:wheeeUntil) {
        if ($script:lastMood -ne 'wheee') { Set-TempMood 'wheee' 'wheee' 1.5; Update-Face }
    }

    # --- the slow loop: what is going on out there -------------------------
    if ($script:ticks % ([int](1000 / $TICK_MS) * $POLL_S) -eq 0) {
        Update-Signals
        Update-Footer
        Update-Face
    }

    # --- and every few minutes it just does something ----------------------
    if ($now -gt $script:anticAt) {
        $script:anticAt = $now.AddMinutes((Get-Random -Minimum 4 -Maximum 13))
        if ($script:ambMood -eq 'content' -and -not $script:tmpMood) {
            $a = $ANTICS | Get-Random
            Set-TempMood $a.Mood $a.Line (Get-Random -Minimum 5 -Maximum 10)
            if ($a.Hop) { $script:hopUntil = $now.AddMilliseconds($HOP_MS) }
            Update-Face
        }
    }

    # --- talking to itself when nothing at all is happening ----------------
    if ($now -gt $script:chatterAt) {
        $script:chatterAt = $now.AddMinutes((Get-Random -Minimum 2 -Maximum 6))
        if ($script:ambMood -eq 'content' -and -not $script:tmpMood) {
            $script:ambLine = $CHATTER | Get-Random
            Update-Face
        }
    }

    # the mood may simply have run out
    if ($script:lastMood -ne (Get-CurrentMood)) { Update-Face }

    if ($now -gt $script:saveAt) { $script:saveAt = $now.AddMinutes(5); Save-WidgetState }
})

Register-WidgetChrome $win

$win.Add_ContentRendered({
    $script:placed = $true     # the chrome has put the card back by now
    Set-Style ([string]$script:S.style)
    Update-Signals
    Update-Footer

    $today = (Get-Date).ToString('yyyy-MM-dd')
    if ($script:S.lastSeen -and $script:S.lastSeen -ne $today -and -not $script:wasEgg) {
        Set-TempMood 'happy' 'morning!' 12
    }
    $script:S.lastSeen = $today
    Save-WidgetState
    Update-Face
})
$win.Add_Closing({ $timer.Stop() })

$timer.Start()
Update-Face
Write-WidgetLog 'show'
[void]$win.ShowDialog()
Write-WidgetLog 'exited'

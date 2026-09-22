<#
    Widget - the shared plumbing behind every widget in this folder.

    A widget is a single .ps1 file under widgets\ that

        Import-Module  lib\Widget.psm1
        Initialize-Widget  -Id/-Name/-StateFile/-Defaults
        New-WidgetWindow   -Body <xaml> -MenuItems <xaml>
        Register-WidgetChrome $win
        ...draws its own content...
        [void]$win.ShowDialog()

    Everything the widgets have in common lives here: the frameless translucent
    card, the tint palettes and the scroll-to-veil, blur / Windows-acrylic, the
    drag-anywhere behaviour, the remembered position, the Background and
    "Start with Windows" menu entries and the little JSON state file.
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, Microsoft.VisualBasic

# --------------------------------------------------------------------------
# native bits: blur behind, acrylic backdrop, rounded corners, tool-window style
# --------------------------------------------------------------------------
if (-not ('WC.Native' -as [type])) {
Add-Type -Namespace WC -Name Native -MemberDefinition @'
    [StructLayout(LayoutKind.Sequential)]
    private struct MARGINS { public int cxLeft, cxRight, cyTop, cyBottom; }

    [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")] private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr hWnd, int attr, ref int val, int size);
    [DllImport("dwmapi.dll")] private static extern int DwmExtendFrameIntoClientArea(IntPtr hWnd, ref MARGINS m);
    [DllImport("user32.dll")] private static extern int SetWindowCompositionAttribute(IntPtr hWnd, ref WinCompAttrData data);
    [StructLayout(LayoutKind.Sequential)]
    private struct AccentPolicy { public int AccentState; public int AccentFlags; public int GradientColor; public int AnimationId; }
    [StructLayout(LayoutKind.Sequential)]
    private struct WinCompAttrData { public int Attribute; public IntPtr Data; public int SizeOfData; }

    // Win11 acrylic backdrop + rounded corners, drawn by the compositor itself.
    public static void ApplyAcrylic(IntPtr hWnd) {
        int dark = 1;     DwmSetWindowAttribute(hWnd, 20, ref dark, 4);      // immersive dark mode
        int round = 2;    DwmSetWindowAttribute(hWnd, 33, ref round, 4);     // small rounded corners
        int backdrop = 3; DwmSetWindowAttribute(hWnd, 38, ref backdrop, 4);  // acrylic
        MARGINS m = new MARGINS();
        m.cxLeft = -1; m.cxRight = -1; m.cyTop = -1; m.cyBottom = -1;
        DwmExtendFrameIntoClientArea(hWnd, ref m);
    }

    // Blur whatever sits behind the window. state 3 = blur behind, 4 = acrylic
    // blur behind; the tint itself is painted by WPF on top of it.
    public static void ApplyBlur(IntPtr hWnd, int state, int gradientAbgr) {
        AccentPolicy p = new AccentPolicy();
        p.AccentState   = state;
        p.AccentFlags   = 2;
        p.GradientColor = gradientAbgr;
        int size = Marshal.SizeOf(p);
        IntPtr buf = Marshal.AllocHGlobal(size);
        try {
            Marshal.StructureToPtr(p, buf, false);
            WinCompAttrData d = new WinCompAttrData();
            d.Attribute = 19;   // WCA_ACCENT_POLICY
            d.Data = buf;
            d.SizeOfData = size;
            SetWindowCompositionAttribute(hWnd, ref d);
        } finally { Marshal.FreeHGlobal(buf); }
    }

    // The blur fills the whole window rect, so let the DWM round it off -
    // SetWindowRgn is ignored on layered windows, the corner preference is not.
    public static void RoundCorners(IntPtr hWnd) {
        int round = 2;   // DWMWA_WINDOW_CORNER_PREFERENCE = round
        DwmSetWindowAttribute(hWnd, 33, ref round, 4);
    }

    // keep the widget out of alt-tab
    public static void MakeToolWindow(IntPtr hWnd) {
        int ex = GetWindowLong(hWnd, -20);
        SetWindowLong(hWnd, -20, ex | 0x00000080);
    }
'@
}

# --------------------------------------------------------------------------
# looks - shared by every widget so they sit next to each other as a set
# --------------------------------------------------------------------------
# How much the card veils what is behind it: 12% .. 85%. At the low end the
# desktop reads straight through the widget, at the high end it is a solid card.
# Scroll on the widget or use the Background menu to walk through them.
$script:TINTS_GLASS   = @('#1F151B26', '#33151B26', '#4D141A24', '#70131822', '#A312171F', '#D611151C')
$script:TINTS_FROSTED = @('#33FFFFFF', '#1AFFFFFF', '#00FFFFFF', '#26000000', '#4D000000', '#7A000000')
$script:CORNER        = 8     # matches the radius the DWM rounds the window with
$script:WIDGET_WIDTH  = 258   # every widget is exactly this wide

# the palette the widgets paint with
$script:Palette = @{
    Accent  = '#5AA9FF'   # blue - the running / upcoming thing
    Soft    = '#79B4FF'   # the smaller blue next to it
    Good    = '#48D28A'   # done / free
    Warn    = '#F0B04A'   # paused / stale
    Title   = '#F1F6FB'   # the big numbers
    Text    = '#C3CEDA'   # normal text
    Muted   = '#8DA0B3'   # status line
    Faint   = '#74838F'   # footer
}

$script:Widget = @{}

# --------------------------------------------------------------------------
# state
# --------------------------------------------------------------------------
function Initialize-Widget {
<#
    .SYNOPSIS Claim the single-instance mutex, work out the paths, load state.
#>
    param(
        [Parameter(Mandatory)][string]$Id,           # 'workclock' - file / mutex name
        [Parameter(Mandatory)][string]$Name,         # 'WorkClock' - window title, shortcut name
        [string]$StateFile,                          # default "<id>.json"
        [System.Collections.IDictionary]$Defaults = @{},
        [string]$ScriptPath,
        [scriptblock]$BeforeSave
    )

    if (-not $StateFile) { $StateFile = "$Id.json" }
    if (-not $ScriptPath) { $ScriptPath = $PSCommandPath }

    # only one of each widget at a time - a second one would fight over the state file
    $isFirst = $false
    $mutex = New-Object System.Threading.Mutex($true, "Local\Widget_$Id", [ref]$isFirst)
    if (-not $isFirst) { exit }

    $appDir = Join-Path $env:LOCALAPPDATA 'WorkClock'
    $root   = Split-Path (Split-Path $ScriptPath -Parent) -Parent

    # everything every widget keeps, plus whatever the widget adds on top
    $state = [ordered]@{
        backdrop = 'glass'    # 'glass' (see-through) | 'acrylic' (Win11 frosted)
        tint     = 3          # index into the tint palettes, 0 = most transparent
        blur     = $false     # blur what is behind the card instead of showing it
        left     = $null
        top      = $null
    }
    foreach ($k in $Defaults.Keys) { $state[$k] = $Defaults[$k] }

    $script:Widget = @{
        Id         = $Id
        Name       = $Name
        AppDir     = $appDir
        StatePath  = Join-Path $appDir $StateFile
        ScriptPath = $ScriptPath
        RootDir    = $root
        Mutex      = $mutex
        State      = $state
        BeforeSave = $BeforeSave
        Window     = $null
        Hwnd       = $null
        Placed     = $false   # the card still sits wherever the OS dropped it
        Rescued    = $false   # ...or had to be pulled back from a missing screen
        UserMoved  = $false   # the card has been dragged by hand this session
    }

    Import-WidgetState
    # remember the persisted position: LocationChanged overwrites left/top as
    # soon as the window is created, before we get a chance to restore it
    $script:Widget.SavedLeft = $state.left
    $script:Widget.SavedTop  = $state.top

    # the Windows 11 system acrylic is only worth offering where it exists
    $script:Widget.Frosted = ($state.backdrop -eq 'acrylic') -and ([Environment]::OSVersion.Version.Build -ge 22000)

    $state
}

function Get-WidgetState { $script:Widget.State }

function Get-WidgetPalette { $script:Palette }

function Get-WidgetDataPath([Parameter(Mandatory)][string]$FileName) {
    Join-Path $script:Widget.AppDir $FileName
}

function Import-WidgetState {
    $path = $script:Widget.StatePath
    if (-not (Test-Path $path)) { return }
    try {
        $j = Get-Content $path -Raw | ConvertFrom-Json
        $s = $script:Widget.State
        foreach ($k in @($s.Keys)) {
            if ($null -ne $j.PSObject.Properties[$k]) { $s[$k] = $j.$k }
        }
    } catch { }
}

function Save-WidgetState {
    try {
        if ($script:Widget.BeforeSave) { & $script:Widget.BeforeSave }
        if (-not (Test-Path $script:Widget.AppDir)) {
            New-Item -ItemType Directory -Path $script:Widget.AppDir -Force | Out-Null
        }
        # depth: flat values need none, a widget keeping a list of things does
        ([pscustomobject]$script:Widget.State) | ConvertTo-Json -Depth 6 |
            Set-Content -Path $script:Widget.StatePath -Encoding UTF8
    } catch { }
}

function Write-WidgetLog([string]$Message) {
    if ($env:WORKCLOCK_DEBUG) {
        try {
            Add-Content -Path (Join-Path $env:TEMP 'workclock.log') `
                        -Value ("{0} [{1}] {2}" -f (Get-Date -f 'HH:mm:ss'), $script:Widget.Id, $Message)
        } catch { }
    }
}

# --------------------------------------------------------------------------
# little helpers the widgets share
# --------------------------------------------------------------------------
function Format-HM([double]$Seconds) {
    $neg   = $Seconds -lt 0
    $total = [math]::Floor([math]::Abs($Seconds))
    $h     = [int][math]::Floor($total / 3600)
    $m     = [int][math]::Floor(($total % 3600) / 60)
    '{0}{1}:{2:00}' -f $(if ($neg) { '-' } else { '' }), $h, $m
}

function ConvertFrom-Duration([string]$Text) {
    # accepts "7:30", "7h30", "450m", "7.5"
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim().ToLower() -replace '\s', ''
    if ($t -match '^(\d+)[:h](\d{1,2})m?$') { return [int]$matches[1] * 60 + [int]$matches[2] }
    if ($t -match '^(\d+)m$')               { return [int]$matches[1] }
    if ($t -match '^(\d+)h$')               { return [int]$matches[1] * 60 }
    if ($t -match '^(\d+([.,]\d+)?)$')      { return [int][math]::Round([double]($matches[1] -replace ',', '.') * 60) }
    return $null
}

function Show-WidgetInput([string]$Prompt, [string]$Default = '') {
    [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $script:Widget.Name, $Default)
}

function Show-WidgetMessage([string]$Text) {
    [Windows.MessageBox]::Show($Text, $script:Widget.Name) | Out-Null
}

function Show-WidgetConfirm([string]$Text) {
    [Windows.MessageBox]::Show($Text, $script:Widget.Name,
        [Windows.MessageBoxButton]::OKCancel, [Windows.MessageBoxImage]::Question) -eq 'OK'
}

function Get-WidgetBrush([string]$Color) {
    ([Windows.Media.BrushConverter]::new()).ConvertFrom($Color)
}

function Test-WidgetOnScreen([double]$Left, [double]$Top, [double]$Width, [double]$Height) {
<#
    .SYNOPSIS Is enough of a card at this spot reachable with the mouse?

    A remembered position can point at a monitor that has since been unplugged,
    and a card placed there is simply invisible - it looks like the widget never
    started. Enough of the title bar area has to land inside the virtual desktop.
#>
    $l = [Windows.SystemParameters]::VirtualScreenLeft
    $t = [Windows.SystemParameters]::VirtualScreenTop
    $r = $l + [Windows.SystemParameters]::VirtualScreenWidth
    $b = $t + [Windows.SystemParameters]::VirtualScreenHeight
    ($Left + 60 -lt $r) -and ($Left + $Width - 60 -gt $l) -and
    ($Top + 24 -lt $b)  -and ($Top + $Height - 24 -gt $t)
}

# --------------------------------------------------------------------------
# the card
# --------------------------------------------------------------------------
function New-WidgetWindow {
<#
    .SYNOPSIS Build the frameless translucent card every widget lives in.
    .PARAMETER Body      XAML that goes inside the card's Grid (rows and all).
    .PARAMETER MenuItems XAML for the widget's own context-menu entries; the
                         shared Background / autostart / Quit entries follow.
    .PARAMETER Height    fixed height, or 0 to grow with the content.
#>
    param(
        [Parameter(Mandatory)][string]$Body,
        [string]$MenuItems = '',
        [int]$Width = 0,
        [int]$Height = 0,
        [string]$Padding = '15,9,15,13'
    )
    if ($Width -le 0) { $Width = $script:WIDGET_WIDTH }
    $frosted = $script:Widget.Frosted

    if ($frosted) {
        $winAttrs  = 'AllowsTransparency="False" Background="Transparent"'
        $cardAttrs = 'Margin="0" CornerRadius="0" BorderThickness="0"'
    } else {
        # no outer margin: the blur fills the window rect and is clipped to a
        # rounded region, so the card has to be the window
        $winAttrs  = 'AllowsTransparency="True" Background="Transparent"'
        $cardAttrs = "Margin=`"0`" CornerRadius=`"$script:CORNER`" BorderBrush=`"#30FFFFFF`" BorderThickness=`"1`""
    }
    $sizeAttrs = if ($Height -gt 0) { "Height=`"$Height`"" } else { 'SizeToContent="Height"' }
    $title     = [Security.SecurityElement]::Escape($script:Widget.Name)

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$title" Width="$Width" $sizeAttrs
        WindowStyle="None" ResizeMode="NoResize" Topmost="True"
        ShowInTaskbar="False" WindowStartupLocation="Manual"
        FontFamily="Segoe UI" UseLayoutRounding="True" $winAttrs>
  <Window.Resources>
    <Style x:Key="Glyph" TargetType="Button">
      <Setter Property="Foreground" Value="#C3CEDA"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Width" Value="22"/>
      <Setter Property="Height" Value="20"/>
      <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bg" CornerRadius="5" Background="{TemplateBinding Background}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bg" Property="Background" Value="#30FFFFFF"/>
                <Setter Property="Foreground" Value="#FFFFFF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Border x:Name="Card" $cardAttrs Padding="$Padding">
    <Border.ContextMenu>
      <ContextMenu x:Name="Menu">
$MenuItems
        <MenuItem Header="Background">
          <MenuItem x:Name="MiLighter"  Header="Lighter&#9;(scroll up on the widget)"/>
          <MenuItem x:Name="MiDarker"   Header="Darker&#9;(scroll down)"/>
          <Separator/>
          <MenuItem x:Name="MiBlur"     Header="Blur what is behind" IsCheckable="True"/>
          <MenuItem x:Name="MiBackdrop" Header="Windows acrylic" IsCheckable="True"/>
        </MenuItem>
        <MenuItem x:Name="MiStartup" Header="Start with Windows" IsCheckable="True"/>
        <Separator/>
        <MenuItem x:Name="MiQuit"    Header="Quit"/>
      </ContextMenu>
    </Border.ContextMenu>

    <Grid>
      <Grid.Effect>
        <DropShadowEffect BlurRadius="6" ShadowDepth="0" Opacity="0.65" Color="#000000"/>
      </Grid.Effect>
$Body
    </Grid>
  </Border>
</Window>
"@

    $win = [Windows.Markup.XamlReader]::Parse($xaml)
    $script:Widget.Window = $win
    $win
}

function Get-WidgetElements {
<#
    .SYNOPSIS Pull the named elements out of the window into a hashtable.
#>
    param([Parameter(Mandatory)]$Window, [Parameter(Mandatory)][string[]]$Names)
    $h = @{}
    foreach ($n in $Names) { $h[$n] = $Window.FindName($n) }
    $h
}

# --------------------------------------------------------------------------
# background: tint, blur, Windows acrylic
# --------------------------------------------------------------------------
function Set-WidgetTint([int]$Level) {
    $pal   = if ($script:Widget.Frosted) { $script:TINTS_FROSTED } else { $script:TINTS_GLASS }
    $Level = [math]::Max(0, [math]::Min($pal.Count - 1, $Level))
    $script:Widget.State.tint = $Level
    $card = $script:Widget.Window.FindName('Card')
    if ($card) { $card.Background = Get-WidgetBrush $pal[$Level] }
}

function Set-WidgetBlur {
    # 3 = blur whatever is behind the window, 0 = leave it sharp and see-through
    if ($script:Widget.Frosted -or $null -eq $script:Widget.Hwnd) { return }
    [WC.Native]::ApplyBlur($script:Widget.Hwnd, $(if ($script:Widget.State.blur) { 3 } else { 0 }), 0)
}

# --------------------------------------------------------------------------
# start with Windows
# --------------------------------------------------------------------------
function Get-WidgetStartupLink {
    Join-Path ([Environment]::GetFolderPath('Startup')) ("{0}.lnk" -f $script:Widget.Name)
}

function Get-WidgetStartupTarget {
    # what the startup shortcut has to run to bring *this* copy back up
    $vbs = Join-Path $script:Widget.RootDir 'Widget.vbs'
    if (Test-Path $vbs) {
        @{ Target = "$env:WINDIR\System32\wscript.exe"
           Args   = "`"$vbs`" $($script:Widget.Id)" }
    } else {
        @{ Target = 'powershell.exe'
           Args   = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$($script:Widget.ScriptPath)`"" }
    }
}

function Set-WidgetStartup([bool]$Enabled) {
    $lnkPath = Get-WidgetStartupLink
    if (-not $Enabled) { Remove-Item $lnkPath -Force -ErrorAction SilentlyContinue; return }

    $want = Get-WidgetStartupTarget
    $sh   = New-Object -ComObject WScript.Shell
    $lnk  = $sh.CreateShortcut($lnkPath)
    $lnk.TargetPath       = $want.Target
    $lnk.Arguments        = $want.Args
    $lnk.WorkingDirectory = $script:Widget.RootDir
    $lnk.Description      = $script:Widget.Name
    $lnk.Save()
}

function Repair-WidgetStartup {
<#
    .SYNOPSIS Point a dead startup shortcut back at the copy that is running.

    The shortcut holds the full path of the folder it was made from, so moving
    or renaming that folder leaves a link that quietly starts nothing at all -
    the widget simply never comes up after a restart. Rewrite such a link the
    next time the widget runs; leave a link that still works alone, so a copy
    installed under %LOCALAPPDATA% is not hijacked by a run from the sources.
#>
    $lnkPath = Get-WidgetStartupLink
    if (-not (Test-Path $lnkPath)) { return }
    $dead = $false
    try {
        $lnk = (New-Object -ComObject WScript.Shell).CreateShortcut($lnkPath)
        # both forms quote the script they hand to wscript / powershell
        if ($lnk.Arguments -match '"([^"]+)"') { $dead = -not (Test-Path $matches[1]) }
    } catch { return }
    if ($dead) {
        Write-WidgetLog "startup shortcut was stale - repointing it at $($script:Widget.RootDir)"
        Set-WidgetStartup $true
    }
}

# --------------------------------------------------------------------------
# wire the shared behaviour to the window
# --------------------------------------------------------------------------
function Register-WidgetChrome {
<#
    .SYNOPSIS Drag, hover, scroll-to-veil, the shared menu entries, the
              remembered position and saving on close.
#>
    param([Parameter(Mandatory)]$Window)

    $ctx  = $script:Widget
    $S    = $ctx.State
    $card = $Window.FindName('Card')
    $ctrl = $Window.FindName('Controls')

    $miLighter  = $Window.FindName('MiLighter')
    $miDarker   = $Window.FindName('MiDarker')
    $miBlur     = $Window.FindName('MiBlur')
    $miBackdrop = $Window.FindName('MiBackdrop')
    $miStartup  = $Window.FindName('MiStartup')
    $miQuit     = $Window.FindName('MiQuit')

    $miLighter.Add_Click({ Set-WidgetTint ([int]$script:Widget.State.tint - 1); Save-WidgetState })
    $miDarker.Add_Click({  Set-WidgetTint ([int]$script:Widget.State.tint + 1); Save-WidgetState })
    $card.Add_MouseWheel({
        $step = if ($args[1].Delta -gt 0) { -1 } else { 1 }
        Set-WidgetTint ([int]$script:Widget.State.tint + $step)
        Save-WidgetState
    })

    $miBlur.IsChecked = [bool]$S.blur
    $miBlur.IsEnabled = -not $ctx.Frosted
    $miBlur.Add_Click({
        $script:Widget.State.blur = [bool]$script:Widget.Window.FindName('MiBlur').IsChecked
        Set-WidgetBlur; Save-WidgetState
    })

    $miBackdrop.IsChecked = $ctx.Frosted
    $miBackdrop.Add_Click({
        $w = $script:Widget
        $w.State.backdrop = if ($w.Window.FindName('MiBackdrop').IsChecked) { 'acrylic' } else { 'glass' }
        Save-WidgetState
        Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
            "Start-Sleep -Milliseconds 1500; & '$($w.ScriptPath)'")
        $w.Window.Close()
    })

    try { Repair-WidgetStartup } catch { }
    $miStartup.IsChecked = Test-Path (Get-WidgetStartupLink)
    $miStartup.Add_Click({
        try { Set-WidgetStartup ([bool]$script:Widget.Window.FindName('MiStartup').IsChecked) }
        catch { Show-WidgetMessage "Could not change autostart: $($_.Exception.Message)" }
    })

    $miQuit.Add_Click({ Save-WidgetState; $script:Widget.Window.Close() })

    # drag anywhere, hover reveals the buttons
    $card.Add_MouseLeftButtonDown({
        try { $script:Widget.Window.DragMove(); $script:Widget.UserMoved = $true } catch { }
    })
    if ($ctrl) {
        $card.Add_MouseEnter({ $script:Widget.Window.FindName('Controls').Opacity = 1.0 })
        $card.Add_MouseLeave({ $script:Widget.Window.FindName('Controls').Opacity = 0.18 })
    }

    $Window.Add_SourceInitialized({
        $w = $script:Widget
        $h = (New-Object Windows.Interop.WindowInteropHelper $w.Window).Handle
        [WC.Native]::MakeToolWindow($h)
        if ($w.Frosted) {
            $src = [Windows.Interop.HwndSource]::FromHwnd($h)
            $src.CompositionTarget.BackgroundColor = [Windows.Media.Colors]::Transparent
            [WC.Native]::ApplyAcrylic($h)
        } else {
            $w.Hwnd = $h
            [WC.Native]::RoundCorners($h)
            Set-WidgetBlur
        }
    })

    # A drag only changes the numbers in memory. Quit saves them, but a Windows
    # restart kills the widget without ever raising Closing, so without this the
    # card comes back wherever it happened to be the last time something else
    # wrote the state file - hours or days ago. Write the new spot out a moment
    # after the card comes to rest.
    $moveSave = New-Object Windows.Threading.DispatcherTimer
    $moveSave.Interval = [timespan]::FromSeconds(1.5)
    $moveSave.Add_Tick({ $script:Widget.MoveSave.Stop(); Save-WidgetState })
    $ctx.MoveSave = $moveSave

    $Window.Add_ContentRendered({
        $w   = $script:Widget
        if ($w.Placed) { return }
        $win = $w.Window
        $wa  = [Windows.SystemParameters]::WorkArea
        if ($null -ne $w.SavedLeft -and $null -ne $w.SavedTop -and
            (Test-WidgetOnScreen ([double]$w.SavedLeft) ([double]$w.SavedTop) $win.ActualWidth $win.ActualHeight)) {
            $win.Left = [double]$w.SavedLeft
            $win.Top  = [double]$w.SavedTop
        } else {
            # no position yet, or it points at a screen that is not there right
            # now - park it in the lower right corner of the primary screen
            $win.Left = $wa.Right  - $win.ActualWidth  - 24
            $win.Top  = $wa.Bottom - $win.ActualHeight - 24
            # a card pulled back from a missing monitor keeps its old spot on
            # file, so unplugging a screen for an hour does not lose the layout
            $w.Rescued = ($null -ne $w.SavedLeft)
        }
        $w.Placed = $true
        if (-not $w.Rescued) { $w.State.left = $win.Left; $w.State.top = $win.Top }
        Write-WidgetLog "placed at $($win.Left),$($win.Top) size $($win.ActualWidth)x$($win.ActualHeight) frosted=$($w.Frosted) rescued=$($w.Rescued)"
    })

    $Window.Add_LocationChanged({
        $w = $script:Widget
        if (-not $w.Placed) { return }                      # still the OS default spot
        if ($w.Rescued -and -not $w.UserMoved) { return }   # keep the spot on the missing screen
        $w.State.left = $w.Window.Left
        $w.State.top  = $w.Window.Top
        $w.MoveSave.Stop(); $w.MoveSave.Start()
    })

    # restart / shutdown: Windows never sends this window a close, it just goes
    # away, so take the chance to flush whatever the debounce is still holding
    $ctx.OnSessionEnding = [Microsoft.Win32.SessionEndingEventHandler]{
        try { $script:Widget.Window.Dispatcher.Invoke([action]{ Save-WidgetState }) } catch { }
    }
    try { [Microsoft.Win32.SystemEvents]::add_SessionEnding($ctx.OnSessionEnding) } catch { }
    $Window.Add_Closed({
        try { [Microsoft.Win32.SystemEvents]::remove_SessionEnding($script:Widget.OnSessionEnding) } catch { }
    })

    $Window.Add_Closing({
        $w = $script:Widget
        Write-WidgetLog 'closing'
        $w.MoveSave.Stop()
        if ($w.Placed -and -not ($w.Rescued -and -not $w.UserMoved)) {
            $w.State.left = $w.Window.Left
            $w.State.top  = $w.Window.Top
        }
        Save-WidgetState
    })

    Set-WidgetTint ([int]$S.tint)
}

Export-ModuleMember -Function Initialize-Widget, Get-WidgetState, Get-WidgetPalette, Get-WidgetDataPath,
    Save-WidgetState, Write-WidgetLog, Format-HM, ConvertFrom-Duration,
    Show-WidgetInput, Show-WidgetMessage, Show-WidgetConfirm, Get-WidgetBrush,
    New-WidgetWindow, Get-WidgetElements, Register-WidgetChrome,
    Set-WidgetTint, Set-WidgetBlur, Set-WidgetStartup, Get-WidgetStartupLink, Repair-WidgetStartup,
    Test-WidgetOnScreen

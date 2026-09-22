<#
    Todos - a short checklist that lives on the desktop.

    Same card as the other widgets: a line to type into on top, the list
    underneath. Enter adds, the checkbox crosses it out, the x on the line
    removes it. The list scrolls once it is longer than the card.

    Run:  powershell -NoProfile -ExecutionPolicy Bypass -File widgets\todos.ps1
    Silent launcher: Widget.vbs todos
#>

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'lib\Widget.psm1') -Force

$ROW_H    = 22    # one todo line
$MAX_ROWS = 9     # beyond this the list scrolls instead of growing the card

# --------------------------------------------------------------------------
# state
# --------------------------------------------------------------------------
$script:S = Initialize-Widget -Id 'todos' -Name 'Todos' -ScriptPath $PSCommandPath `
    -BeforeSave { $script:S.items = @($script:items) } `
    -Defaults ([ordered]@{
        items        = @()
        showFinished = $true
    })

# The state file hands back PSCustomObjects; everything below wants dictionaries
# it can write to in place. An ArrayList rather than a List[object] because
# @($list) on a generic list throws on Windows PowerShell 5.1.
$script:items = New-Object Collections.ArrayList
foreach ($it in @($script:S.items)) {
    if (-not $it) { continue }
    $text = [string]$it.text
    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    [void]$script:items.Add([ordered]@{
        text    = $text
        done    = [bool]$it.done
        created = [string]$it.created
        doneAt  = [string]$it.doneAt
    })
}

$script:wantOffset  = 0.0
$script:scrollToEnd = $false

function Add-Todo([string]$Text) {
    $t = $Text.Trim()
    if (-not $t) { return }
    [void]$script:items.Add([ordered]@{
        text    = $t
        done    = $false
        created = (Get-Date).ToString('s')
        doneAt  = ''
    })
    $script:scrollToEnd = $true
    Save-WidgetState; Update-Ui
}

function Set-TodoDone($Item, [bool]$Done) {
    if (-not $Item) { return }
    $Item.done   = $Done
    $Item.doneAt = if ($Done) { (Get-Date).ToString('s') } else { '' }
    Save-WidgetState; Update-Ui
}

function Remove-Todo($Item) {
    if (-not $Item) { return }
    $script:items.Remove($Item)
    Save-WidgetState; Update-Ui
}

function Clear-Finished {
    $done = @($script:items | Where-Object { $_.done })
    if (-not $done.Count) { return }
    foreach ($d in $done) { $script:items.Remove($d) }
    Save-WidgetState; Update-Ui
}

# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
$menu = @'
        <MenuItem x:Name="MiShowDone" Header="Show finished" IsCheckable="True"/>
        <Separator/>
        <MenuItem x:Name="MiClear"    Header="Clear finished"/>
        <MenuItem x:Name="MiClearAll" Header="Remove all todos..."/>
        <Separator/>
'@

$body = @"
      <Grid.Resources>
        <!-- the checkbox: a small rounded box that fills in green when ticked -->
        <Style x:Key="TodoCheck" TargetType="CheckBox">
          <Setter Property="Cursor" Value="Hand"/>
          <Setter Property="Focusable" Value="False"/>
          <Setter Property="VerticalAlignment" Value="Center"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="CheckBox">
                <Border x:Name="box" Width="14" Height="14" CornerRadius="4"
                        Background="#14FFFFFF" BorderBrush="#5D6E7E" BorderThickness="1.2">
                  <TextBlock x:Name="tick" Text="&#xE73E;" FontFamily="Segoe MDL2 Assets"
                             FontSize="8" Foreground="#101820" Visibility="Collapsed"
                             HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <ControlTemplate.Triggers>
                  <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="box" Property="BorderBrush" Value="#9BB0C4"/>
                    <Setter TargetName="box" Property="Background"  Value="#26FFFFFF"/>
                  </Trigger>
                  <Trigger Property="IsChecked" Value="True">
                    <Setter TargetName="box"  Property="Background"  Value="#48D28A"/>
                    <Setter TargetName="box"  Property="BorderBrush" Value="#48D28A"/>
                    <Setter TargetName="tick" Property="Visibility"  Value="Visible"/>
                  </Trigger>
                </ControlTemplate.Triggers>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Style>

        <!-- a hairline scrollbar instead of the grey system one on the glass -->
        <Style TargetType="ScrollBar">
          <Setter Property="Width" Value="4"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="ScrollBar">
                <Border Background="#10FFFFFF" CornerRadius="2" Width="4">
                  <Track x:Name="PART_Track" IsDirectionReversed="True"
                         Minimum="{TemplateBinding Minimum}" Maximum="{TemplateBinding Maximum}"
                         ViewportSize="{TemplateBinding ViewportSize}"
                         Value="{Binding Value, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}">
                    <Track.Thumb>
                      <Thumb>
                        <Thumb.Template>
                          <ControlTemplate TargetType="Thumb">
                            <Border CornerRadius="2" Background="#4DFFFFFF" MinHeight="18"/>
                          </ControlTemplate>
                        </Thumb.Template>
                      </Thumb>
                    </Track.Thumb>
                  </Track>
                </Border>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Style>
      </Grid.Resources>

      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Grid Grid.Row="0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse x:Name="Dot" Width="7" Height="7" Fill="#5AA9FF" Margin="1,0,8,0"/>
          <TextBlock x:Name="StatusText" Text="TODOS" FontSize="9" FontWeight="SemiBold" Foreground="#8DA0B3"/>
        </StackPanel>
        <StackPanel x:Name="Controls" Orientation="Horizontal" HorizontalAlignment="Right" Opacity="0.18">
          <Button x:Name="BtnClear" Style="{StaticResource Glyph}" Content="&#xE74D;" ToolTip="Clear finished"/>
          <Button x:Name="BtnQuit"  Style="{StaticResource Glyph}" Content="&#xE8BB;" ToolTip="Quit"/>
        </StackPanel>
      </Grid>

      <Grid x:Name="AddRow" Grid.Row="1" Height="$ROW_H" Background="Transparent" Margin="0,8,0,6">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="&#xE710;" FontFamily="Segoe MDL2 Assets" FontSize="10"
                   Width="14" TextAlignment="Center" Foreground="#5AA9FF"
                   VerticalAlignment="Center" Margin="1,0,9,0"/>
        <Grid Grid.Column="1">
          <TextBlock x:Name="Hint" Text="add a todo..." FontSize="11.5" Foreground="#6C7B89"
                     VerticalAlignment="Center" IsHitTestVisible="False"/>
          <TextBox x:Name="AddBox" MaxLength="140" FontSize="11.5" Foreground="#E2EAF2"
                   Background="Transparent" BorderThickness="0" Padding="0" CaretBrush="#5AA9FF"
                   SelectionBrush="#5AA9FF" VerticalContentAlignment="Center"/>
        </Grid>
      </Grid>

      <Border Grid.Row="2" Height="1" Background="#20FFFFFF" Margin="0,0,0,6"/>

      <ScrollViewer x:Name="Scroller" Grid.Row="3" MaxHeight="$($MAX_ROWS * $ROW_H)"
                    VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                    Padding="0,0,3,0">
        <StackPanel x:Name="List"/>
      </ScrollViewer>

      <Grid Grid.Row="4" Margin="0,7,0,0">
        <TextBlock x:Name="FootLeft"  Text="" FontSize="10.5" Foreground="#74838F"/>
        <TextBlock x:Name="FootRight" HorizontalAlignment="Right" Text="" FontSize="10.5" Foreground="#74838F"/>
      </Grid>
"@

$win = New-WidgetWindow -Body $body -MenuItems $menu
$script:UI = Get-WidgetElements $win @(
    'Card','Dot','StatusText','Controls','BtnClear','BtnQuit','AddRow','Hint','AddBox',
    'Scroller','List','FootLeft','FootRight','MiShowDone','MiClear','MiClearAll')

$P = Get-WidgetPalette
$brushAccent = Get-WidgetBrush $P.Accent
$brushGood   = Get-WidgetBrush $P.Good
$brushText   = Get-WidgetBrush $P.Text
$brushFaint  = Get-WidgetBrush $P.Faint

# the row styles live in the body's Grid.Resources, so ask from inside the Grid
$script:checkStyle = $UI.List.FindResource('TodoCheck')
$script:glyphStyle = $win.FindResource('Glyph')

# --------------------------------------------------------------------------
# rendering
# --------------------------------------------------------------------------
function Get-TodoTip($Item) {
    $lines = @($Item.text)
    if ($Item.created) {
        try { $lines += 'added ' + ([datetime]$Item.created).ToString('ddd HH:mm') } catch { }
    }
    if ($Item.done -and $Item.doneAt) {
        try { $lines += 'done ' + ([datetime]$Item.doneAt).ToString('ddd HH:mm') } catch { }
    }
    ($lines -join "`n")
}

function New-TodoRow($Item) {
    $g = New-Object Windows.Controls.Grid
    $g.Height     = $ROW_H
    $g.Background = [Windows.Media.Brushes]::Transparent   # so the whole line reacts to hover
    $g.ToolTip    = Get-TodoTip $Item
    foreach ($w in @(0, -1, 0)) {
        $c = New-Object Windows.Controls.ColumnDefinition
        $c.Width = if ($w -lt 0) { [Windows.GridLength]::new(1, 'Star') } else { [Windows.GridLength]::Auto }
        $g.ColumnDefinitions.Add($c)
    }

    # every handler reads what it works on off the sender's Tag - a scriptblock
    # closing over $Item would see the last row of the loop instead
    $cb = New-Object Windows.Controls.CheckBox
    $cb.Style     = $script:checkStyle
    $cb.IsChecked = [bool]$Item.done
    $cb.Margin    = '1,0,9,0'
    $cb.Tag       = $Item
    $cb.Add_Click({ param($s, $e) Set-TodoDone $s.Tag ([bool]$s.IsChecked) })
    [Windows.Controls.Grid]::SetColumn($cb, 0)

    $t = New-Object Windows.Controls.TextBlock
    $t.Text              = [string]$Item.text
    $t.FontSize          = 11.5
    $t.VerticalAlignment = 'Center'
    $t.TextTrimming      = 'CharacterEllipsis'
    if ($Item.done) {
        $t.Foreground      = $brushFaint
        $t.TextDecorations = [Windows.TextDecorations]::Strikethrough
    } else {
        $t.Foreground = $brushText
    }
    [Windows.Controls.Grid]::SetColumn($t, 1)

    $x = New-Object Windows.Controls.Button
    $x.Style   = $script:glyphStyle
    $x.Content = [char]0xE711
    $x.Width   = 18
    $x.Height  = 18
    $x.Opacity = 0            # only the line under the pointer shows its x
    $x.ToolTip = 'Remove'
    $x.Tag     = $Item
    $x.Add_Click({ param($s, $e) Remove-Todo $s.Tag })
    [Windows.Controls.Grid]::SetColumn($x, 2)

    $g.Tag = $x            # the row hands its own x back to the hover handlers
    $g.Add_MouseEnter({ param($s, $e) $s.Tag.Opacity = 1 })
    $g.Add_MouseLeave({ param($s, $e) $s.Tag.Opacity = 0 })

    [void]$g.Children.Add($cb)
    [void]$g.Children.Add($t)
    [void]$g.Children.Add($x)
    $g
}

function New-NoteRow([string]$Text) {
    $t = New-Object Windows.Controls.TextBlock
    $t.Text              = $Text
    $t.Height            = $ROW_H
    $t.FontSize          = 11.5
    $t.Foreground        = $brushFaint
    $t.Margin            = '24,0,0,0'
    $t.VerticalAlignment = 'Center'
    $t
}

function Update-Ui {
    $ui   = $script:UI
    $open = @($script:items | Where-Object { -not $_.done })
    $done = @($script:items | Where-Object { $_.done })
    $show = if ([bool]$script:S.showFinished) { @($script:items) } else { $open }

    $script:wantOffset = $ui.Scroller.VerticalOffset
    $ui.List.Children.Clear()
    foreach ($it in $show) { [void]$ui.List.Children.Add((New-TodoRow $it)) }
    if (-not $show.Count) {
        [void]$ui.List.Children.Add((New-NoteRow $(
            if ($done.Count) { 'nothing open' } else { 'nothing to do' })))
    }

    # rebuilding the list drops the scroll position - put it back once the new
    # rows have been measured. Priority first: the other overload would hand the
    # priority to a scriptblock that takes no arguments.
    [void]$ui.Scroller.Dispatcher.BeginInvoke(
        [Windows.Threading.DispatcherPriority]::Loaded, [action]{
            $sv = $script:UI.Scroller
            if ($script:scrollToEnd) { $script:scrollToEnd = $false; $sv.ScrollToEnd() }
            else { $sv.ScrollToVerticalOffset($script:wantOffset) }
        })

    if ($script:items.Count -eq 0) {
        $ui.StatusText.Text = 'TODOS'
        $ui.Dot.Fill        = $brushFaint
    } elseif ($open.Count -eq 0) {
        $ui.StatusText.Text = 'ALL DONE'
        $ui.Dot.Fill        = $brushGood
    } else {
        $ui.StatusText.Text = 'TODOS'
        $ui.Dot.Fill        = $brushAccent
    }

    $ui.FootLeft.Text  = if ($open.Count -eq 0) { '' }
                         elseif ($open.Count -eq 1) { '1 open' }
                         else { "$($open.Count) open" }
    $ui.FootRight.Text = if ($done.Count) { "$($done.Count) done" } else { '' }

    $ui.BtnClear.Visibility = if ($done.Count) { 'Visible' } else { 'Collapsed' }
    $ui.MiClear.IsEnabled   = [bool]$done.Count
}

# --------------------------------------------------------------------------
# behaviour
# --------------------------------------------------------------------------
$UI.AddBox.Add_TextChanged({
    $ui = $script:UI
    $ui.Hint.Visibility = if ($ui.AddBox.Text.Length) { 'Collapsed' } else { 'Visible' }
})

$UI.AddBox.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq 'Return') {
        $e.Handled = $true
        Add-Todo $s.Text
        $s.Clear()
    } elseif ($e.Key -eq 'Escape') {
        $e.Handled = $true
        $s.Clear()
    }
})

# the whole line is the input, not just the letters in it
$UI.AddRow.Add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true; [void]$script:UI.AddBox.Focus() })

# scrolling the list scrolls it; scrolling a list that fits veils the card like
# everywhere else, which the ScrollViewer would otherwise swallow
$UI.Scroller.Add_PreviewMouseWheel({
    param($s, $e)
    if ($s.ScrollableHeight -le 0) {
        $e.Handled = $true
        Set-WidgetTint ([int]$script:S.tint + $(if ($e.Delta -gt 0) { -1 } else { 1 }))
        Save-WidgetState
    }
})

$clear = { Clear-Finished }
$UI.BtnClear.Add_Click($clear)
$UI.MiClear.Add_Click($clear)

$UI.MiClearAll.Add_Click({
    $n = $script:items.Count
    if ($n -eq 0) { return }
    if (-not (Show-WidgetConfirm "Remove all $n todos?")) { return }
    $script:items.Clear()
    Save-WidgetState; Update-Ui
})

$UI.MiShowDone.IsChecked = [bool]$script:S.showFinished
$UI.MiShowDone.Add_Click({
    $script:S.showFinished = [bool]$script:UI.MiShowDone.IsChecked
    Save-WidgetState; Update-Ui
})

$UI.BtnQuit.Add_Click({ Save-WidgetState; $win.Close() })

Register-WidgetChrome $win

Update-Ui
Write-WidgetLog 'show'
[void]$win.ShowDialog()
Write-WidgetLog 'exited'

<#
    Copies the widgets to %LOCALAPPDATA%\WorkClock\app so they run without WSL,
    adds Start-menu shortcuts and (optionally) starts them with Windows.

    Usage:
      powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 [-Autostart]
      ... -Widgets workclock,agenda      only these (default: all of them)
      ... -Uninstall                     remove everything again
#>
param([switch]$Autostart, [switch]$Uninstall, [string[]]$Widgets)

$ErrorActionPreference = 'Stop'
$target  = Join-Path $env:LOCALAPPDATA 'WorkClock\app'
$startup = [Environment]::GetFolderPath('Startup')
$menu    = [Environment]::GetFolderPath('Programs')
$sh      = New-Object -ComObject WScript.Shell

$all = Get-Content (Join-Path $PSScriptRoot 'widgets.json') -Raw | ConvertFrom-Json
if ($Widgets) {
    $all = @($all | Where-Object { $Widgets -contains $_.id })
    if (-not $all) { throw "No widget matched -Widgets $($Widgets -join ',')" }
}

function New-Lnk($path, $widget) {
    $lnk = $sh.CreateShortcut($path)
    $lnk.TargetPath       = "$env:WINDIR\System32\wscript.exe"
    $lnk.Arguments        = "`"$target\Widget.vbs`" $($widget.id)"
    $lnk.WorkingDirectory = $target
    $lnk.Description      = "$($widget.name) - $($widget.description)"
    $lnk.Save()
}

if ($Uninstall) {
    foreach ($w in $all) {
        Remove-Item (Join-Path $startup "$($w.name).lnk"), (Join-Path $menu "$($w.name).lnk") `
                    -Force -ErrorAction SilentlyContinue
    }
    Remove-Item (Join-Path $startup 'WorkClock.lnk') -Force -ErrorAction SilentlyContinue   # pre-2.0 shortcut
    Remove-Item $target -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Removed. Your times and settings stay in $env:LOCALAPPDATA\WorkClock"
    return
}

New-Item -ItemType Directory -Path $target -Force | Out-Null
foreach ($dir in 'lib', 'widgets') {
    $dest = Join-Path $target $dir
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    Copy-Item (Join-Path $PSScriptRoot "$dir\*") $dest -Recurse -Force
}
# WorkClock.vbs is the pre-2.0 launcher - copy it along if it is still around
$loose = @('Widget.vbs', 'WorkClock.vbs', 'widgets.json') |
         ForEach-Object { Join-Path $PSScriptRoot $_ } | Where-Object { Test-Path $_ }
Copy-Item $loose $target -Force

foreach ($w in $all) {
    New-Lnk (Join-Path $menu "$($w.name).lnk") $w
    if ($Autostart) { New-Lnk (Join-Path $startup "$($w.name).lnk") $w }
}

Write-Host "Installed to $target"
Write-Host ("Widgets: " + (($all | ForEach-Object { $_.name }) -join ', '))
Write-Host "Start them from the Start menu, or run: wscript `"$target\Widget.vbs`" <widget>"
if ($Autostart) { Write-Host "They will also start with Windows." }

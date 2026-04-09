$LauncherDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "GhostPractice Launcher.lnk"
$targetPath = Join-Path $LauncherDir "Launch.bat"
$iconPath = Join-Path $LauncherDir "icon.ico"

$ws = New-Object -ComObject WScript.Shell
$shortcut = $ws.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetPath
$shortcut.WorkingDirectory = $LauncherDir
$shortcut.IconLocation = "$iconPath,0"
$shortcut.WindowStyle = 7
$shortcut.Description = "Launch GhostPractice with database selection"
$shortcut.Save()

Write-Host "Shortcut created: $shortcutPath"

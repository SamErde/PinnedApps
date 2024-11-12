$PinnedAppList = Get-PinnedApp
Write-Host 'All Pinned Apps:' -ForegroundColor Magenta -BackgroundColor Black
$PinnedAppList.Name

$StartMenuPinnedApps = Get-PinnedApp -StartMenu
Write-Host 'Apps pinned to the start menu: ' -ForegroundColor Green -BackgroundColor Black -NoNewline
Write-Host "$($StartMenuPinnedApps | Select-Object -ExpandProperty Name | Join-String -Separator ', ')" -ForegroundColor White -BackgroundColor Black -NoNewline
Write-Host ''

$TaskBarPinnedApps = Get-PinnedApp -Taskbar
Write-Host 'Apps pinned to the taskbar: ' -ForegroundColor Green -BackgroundColor Black -NoNewline
Write-Host "$($TaskbarPinnedApps | Select-Object -ExpandProperty Name | Join-String -Separator ', ')" -ForegroundColor White -BackgroundColor Black -NoNewline
Write-Host ''

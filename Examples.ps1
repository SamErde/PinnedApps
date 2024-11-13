$TaskBarPinnedApps = Get-PinnedApp
Write-Host 'Apps pinned to the taskbar: ' -ForegroundColor Green -BackgroundColor Black -NoNewline
Write-Host "$($TaskbarPinnedApps | Select-Object -ExpandProperty Name | Join-String -Separator ', ')" -ForegroundColor White -BackgroundColor Black -NoNewline
Write-Host ''

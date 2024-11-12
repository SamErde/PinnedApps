function Get-PinnedApps {
    [CmdletBinding()]
    param (
        # Start Menu
        [Parameter()]
        [switch]
        $StartMenu,

        # Taskbar
        [Parameter()]
        [switch]
        $TaskBar
    )

    switch ($PSBoundParameters.Keys) {
        'StartMenu' {
            $Pattern = 'Unpi&n from Start'
        }
        'TaskBar' {
            $Pattern = 'Unpin from tas&kbar'
        }
        Default {
            $Pattern = 'Unpi[&]?n'
        }
    }
    if (-not $Pattern) {
        $Pattern = 'Unpi[&]?n'
    }
    Write-Debug -Message $Pattern

    $Apps = (New-Object -ComObject Shell.Application).Namespace('shell:AppsFolder').Items()

    [System.Collections.Generic.List[Object]]$PinnedAppList = @()
    foreach ($app in $Apps) {
        $AppVerbs = $app.Verbs()
        $PinnedApp = $AppVerbs | Where-Object { $_.Name -match $Pattern }
        if ($PinnedApp) {
            $PinnedAppList.Add($App)
        }
    }

    # Output the list of pinned apps
    $PinnedAppList # | ForEach-Object { Write-Host "Pinned app: $_" }
    Write-Debug -Message $PinnedAppList.Name
}

$PinnedAppList = Get-PinnedApps
$PinnedAppList.ToString()

$StartMenuPinnedApps = Get-PinnedApps -StartMenu
Write-Host "Start Menu:`n$($StartMenuPinnedApps | Select-Object Name)"

$TaskBarPinnedApps = Get-PinnedApps -Taskbar
Write-Host "Taskbar:`n$($TaskBarPinnedApps | Select-Object Name)"

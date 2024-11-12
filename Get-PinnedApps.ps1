function Get-PinnedApps {
    [CmdletBinding()]
    param (
        # Start Menu
        [Parameter(Mandatory, ParameterSetName = 'StartMenu')]
        [switch]
        $StartMenu,

        # Taskbar
        [Parameter(Mandatory, ParameterSetName = 'Taskbar')]
        [switch]
        $TaskBar
    )

    switch ($PSBoundParameters.Keys) {
        'StartMenu' {
            $Filter = 'Unpi&n from Start'
        }
        'TaskBar' {
            $Filter = 'Unpin from tas&kbar'
        }
    }

    $Apps = (New-Object -ComObject Shell.Application).Namespace('shell:AppsFolder').Items()

    [System.Collections.Generic.List[Object]]$PinnedAppList = @()
    foreach ($app in $Apps) {
        $AppVerbs = $app.Verbs()
        $PinnedApp = $AppVerbs | Where-Object { $_.Name -eq $Filter }
        if ($PinnedApp) {
            $PinnedAppList.Add($App)
        }
    }
    
    # Output the list of pinned apps
    $PinnedAppList # | ForEach-Object { Write-Host "Pinned app: $_" }
}

$StartMenuPinnedApps = Get-PinnedApps -StartMenu
Write-Host "Start Menu:`n$($StartMenuPinnedApps | Select-Object Name)"

$TaskBarPinnedApps = Get-PinnedApps -Taskbar
Write-Host "Taskbar:`n$($TaskBarPinnedApps | Select-Object Name)"

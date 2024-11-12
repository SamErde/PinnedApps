function Get-PinnedApp {
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

    $PinnedAppList
}

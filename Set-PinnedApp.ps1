function Set-PinnedApp {
    <#
        .SYNOPSIS
        Pin an app to the Start Menu or Taskbar

        .DESCRIPTION
        This function will help you pin an app to the start menu, the taskbar, or both.

        .EXAMPLE
        Set-PinnedApp -App 'Terminal' -Target StartMenu,Taskbar

        Pin the Windows Terminal app to the start menu and the taskbar.

        .NOTES
        Author: Sam Erde
        Version: 0.0.1
        Modified: 2024-11-12
    #>

    [CmdletBinding()]
    [Alias('Pin-App')]
    param (
        # App name or ID
        [Parameter(Mandatory)]
        [string]
        $App,

        # Target[s] to pin the app to
        [Parameter()]
        [ValidateSet('StartMenu', 'Taskbar')]
        [string[]]
        $Target
    )

    begin {

    } # end begin block

    process {

    } # end process block

    end {

    } # end end block

} # end function Set-PinnedApp

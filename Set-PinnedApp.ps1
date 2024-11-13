function Set-PinnedApp {
    <#
        .SYNOPSIS
        Pin an app to the Windows taskbar.

        .DESCRIPTION
        This function will help you pin an app to the Windows taskbar.

        .EXAMPLE
        Set-PinnedApp -App 'Terminal'

        Pin the Windows Terminal app to taskbar.

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
        $App
    )

    begin {

    } # end begin block

    process {

    } # end process block

    end {

    } # end end block

} # end function Set-PinnedApp

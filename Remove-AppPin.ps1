# Work in Progress.
# sde
# 2024/11/12

function Remove-AppPin {
    <#
        .SYNOPSIS
        Remove a pinned app from the taskbar.

        .DESCRIPTION
        This function will help you remove a pinned app from the taskbar.

        .EXAMPLE
        Remove-AppPin -App 'Terminal'

        Un-pin the Terminal app from the taskbar.

        .NOTES
        Author: Sam Erde
        Version: 0.0.1
        Modified: 2024-11-12
    #>

    [CmdletBinding()]
    [Alias('Unpin-App')]
    param (

    )

    begin {

    } # end begin block

    process {
        $Shell = New-Object -ComObject 'Shell.Application'
        $AppsFolder = $Shell.NameSpace('shell:AppsFolder')
        $Apps = $AppsFolder.Items() | Where-Object { $_.Name -eq $App }

        if ($App) {
            $AppVerbs = $App.Verbs()
            $UnpinVerb = $AppVerbs | Where-Object { $_.Name -eq 'Unpin from tas&kbar' }
            if ($UnpinVerb) {
                $App.InvokeVerb('Unpin from tas&kbar')
            }
        }
    } # end process block

    end {

    } # end end block

} # end function Remove-AppPin

# Get all apps
$Apps = (New-Object -ComObject Shell.Application).Namespace('shell:AppsFolder').Items()

# Unpin each app
foreach ($App in $Apps) {
    Remove-AppPin -AppName $App
}

# Restart Explorer to apply changes
Stop-Process -Name explorer -Force
Start-Process explorer

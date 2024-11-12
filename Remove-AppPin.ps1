# Work in Progress. Does not successfully un-pin apps yet. Have only tried with start menu so far.
# sde
# 2024/11/12

function Remove-AppPin {
    <#
        .SYNOPSIS
        Remove a pinned app from the start menu or taskbar.

        .DESCRIPTION
        This function will help you remove a pinned app from the start menu or task bar (or both).

        .EXAMPLE
        Remove-AppPin -App 'Terminal' -Target StartMenu

        Un-pin the Terminal app from the start menu.

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
            $UnpinVerb = $AppVerbs | Where-Object { $_.Name -eq 'Unpi&n from Start' }
            if ($UnpinVerb) {
                $App.InvokeVerb('Unpi&n from Start')
            }
        }
    } # end process block

    end {

    } # end end block

} # end function Remove-AppPin

# Get all Start Menu apps
$StartApps = (New-Object -ComObject Shell.Application).Namespace('shell:AppsFolder').Items()

# Unpin each app
foreach ($App in $StartApps) {
    Remove-AppPin -AppName $App
}

# Restart Explorer to apply changes
Stop-Process -Name explorer -Force
Start-Process explorer

function Get-PinnedApp {
    [CmdletBinding()]
    param ( )

    $Pattern = 'Unpin from tas&kbar'
    $Apps = (New-Object -ComObject Shell.Application).Namespace('shell:AppsFolder').Items() |
        Where-Object {
        ($_.Verbs() | Select-Object -ExpandProperty Name) -contains $Pattern
        }

    [System.Collections.Generic.List[PSObject]]$PinnedAppList = @()
    foreach ($app in $Apps) {
        Write-Verbose $app.Name
        Write-Verbose $app.Path
        $Properties = [ordered]@{
            Name = $app.Name
            Path = $app.Path
        }
        $item = New-Object -TypeName psobject -Property $Properties
        Write-Verbose $item
        $PinnedAppList.Add($item)
    }
    $PinnedAppList
}

$PinnedAppList = Get-PinnedApp
$PinnedAppList

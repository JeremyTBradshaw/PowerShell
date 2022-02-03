function Get-ConfigurationDataAsObject {
    <#
        .Synopsis
        PowerShell 4.0 alternative for Import-PowerShellDataFile, which was only introduced in PowerShell 5.0.

        .Link
        https://powershellmagazine.com/2016/05/11/pstip-convert-powershell-data-file-to-an-object
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [Microsoft.PowerShell.DesiredStateConfiguration.ArgumentToConfigurationDataTransformation()]
        [hashtable] $ConfigurationData
    )
    $ConfigurationData
}

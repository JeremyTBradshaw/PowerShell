function Get-WebsiteLogLocation {
<#
    .Synopsis
    Get-WebsiteLogLocation

    .Description
    Results are grouped by number of distinct users per day.

    .Parameter Credential
    Credential object to be used with Invoke-Command against IIS servers.

    .Parameter Server
    One or more server names/FQDN's (array or single).  Default is local computer.
#>
#Requires -Version 4
[CmdletBinding()]
param(
    [System.Management.Automation.PSCredential]$Credential,
    [string[]]$Server = @("$($env:COMPUTERNAME)")
)

begin {

    $LogUNCPaths = @()
}

process {

    foreach ($s in $Server) {

        $InvokeCmdProps = @{
            ComputerName        = $s
            HideComputerName    = $true
            ScriptBlock         = {
                                    Get-WebSite 'Default Web Site' |
                                    Select-Object -ExpandProperty LogFile
                                }
            ErrorAction         = 'Stop'
        }

        if ($PSBoundParameters.ContainsKey('Credential')) {$InvokeCmdProps['Credential'] = $Credential}

        $LogDirectory = Invoke-Command @InvokeCmdProps
        $LogUNCPaths += "\\$($s)\$($LogDirectory.Directory -replace ':','$' -replace '%SystemDrive%','C$')\W3SVC1"
    }

    $LogUNCPaths
}

end {}
}

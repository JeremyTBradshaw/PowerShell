<#
    .SYNOPSIS
    Wrapper for Get-MessageTrackingLog (no trailing 's') which repeats the search against all Transport Servers.

    .NOTES
    2023-08-11: Script is not complete, not intended for use, etc.
#>
[CmdletBinding()]
param (
    [string[]]$Recipients,
    [datetime]$Start#,
    #Remaining parameters to come soon.
)
begin {
    Write-Warning -Message '2023-08-11: Script is not complete, not intended for use, etc.'
    $TS = Get-TransportService
}
process {}
end {
    foreach ($_ts in $TS) {
        Get-MessageTrackingLog -Server $_ts.Name -Recipients $Recipients -Start $Start
    }
}

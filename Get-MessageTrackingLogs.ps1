<#
    .SYNOPSIS
    Wrapper for Get-MessageTrackingLog (no trailing 's') which repeats the search against all Transport Servers.

    .NOTES
    2023-08-11: Script is not complete, not intended for use, etc.
#>
[CmdletBinding()]
param (
    [string[]]$Recipients,
    [Parameter(Mandatory, HelpMessage = 'I insist.')]
    [datetime]$Start#,
    #Remaining parameters to come soon.
)
begin {
    Write-Warning -Message '2023-08-11: Script is not complete, not intended for use, etc.'

    # Eventually will use the following to enable all of the original parameters for pass-through:
    $cmdGetMsgTrkLog = Get-Command Get-MessageTrackingLog
    $cmdGetMsgTrkLog.Parameters.Values | ForEach-Object {
        $_ | Select-Object Name,
        @{
            Name       = 'Type'
            Expression = { $_.ParameterType.Name }
        },
        IsDynamic, SwitchParameter,
        @{
            Name       = 'Aliases'
            Expression = { $_.Aliases -join ', ' }
        },
        @{
            Name       = 'Mandatory'
            Expression = { $_.Attributes.Mandatory }
        },
        @{
            Name       = 'HelpMessage'
            Expression = { $_.Attributes.HelpMessage }
        },
        @{
            Name       = 'ValueFromPipeline'
            Expression = { $_.Attributes.ValueFromPipeline }
        },
        @{
            Name       = 'ValueFromPipelineByPropertyName'
            Expression = { $_.Attributes.ValueFromPipelineByPropertyName }
        }
    }

    # Will loop through these servers' MsgTrkLog's:
    $TS = Get-TransportService
}
process {}
end {
    foreach ($_ts in $TS) {
        Get-MessageTrackingLog -Server $_ts.Name -Recipients $Recipients -Start $Start
    }
}

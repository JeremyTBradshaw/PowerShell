<#
    .SYNOPSIS
    Helper script for Get-Command to get detailed information about a command's paramters.

    .PARAMETER CommandName
    Passthrough to Get-Command's -Name parameter (expects the entire Verb-Noun, e.g., Get-ChildItem)
#>
[CmdletBinding()]
param (
    [string]$CommandName
)
try {
    $Command = Get-Command -Name $CommandName -ErrorAction:Stop
    $ParameterSets = $Command.Parameters.Values.ParameterSets.Keys | Select-Object -Unique
    foreach ($pSet in $ParameterSets) {
        $Command.Parameters.Values | Where-Object { $_.ParameterSets.Keys -eq $pSet } |
        Select-Object Name,
        @{
            Name       = 'ParameterSetName'
            Expression = { $pSet }
        },
        @{
            Name       = 'Position'
            Expression = { $_.ParameterSets.$pSet.Position }
        },
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
            Expression = { $_.ParameterSets.$pSet.IsMandatory }
        },
        @{
            Name       = 'HelpMessage'
            Expression = { $_.ParameterSets.$pSet.HelpMessage }
        },
        @{
            Name       = 'ValueFromPipeline'
            Expression = { $_.ParameterSets.$pSet.ValueFromPipeline }
        },
        @{
            Name       = 'ValueFromPipelineByPropertyName'
            Expression = { $_.ParameterSets.$pSet.ValueFromPipelineByPropertyName }
        }
    }
}
catch { throw }

<#
    .SYNOPSIS
    Check Exchange Transport Servers' Back Pressure status (and thresholds).

    .PARAMETER Server
    Specify one or more Exchange server names.

    .PARAMETER All
    Check all servers (found by Get-TransportService).

    .EXAMPLE
    .\Get-BackPressureStatus.ps1

    .EXAMPLE
    .\Get-BackPressureStatus.ps1 -Server DAG02MB05

    .EXAMPLE
    Get-ExchangeServer DAG03MB0* | .\Get-BackPressureStatus.ps1 | ft -GroupBy Server

    .LINK
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-BackPressureStatus.ps1

    .LINK
    https://learn.microsoft.com/en-us/exchange/mail-flow/back-pressure#view-back-pressure-resource-thresholds-and-utilization-levels
#>
#Requires -Version 4.0
[CmdletBinding(DefaultParameterSetName = 'AllServers')]
param (
    [Parameter(
        ParameterSetName = 'IndividualServers',
        ValueFromPipeline, ValueFromPipelineByPropertyName
    )]
    [Alias('ServerName')]
    [string[]]$Server,
    [Parameter(ParameterSetName = 'AllServers')]
    [switch]$All
)
begin {
    $_requiredCommands = @('Get-ExchangeDiagnosticInfo')
    if ($PSCmdlet.ParameterSetName -eq 'AllServers') { $_requiredCommands += 'Get-TransportService' }

    foreach ($_cmd in $_requiredCommands) { if (-not (Get-Command $_cmd -ea SilentlyContinue)) { $_missingCommands += $_cmd } }
    if ($_missingCommands.Count -ge 1) { throw "Missing required commands: $($_missingCommands -join ', ')" }

    if ($PSCmdlet.ParameterSetName -eq 'AllServers') { $Script:Server = Get-TransportService -ea Stop | Select-Object -ExpandProperty Name }
}
process {
    foreach ($srv in $Server) {

        [xml]$perServerBPDiagInfo = Get-ExchangeDiagnosticInfo -Server $srv -Process EdgeTransport -Component ResourceThrottling -ea Stop
        foreach ($rsrc in $perServerBPDiagInfo.Diagnostics.Components.ResourceThrottling.ResourceTracker.ResourceMeter) {

            $rsrc | Select-Object @{Name = 'Server'; Expression = { $srv } },
            @{Name = 'Resource'; Expression = { $_.Resource -replace '\[.*' } },
            CurrentResourceUse,
            PreviousResourceUse,
            Pressure,
            @{
                Name       = 'PressureTransitions'
                Expression = { $_.PressureTransitions -replace '(Pressure.*\:\s)|(ow)|(edium)|(igh)' -replace '(To)', '>' }
            },
            @{Name = 'ResourceFullName'; Expression = { $_.Resource } }
        }
    }
}
end {}

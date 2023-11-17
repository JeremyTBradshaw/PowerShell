<#
    .SYNOPSIS
    Helper script to overcome the challenges with paginated results from Get-QuarantineMessage.  All parameters except
    for -Identity, -Page, and -PageSize are supported.  All pages will be retuned (or EXO will barf, one or the other).

    .NOTES
    Last updated: 2023-11-17
    Every parameter is a direct passthrough to Get-QuarantineMessage.  See the help file for that and just know that
    pagination is taken care of for you here, so nevermind the -Page / -PageSize parameters.  The -Identity parameter
    has also been omitted since it doesn't make sense in the use case for this script.

    .OUTPUTS
    The output matches that of Get-QuarantineMessage from the ExchangeOnlineManagement module.
#>
#Requires -Modules @{ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.4.0'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'}

[CmdletBinding(DefaultParameterSetName = 'Summary')]
param (
    [Parameter(ParameterSetName = 'Summary')]
    [ValidateSet('Inbound', 'Outbound')]
    [string]$Direction,

    [Parameter(ParameterSetName = 'Summary')]
    [string[]]$Domain,

    [Parameter(ParameterSetName = 'Summary')]
    [datetime]$EndExpiresDate,

    [Parameter(ParameterSetName = 'Summary')]
    [datetime]$EndReceivedDate,

    [Parameter(ParameterSetName = 'Summary')]
    [ValidateSet('Email', 'SharePointOnline', 'Teams')]
    [string]$EntityType,

    [Parameter(ParameterSetName = 'Summary')]
    [string]$MessageId,

    [Parameter(ParameterSetName = 'Summary')]
    [switch]$MyItems,

    [Parameter(ParameterSetName = 'Summary')]
    [string]$PolicyName,

    [Parameter(ParameterSetName = 'Summary')]
    [ValidateSet('AntiMalwarePolicy', 'AntiPhishPolicy', 'ExchangeTransportRule', 'HostedContentFilterPolicy', 'SafeAttachmentPolicy')]
    [string[]]$PolicyTypes,

    [Parameter(ParameterSetName = 'Summary')]
    [ValidateSet('Bulk', 'FileTypeBlock', 'HighConfPhish', 'Malware', 'Phish', 'Spam', 'SPOMalware', 'TransportRule')]
    [string[]]$QuarantineTypes,

    [Parameter(ParameterSetName = 'Summary')]
    [Parameter(ParameterSetName = 'Details')]
    [string[]]$RecipientAddress,

    [Parameter(ParameterSetName = 'Summary')]
    [string[]]$RecipientTag,

    [Parameter(ParameterSetName = 'Summary')]
    [ValidateSet('Approved', 'Denied', 'Error', 'NotReleased', 'PreparingToRelease', 'Released', 'Requested')]
    [string[]]$ReleaseStatus,

    [Parameter(ParameterSetName = 'Summary')]
    [bool]$Reported,

    [Parameter(ParameterSetName = 'Summary')]
    [string[]]$SenderAddress,

    [Parameter(ParameterSetName = 'Summary')]
    [datetime]$StartExpiresDate,

    [Parameter(ParameterSetName = 'Summary')]
    [datetime]$StartReceivedDate,

    [Parameter(ParameterSetName = 'Summary')]
    [string]$Subject,

    [Parameter(ParameterSetName = 'Summary')]
    [ValidateSet('Bulk', 'HighConfPhish', 'Malware', 'Phish', 'Spam', 'SPOMalware', 'TransportRule')]
    [string]$Type
)

if (-not (Get-Command Get-QuarantineMessage -ErrorAction SilentlyContinue)) {

    throw "This script requires an active connection to Exchange Online PowerShell (using v3.0.0 module or newer), and access to the Get-QuarantineMessage cmdlet."
}

$dtNow = [datetime]::Now
$Progress = @{
    Activity        = "$($PSCmdlet.MyInvocation.MyCommand.Name) - Start time: $($dtNow)"
    PercentComplete = -1
}

$gqmParams = @{
    Page        = 0
    PageSize    = 1000
    ErrorAction = 'Stop'
}
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Direction')) { $gqmParams['Direction'] = $Direction }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Domain')) { $gqmParams['Domain'] = $Domain }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('EndExpiresDate')) { $gqmParams['EndExpiresDate'] = $EndExpiresDate }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('EndReceivedDate')) { $gqmParams['EndReceivedDate'] = $EndReceivedDate }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('EntityType')) { $gqmParams['EntityType'] = $EntityType }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('MessageId')) { $gqmParams['MessageId'] = $MessageId }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('MyItems')) { $gqmParams['MyItems'] = $MyItems }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('PolicyName')) { $gqmParams['PolicyName'] = $PolicyName }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('PolicyTypes')) { $gqmParams['PolicyTypes'] = $PolicyTypes }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('QuarantineTypes')) { $gqmParams['QuarantineTypes'] = $QuarantineTypes }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('RecipientAddress')) { $gqmParams['RecipientAddress'] = $RecipientAddress }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('RecipientTag')) { $gqmParams['RecipientTag'] = $RecipientTag }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('ReleaseStatus')) { $gqmParams['ReleaseStatus'] = $ReleaseStatus }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Reported')) { $gqmParams['Reported'] = $Reported }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('SenderAddress')) { $gqmParams['SenderAddress'] = $SenderAddress }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('StartExpiresDate')) { $gqmParams['StartExpiresDate'] = $StartExpiresDate }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('StartReceivedDate')) { $gqmParams['StartReceivedDate'] = $StartReceivedDate }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Subject')) { $gqmParams['Subject'] = $Subject }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Type')) { $gqmParams['Type'] = $Type }

$InLoopResults = @(1)
$TotalResultsCount = 0

try {
    do {
        $gqmParams['Page']++
        # Next line is to get around EXO v3* issue which sets $ProgressPreference to SilentlyContinue globally:
        $Global:ProgressPreference = 'Continue'
        Write-Progress @Progress -Status "Page: $($gqmParams['Page']) | Quarantine messages found: $($TotalResultsCount)"

        $InLoopResults = @(Get-QuarantineMessage @gqmParams)
        if ($InLoopResults.Count -gt 0) {
            (1..$InLoopResults.Count) | ForEach-Object { $TotalResultsCount++ }
            $InLoopResults
        }
    }
    until (($InLoopResults.Count -eq 0) -or ($gqmParams['Page'] -eq 1000))
}
catch { throw }

Write-Progress @Progress -Completed

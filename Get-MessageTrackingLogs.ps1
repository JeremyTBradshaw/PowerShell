<#
    .SYNOPSIS
    Wrapper for Get-MessageTrackingLog (no trailing 's') which repeats the search against all Transport Servers.  All
    parameters except for -Server are supported, since -Server is handled automatically (all available servers are
    searched).

    .NOTES
    Last updated: 2023-11-20
#>
[CmdletBinding()]
param (
    [object]$DomainController,
    [datetime]$End,
    [ValidateSet(
        'AGENTINFO', 'BADMAIL', 'CLIENTSUBMISSION', 'DEFER', 'DELIVER', 'DELIVERFAIL', 'DROP', 'DSN', 'DUPLICATEDELIVER', 'DUPLICATEEXPAND',
        'DUPLICATEREDIRECT', 'EXPAND', 'FAIL', 'HADISCARD', 'HARECEIVE', 'HAREDIRECT', 'HAREDIRECTFAIL', 'INITMESSAGECREATED', 'LOAD',
        'MODERATIONEXPIRE', 'MODERATORAPPROVE', 'MODERATORREJECT', 'MODERATORSALLNDR', 'NOTIFYMAPI', 'NOTIFYSHADOW', 'POISONMESSAGE',
        'PROCESS', 'PROCESSMEETINGMESSAGE', 'RECEIVE', 'REDIRECT', 'RESOLVE', 'RESUBMIT', 'RESUBMITDEFER', 'RESUBMITFAIL', 'SEND', 'SUBMIT',
        'SUBMITDEFER', 'SUBMITFAIL', 'SUPPRESSED', 'THROTTLE', 'TRANSFER'
    )]
    [object]$EventId,
    [object]$InternalMessageId,
    [object]$MessageId,
    [object]$MessageSubject,
    [object]$NetworkMessageId,
    [object[]]$Recipients,
    [object]$Reference,
    [object]$ResultSize,
    [object]$Sender,
    [ValidateSet(
        'ADMIN', 'AGENT', 'APPROVAL', 'BOOTLOADER', 'DNS', 'DSN', 'GATEWAY', 'MAILBOXRULE', 'MEETINGMESSAGEPROCESSOR', 'ORAR',
        'PICKUP', 'POISONMESSAGE', 'PUBLICFOLDER', 'QUEUE', 'REDUNDANCY', 'RESOLVER', 'ROUTING', 'SAFETYNET', 'SMTP', 'STOREDRIVER'
    )]
    [object]$Source,
    [datetime]$Start,
    [object]$TransportTrafficType
)

if (-not (Get-Command Get-MessageTrackingLog, Get-TransportService -ErrorAction SilentlyContinue)) {

    throw 'This script requires an active connection to Exchange server remote PowerShell, and access to the Get-MessageTrackingLog and Get-TransportService cmdlets.'
}

try { $TransportServers = Get-TransportService -ErrorAction Stop | Where-Object { $_.MessageTrackingLogEnabled -eq $true } }
catch { throw }

$dtNow = [datetime]::Now
$Progress = @{
    Activity = "$($PSCmdlet.MyInvocation.MyCommand.Name) - Start time: $($dtNow)"
}

$gmtlParams = @{ ErrorAction = 'Stop' }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('MessageId')) { $gmtlParams['MessageId'] = $MessageId }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('ResultSize')) { $gmtlParams['ResultSize'] = $ResultSize }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('NetworkMessageId')) { $gmtlParams['NetworkMessageId'] = $NetworkMessageId }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Source')) { $gmtlParams['Source'] = $Source }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Start')) { $gmtlParams['Start'] = $Start }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Sender')) { $gmtlParams['Sender'] = $Sender }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('InternalMessageId')) { $gmtlParams['InternalMessageId'] = $InternalMessageId }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Recipients')) { $gmtlParams['Recipients'] = $Recipients }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('EventId')) { $gmtlParams['EventId'] = $EventId }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('TransportTrafficType')) { $gmtlParams['TransportTrafficType'] = $TransportTrafficType }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('MessageSubject')) { $gmtlParams['MessageSubject'] = $MessageSubject }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('DomainController')) { $gmtlParams['DomainController'] = $DomainController }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('End')) { $gmtlParams['End'] = $End }
if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Reference')) { $gmtlParams['Reference'] = $Reference }

$SearchResults = @()
$ServerCounter = 0
try {
    foreach ($srv in $TransportServers) {
        $ServerCounter++
        Write-Progress @Progress -Status "Message Tracking Log results found: $($SearchResults.Count)" -PercentComplete (($ServerCounter / $TransportServers.Count) * 100)

        $gmtlParams['Server'] = $srv.Name
        $SearchResults += Get-MessageTrackingLog @gmtlParams
    }
}
catch { throw }

Write-Progress @Progress -Completed

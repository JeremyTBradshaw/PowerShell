<#
    .SYNOPSIS
    A wrapper for Get-CalendarProcessing with output focused on "Booking Options" (as shown in EXO Admin Center).

    .PARAMETER Identity
    Passthrough for Get-CalendarProcessing's -Identity parameter.

    .PARAMETER ReturnAllCalendarProcessingProperties
    Specifies to return all properties that come out of Get-CalendarProcessing, instead of just the typical/default few properties.

    .NOTES
    v2026-03-12 - Added paramter -ReturnAllCalendarProcessingProperties.
#>
#Requires -Version 5.1
[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [string[]]$Identity,
    [switch]$ReturnAllCalendarProcessingProperties
)
begin {
    if ((Get-Command Get-CalendarProcessing, Get-Recipient -ErrorAction SilentlyContinue).Count -ne 2) {
        throw 'An active Exchange PowerShell session is required, along with access to the Get-CalendarProcessing and Get-Recipient cmdlets.'
    }
    $Script:startTime = [datetime]::Now
    $Script:stopwatchMain = [System.Diagnostics.Stopwatch]::StartNew()
    $Script:stopwatchPipeline = [System.Diagnostics.Stopwatch]::new()
    $Script:progress = @{
        Activity        = "$($PSCmdlet.MyInvocation.MyCommand.Name)"
        Status          = "Start time: $($startTime.ToString('yyyy-MM-ddTHH:mm:ss')) | Elapsed: $($stopWatchMain.Elapsed.ToString('hh\:mm\:ss'))"
        PercentComplete = -1
    }
    Write-Progress @progress

    $Script:ht_rcptTracker = @{}
    function getRecipient ([string]$rcptId) {
        if ($Script:ht_rcptTracker.ContainsKey($rcptId)) { $Script:ht_rcptTracker[$rcptId] }
        else {
            $rcpt = @(Get-Recipient -Filter "(Name -eq '$($rcptId)') -or (UserPrincipalName -eq '$($rcptId)') -or LegacyExchangeDNRaw -eq '$($rcptId)'" $rcptId -ErrorAction SilentlyContinue)
            if ($rcpt.Count -eq 1) { $Script:ht_rcptTracker[$rcptId] = $rcpt.PrimarySmtpAddress.ToString() }
            elseif ($rcpt.Count -gt 1) { $Script:ht_rcptTracker[$rcptId] = "AMBIGUOUS_ACE('$($rcptId)'):{$($rcpt.PrimarySmtpAddress -join ', ')}" }
            else { $Script:ht_rcptTracker[$rcptId] = "UNKNOWN_ACE('$($rcptId)')" }
            $Script:ht_rcptTracker[$rcptId]
        }
    }

    $stopWatchPipeline.Start()
}
process {
    try {
        if (($PSCmdlet.MyInvocation.PipelinePosition -eq 0) -or ($stopWatchPipeline.ElapsedMilliseconds -ge 250)) {
            $Script:progress.CurrentOperation = "Resource: $($Identity[0])"
            $Script:progress.Status = "Start time: $($startTime.ToString('yyyy-MM-ddTHH:mm:ss')) | Elapsed: $($stopWatchMain.Elapsed.ToString('hh\:mm\:ss'))"
            Write-Progress @progress
            $stopWatchPipeline.Restart()
        }

        $_calProc = $null; $_calProc = Get-CalendarProcessing -Identity $Identity[0] -ErrorAction Stop
        $_output = [ordered]@{
            Identity                    = $Identity[0]
            AutomateProcessing          = $_calProc.AutomateProcessing
            ForwardRequestsToDelegates  = $_calProc.ForwardRequestsToDelegates
            ResourceDelegates           = if ($_calProc.ResourceDelegates) { ($_calProc.ResourceDelegates | ForEach-Object { getRecipient -rcptId $_ }) -join ', ' } else { $null }
            AllBookInPolicy             = $_calProc.AllBookInPolicy
            AllRequestInPolicy          = $_calProc.AllRequestInPolicy
            AllRequestOutOfPolicy       = $_calProc.AllRequestOutOfPolicy
            BookInPolicy                = if ($_calProc.BookInPolicy) { ($_calProc.BookInPolicy | ForEach-Object { getRecipient -rcptId $_ }) -join ', ' } else { $null }
            RequestInPolicy             = if ($_calProc.RequestInPolicy) { ($_calProc.RequestInPolicy | ForEach-Object { getRecipient -rcptId $_ }) -join ', ' } else { $null }
            RequestOutOfPolicy          = if ($_calProc.RequestOutOfPolicy) { ($_calProc.RequestOutOfPolicy | ForEach-Object { getRecipient -rcptId $_ }) -join ', ' } else { $null }
            AddAdditionalResponse       = $_calProc.AddAdditionalResponse
            AdditionalResponse          = $_calProc.AdditionalResponse
            EnableResponseDetails       = $_calProc.EnableResponseDetails
            EnforceSchedulingHorizon    = $_calProc.EnforceSchedulingHorizon
            ScheduleOnlyDuringWorkHours = $_calProc.ScheduleOnlyDuringWorkHours
            BookingWindowInDays         = $_calProc.BookingWindowInDays
            MaximumDurationInMinutes    = $_calProc.MaximumDurationInMinutes
            MinimumDurationInMinutes    = $_calProc.MinimumDurationInMinutes
            AllowRecurringMeetings      = $_calProc.AllowRecurringMeetings
            AllowConflicts              = $_calProc.AllowConflicts
            ConflictPercentageAllowed   = $_calProc.ConflictPercentageAllowed
            MaximumConflictInstances    = $_calProc.MaximumConflictInstances
        }
        if ($ReturnAllCalendarProcessingProperties) {
            $_output['AddNewRequestsTentatively'] = $_calProc.AddNewRequestsTentatively
            $_output['AddOrganizerToSubject'] = $_calProc.AddOrganizerToSubject
            $_output['AllowDistributionGroup'] = $_calProc.AllowDistributionGroup
            $_output['AllowMultipleResources'] = $_calProc.AllowMultipleResources
            $_output['BookingType'] = $_calProc.BookingType
            $_output['DeleteAttachments'] = $_calProc.DeleteAttachments
            $_output['DeleteComments'] = $_calProc.DeleteComments
            $_output['DeleteNonCalendarItems'] = $_calProc.DeleteNonCalendarItems
            $_output['DeleteSubject'] = $_calProc.DeleteSubject
            $_output['EnableAutoRelease'] = $_calProc.EnableAutoRelease
            $_output['EnforceAdjacencyAsOverlap'] = $_calProc.EnforceAdjacencyAsOverlap
            $_output['EnforceCapacity'] = $_calProc.EnforceCapacity
            $_output['OrganizerInfo'] = $_calProc.OrganizerInfo
            $_output['PostReservationMaxClaimTimeInMinute'] = $_calProc.PostReservationMaxClaimTimeInMinute
            $_output['ProcessExternalMeetingMessages'] = $_calProc.ProcessExternalMeetingMessages
            $_output['RemoveCanceledMeetings'] = $_calProc.RemoveCanceledMeetings
            $_output['RemoveForwardedMeetingNotification'] = $_calProc.RemoveForwardedMeetingNotification
            $_output['RemoveOldMeetingMessages'] = $_calProc.RemoveOldMeetingMessages
            $_output['RemovePrivateProperty'] = $_calProc.RemovePrivateProperty
            $_output['TentativePendingApproval'] = $_calProc.TentativePendingApproval
        }
        [PSCustomObject]$_output
    }
    catch { Write-Warning "Failed on Identity: $($Identity[0])"; throw }
}
end { Write-Progress @progress -Completed }

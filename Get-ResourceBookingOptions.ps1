<#
    .SYNOPSIS
    A wrapper for Get-CalendarProcessing with output focused on "Booking Options" (as shown in EXO Admin Center).

    .PARAMETER Identity
    Passthrough for Get-CalendarProcessing's -Identity parameter.
#>
#Requires -Version 5.1
[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [string[]]$Identity
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
            $rcpt = @(Get-Recipient -Filter "(Name -eq '$($rcptId)') -or (UserPrincipalName -eq '$($rcptId)')" $rcptId -ErrorAction SilentlyContinue)
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

        Get-CalendarProcessing -Identity $Identity[0] -ErrorAction Stop |
        Select-Object @{Name = 'Identity'; Expression = { $Identity[0] } },
        AutomateProcessing,
        ForwardRequestsToDelegates,
        @{
            Name       = 'ResourceDelegates'
            Expression = { ($_.ResourceDelegates | ForEach-Object { getRecipient -rcptId $_ }) -join ', ' }
        },
        AllBookInPolicy, AllRequestInPolicy, AllRequestOutOfPolicy,
        @{
            Name       = 'BookInPolicy'
            Expression = { ($_.BookInPolicy | ForEach-Object { getRecipient -rcptId $_ }) -join ', ' }
        },
        @{
            Name       = 'RequestInPolicy'
            Expression = { ($_.RequestInPolicy | ForEach-Object { getRecipient -rcptId $_ }) -join ', ' }
        },
        @{
            Name       = 'RequestOutOfPolicy'
            Expression = { ($_.RequestOutOfPolicy | ForEach-Object { getRecipient -rcptId $_ }) -join ', ' }
        },
        AddAdditionalResponse, AdditionalResponse,
        EnforceSchedulingHorizon, ScheduleOnlyDuringWorkHours, BookingWindowInDays, MaximumDurationInMinutes, MinimumDurationInMinutes,
        AllowRecurringMeetings, AllowConflicts, ConflictPercentageAllowed, MaximumConflictInstances
    }
    catch { Write-Warning "Failed on Identity: $($Identity[0])"; throw }
}
end { Write-Progress @progress -Completed }

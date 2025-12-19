<#
    .SYNOPSIS
    Clear calendar of events (only individual occurrences, not entire series) within a date range.

    .NOTES
    This script is not ready nor meant for use as of 2025-09-26. But I needed to park the code somewhere for now.
#>
#Requires -PSEdition Core
#Requires -Modules @{ModuleName = 'Microsoft.Graph.Calendar'; ModuleVersion = '2.2.8'; Guid = 'bf2ee476-ae5b-4a53-9fa8-943f3a49bf93'}
[CmdletBinding()]
param(
    [Parameter(Mandatory)][datetime]$seriesStartDate,
    [Parameter(Mandatory)][datetime]$occurrencesStartDateTime,
    [Parameter(Mandatory)][datetime]$occurrencesEndDateTime,
    [Parameter(Mandatory)][object]$UserId,
    [string]$DeclineComment = "This meeting occurrence has been administratively declined using Microsoft Graph PowerShell script 'Clear-MgUserCalendar.ps1'."
)
$Calendar = Get-MgUserCalendar -UserId $UserId -Filter "Name eq 'Calendar'"
$Events = Get-MgUserCalendarEvent -UserId $UserId -CalendarId $Calendar.Id -Filter "start/dateTime ge '$($seriesStartDate.ToUniversalTime().ToString("o"))'"
$Instances = @(foreach ($e in ($Events | Where-Object { $_.Type -eq 'seriesMaster' })) {
        Get-MgUserEventInstance -UserId $UserId -EventId $e.Id -StartDateTime $occurrencesStartDateTime.ToUniversalTime().ToString("o") -EndDateTime $occurrencesEndDateTime.ToUniversalTime().AddDays(1).ToString("o")
    })
# $Instances += $Events | Where-Object { $_.Type -ne 'seriesMaster' }
foreach ($i in $Instances) {
    Invoke-MgDeclineUserEventInstance -UserId $UserId -EventId $i.Id -EventId1 $i.Id -Comment $DeclineComment -SendResponse
    Write-Host -ForegroundColor Green "Declined '$($i.Subject)' occurrence: $(([datetime]$i.Start.DateTime).ToLocalTime())"
}

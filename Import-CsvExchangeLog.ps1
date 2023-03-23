<#
    .SYNOPSIS
    Effectively Import-Csv for Exchange log files, which have a few extra lines at the top of them.

    .LINK
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Import-CsvExchangeLog.ps1

    .NOTES
    Author: Jeremy Bradshaw (Jeremy.Bradshaw@Outlook.com)
    Version: 2023-03-23 11:00 AM (-03:00)
#>
#Requires -Version 4.0
[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
    [Alias('FullName','Path')]
    [System.IO.FileInfo[]]$LogFilePath
)
begin {
    $StartTime = [datetime]::Now
    $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
}
process {
    try {
        if ($StopWatch.ElapsedMilliseconds -ge 200) {
            $Progress = @{
                Activity = "Import-CsvExchangeLog.ps1 - Start time: $($StartTime)"
                PercentComplete = -1
                Status = "Importing CSV: $($LogFilePath)"
            }
            Write-Progress @Progress
            $StopWatch.Restart()
        }
        
        $LogContent = Get-Content -Path $LogFilePath -ErrorAction Stop

        $CsvHeaderIndex = if ($LogContent[0] -like '#Software: Microsoft*Server') { 4 }
        elseif ($LogContent[1] -eq '#Software: Microsoft*Server') { 5 }
        else {
            "Log file '$($LogFilePath)' doesn't fit the expect patterns for this script's supported Exchange log types.  " +
            "The 1st or 2nd row should begin with '#Software: Microsoft Exchange' or #Software: Microsoft Transport Server'.  " +
            "If this is a miss by the script, please let me know via GitHub Discussions on the repository." | Write-Warning
        }

        $CsvHeaders = $LogContent[$CsvHeaderIndex] -replace '#Fields: ' -split ','
        ConvertFrom-Csv -InputObject $LogContent[($CsvHeaderIndex+4)..($LogContent.Count-1)] -Header $CsvHeaders -ErrorAction Stop
    }
    catch { throw }
}

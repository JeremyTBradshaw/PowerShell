function writeLog {
    param(
        [Parameter(Mandatory)]
        [string]$LogName,

        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [System.IO.FileInfo]$Folder,
  
        [ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
		[datetime]$LogDateTime = [datetime]::Now,

        [switch]$DisableLogging
    )

    if (-not $DisableLogging -and -not $WhatIfPreference.IsPresent) {

        # Check for current log file and if necessary create it.
        $LogFile = Join-Path -Path $Folder -ChildPath "$($LogName)_$($LogDateTime.ToString('yyyy-MM-dd_HH-mm-ss')).log"
        if (-not (Test-Path $LogFile)) {
            try {
                [void](New-Item -Path $LogFile -ItemType:File -ErrorAction Stop)
            }
            catch {
                throw "Unable to create log file $($LogFile).  Unable to write to log."
            }
        }

        [string]$Date = Get-Date -Format 'yyyy-MM-dd hh:mm:ss tt'

        # Write message to log file:
        $MessageText = "[ $($Date) ] $($Message)"
        switch ($SectionStart) {

            $true { $MessageText = "`r`n" + $MessageText }
        }
        $MessageText | Out-File -FilePath $LogFile -Append -Encoding UTF8

        # If an error was supplied, write it to the log as well.
        if ($PSBoundParameters.ErrorRecord) {

            # Format the error as it would be displayed in the PS console.
            $ErrorForLog = "$($ErrorRecord.Exception)`r`n" +
            "$($ErrorRecord.InvocationInfo.PositionMessage)`r`n" +
            "`t+ CategoryInfo: " +
            "$($ErrorRecord.CategoryInfo.Category): " +
            "($($ErrorRecord.CategoryInfo.TargetName):$($ErrorRecord.CategoryInfo.TargetType))" +
            "[$($ErrorRecord.CategoryInfo.Activity)], " +
            "$($ErrorRecord.CategoryInfo.Reason)`r`n" +
            "`t+ FullyQualifiedErrorId: $($ErrorRecord.FullyQualifiedErrorId)"

            "[ $($Date) ][Error] $($ErrorForLog)" | Out-File -FilePath $LogFile -Append
        }
    }
}

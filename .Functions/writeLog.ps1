function writeLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$LogName,
        [Parameter(Mandatory)][datetime]$LogDateTime,
        [Parameter(Mandatory)][System.IO.FileInfo]$Folder,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)][string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [switch]$DisableLogging,
        [switch]$PassThru
    )

    if (-not $DisableLogging -and -not $WhatIfPreference.IsPresent) {
        try {
            if (-not (Test-Path -Path $Folder)) {

                [void](New-Item -Path $Folder -ItemType Directory -ErrorAction Stop)
            }

            $LogFile = Join-Path -Path $Folder -ChildPath "$($LogName)_$($LogDateTime.ToString('yyyy-MM-dd_HH-mm-ss')).log"
            if (-not (Test-Path $LogFile)) {

                [void](New-Item -Path $LogFile -ItemType:File -ErrorAction Stop)
            }

            $Date = [datetime]::Now.ToString('yyyy-MM-dd hh:mm:ss tt')

            "[ $($Date) ] $($Message)" | Out-File -FilePath $LogFile -Append

            if ($PSBoundParameters.ErrorRecord) {

                # Format the error as it would be displayed in the PS console.
                "[ $($Date) ][Error] $($ErrorRecord.Exception.Message)`r`n" +
                "$($ErrorRecord.InvocationInfo.PositionMessage)`r`n" +
                "`t+ CategoryInfo: $($ErrorRecord.CategoryInfo.Category): " +
                "($($ErrorRecord.CategoryInfo.TargetName):$($ErrorRecord.CategoryInfo.TargetType))" +
                "[$($ErrorRecord.CategoryInfo.Activity)], $($ErrorRecord.CategoryInfo.Reason)`r`n" +
                "`t+ FullyQualifiedErrorId: $($ErrorRecord.FullyQualifiedErrorId)`r`n" |
                Out-File -FilePath $LogFile -Append -ErrorAction Stop
            }
        }
        catch { throw }
    }

    if ($PassThru) { $Message }
    else { Write-Verbose -Message $Message }
}

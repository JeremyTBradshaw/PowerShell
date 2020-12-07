<#
    .Synopsis
    Testing writeLog with pipeline input of the main message and adding -PassThru logic.

    .Description
    The goal is to enable cleaner code when both logging as well as using Write-Warning/Write-Verbose.  For example:

    'This is a sample message to both log and Write-(Verbose|Warning).' |  writeLog @writeLogParams -PassThru | Write-Verbose

    Write-Warning example:

    'This is a sample message to both log and Write-(Verbose|Warning).' |  writeLog @writeLogParams -PassThru | Write-Warning

    .Outputs
    # Sample log entry:
    [yyyy-MM-dd hh:mm:ss tt] This is a sample message to both log and Write-(Verbose|Warning).

    # Sample Write-Verbose output:
    VERBOSE: This is a sample message to both log and Write-(Verbose|Warning).

    # Sample Write-Warning output:
    WARNING: This is a sample message to both log and Write-(Verbose|Warning).
#>
[CmdletBinding()]
param (
    [switch]$NoLog
)
#region Functions
function writeLog {
    param(
        [Parameter(Mandatory)]
        [string]$LogName,

        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$Message,

        [Parameter(Mandatory)]
        [System.IO.FileInfo]$Folder,

        [ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
        [datetime]$LogDateTime = [datetime]::Now,

        [switch]$DisableLogging,
        [switch]$PassThru
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
    # Output/passthru the message:
    if ($PassThru) { $Message }
}
#endregion Functions

#region Initialization
$dtNow = [datetime]::Now

$writeLogParams = @{

    LogName     = "$($MyInvocation.MyCommand.Name -replace '\.ps1')"
    Folder      = "$($PSScriptRoot)\$($MyInvocation.MyCommand.Name -replace '\.ps1')_Logs"
    LogDateTime = $dtNow
    ErrorAction = 'Stop'
}

if (-not $NoLog.IsPresent) {

    # Check for and if necessary create logs folder:
    if (-not (Test-Path -Path "$($writeLogParams['Folder'])")) {
        
        [void](New-Item -Path "$($writeLogParams['Folder'])" -ItemType Directory -ErrorAction Stop)
    }

    writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand) - Script begin."
    writeLog @writeLogParams -Message "PSScriptRoot: $($PSScriptRoot)"
    writeLog @writeLogParams -Message "Command: $($PSCmdlet.MyInvocation.Line)"
}
else {
    # Disable logging.
    $writeLogParams['DisableLogging'] = $true
}
#endregion Initialization

#region Main Script

<# Code would go here #>

#endregion Main Script

writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand) - Script end."

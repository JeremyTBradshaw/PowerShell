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
#Requires -Version 5.1
using namespace System.Management.Automation

[CmdletBinding(SupportsShouldProcess)]
param (
    [switch]$NoLog
)
#region Functions
function writeLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$LogName,
        [Parameter(Mandatory)][System.IO.FileInfo]$Folder,
        [Parameter(Mandatory, ValueFromPipeline)][string]$Message,
        [ErrorRecord]$ErrorRecord,
        [datetime]$LogDateTime = [datetime]::Now,
        [switch]$DisableLogging,
        [switch]$SectionStart,
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

            $Date = Get-Date -Format 'yyyy-MM-dd hh:mm:ss tt'
            $MessageText = "[ $($Date) ] $($Message)"
            switch ($SectionStart) {

                $true { $MessageText = "`r`n" + $MessageText }
            }
            $MessageText | Out-File -FilePath $LogFile -Append

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
        catch { throw $_ }
    }
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

if ($NoLog.IsPresent) { $writeLogParams['DisableLogging'] = $true }
#endregion Initialization

#region Main Script
"$($PSCmdlet.MyInvocation.MyCommand) - Script begin." | writeLog @writeLogParams -PassThru | Write-Verbose
"PSScriptRoot: $($PSScriptRoot)" | writeLog @writeLogParams -PassThru | Write-Verbose
"Command: $($PSCmdlet.MyInvocation.Line)" | writeLog @writeLogParams -PassThru | Write-Verbose

try { 1/0 }
catch {
    writeLog @writeLogParams -Message "Encountered issue." -ErrorRecord $_ -PassThru |
    Write-Warning
}
try { Get-ChildItem 23 -ErrorAction Stop }
catch {
    writeLog @writeLogParams -Message "Encountered issue." -ErrorRecord $_ -PassThru |
    Write-Warning
}
finally {
    writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand) - Script end." -PassThru | Write-Verbose
}
#endregion Main Script

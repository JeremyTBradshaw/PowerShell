<#
    .SYNOPSIS
    Script file for launching HealthChecker.ps1 via Task Scheduler, and from within the same directory.

    .DESCRIPTION
    This script is intended for use by a Windows Task Scheduler task where the Action as configured as follows:
    - Action: Start a program
    - Program/Script: PowerShell
    - Add arguments (optional): (option1) -Command " & 'E:\.Exchange-Reporting-and-Automation\.CSS-Exchange\HealthChecker.runHelper.ps1'"
    - Add arguments (optional): (option2) -Command " & 'E:\.Exchange-Reporting-and-Automation\.CSS-Exchange\HealthChecker.runHelper.ps1' -ExchangeServers 'exSvr1', 'exSvr2'"

    HealthChecker.ps1 is executed 3 times: once to update itself, once more to process the specified servers,
    generating the XML files, and one final time to generate the HTML report from the XML files.

    The path to the script can be changed as desired and updated in the arguments option for the task's Action.  What
    is important about the path of the script is that it is in the same directory as Healthchecker.ps1.  Once the
    script has been run for the first time, it will create two subfolders in the same directory as the script.  The
    script root folder will then contain the following items:

    <PSScriptRoot>
      - HealthChecker.ps1 #<--: official script (Version 23.05.15.1908 or newer recommended or HTML report will land in script caller's $PWD).
      - HealthChecker.runHelper.ps1 #<--: This script file.
      - HealthChecker.runHelper_Logs #<--: Folder containing a daily log file.  Modify 'LogRotation' in Initialization and Variables region's #1 to be 'Secondly' for a separate log file for every execution of this script.
      - HealthChecker.ps1_Outputs #<--: Folder containing server TXT/XML files, debug log, and other outputs from HealthChecker.ps1, saved in dedicated subfolders for each execution of this script.
      - HealthChecker_Report_2023-05-17T10-25-51.html #<--: The actual HTML reports will be dropped here, named after date/time to prevent overwriting earlier reports.

    .PARAMETER ExchangeServers
    Specifies one or more Exchange server names/FQDNs to target with HealthChecker.ps1.  This is a pass-through to
    HealthChecker.ps1's -Server Parameter.  Optionally, update the parameter's default value witin the param() block.

    .NOTES
    Author: Jeremy Bradshaw (Jeremy.Bradshaw@Outlook.com)
    WhenChanged: 2023-05-17 12:00 PM (Atlantic Standard Time)

    .LINK
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/HealthChecker.runHelper.ps1

    .LINK
    https://microsoft.github.io/CSS-Exchange/

    .LINK
    https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/RunHCViaSchedTask/

    .LINK
    https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1
#>
#Requires -Version 4
[CmdletBinding()]
param (
    [string[]]$ExchangeServers = @('exampleExSvr1', 'exampleExSvr2')
)

#======#-----------#
#region# Functions #
#======#-----------#

function writeLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$LogName,
        [Parameter(Mandatory)][datetime]$LogDateTime,
        [ValidateSet('Daily', 'Secondly')]
        [string]$LogRotation = 'Secondly',
        [Parameter(Mandatory)][System.IO.FileInfo]$Folder,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)][string]$Message,
        [switch]$SectionStart,
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [switch]$DisableLogging,
        [switch]$PassThru
    )

    if (-not $DisableLogging -and -not $WhatIfPreference.IsPresent) {
        try {
            if (-not (Test-Path -Path $Folder)) {

                [void](New-Item -Path $Folder -ItemType Directory -ErrorAction Stop)
            }

            $Rotation = switch ($LogRotation) {

                Secondly { $LogDateTime.ToString('yyyy-MM-dd_HH-mm-ss') }
                Daily { $LogDateTime.ToString('yyyy-MM-dd') }
            }
            $LogFile = Join-Path -Path $Folder -ChildPath "$($LogName)_$($Rotation).log"
            if (-not (Test-Path $LogFile)) {

                [void](New-Item -Path $LogFile -ItemType:File -ErrorAction Stop)
            }

            $Date = [datetime]::Now.ToString('yyyy-MM-dd hh:mm:ss tt')

            $LogOutput = "[ $($Date) ] $($Message)"
            if ($SectionStart) { $LogOutput = $LogOutput.Insert(0, "`r`n") }
            $LogOutput | Out-File -FilePath $LogFile -Append

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

#=========#-----------#
#endregion# Functions #
#=========#-----------#


try {
    #======#------------------------------#
    #region# Initialization and Variables #
    #======#------------------------------#

    # 1. Setup writeLog splat and test writeLog:
    $Script:dtNow = [datetime]::Now
    $Script:writeLogParams = @{

        DisableLogging = $DisableLogging
        LogName        = "$($PSCmdlet.MyInvocation.MyCommand.Name -replace '\.ps1')"
        LogDateTime    = $dtNow
        Folder         = "$($PSCmdlet.MyInvocation.MyCommand.Path -replace '\.ps1')_Logs"
        LogRotation    = 'Daily'
        ErrorAction    = 'Stop'
    }
    writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand.Name) - Start" -SectionStart
    writeLog @writeLogParams -Message "MyCommand: $($PSCmdlet.MyInvocation.Line)"

    # 2. Easy path helpers:
    $Script:scriptRoot = Split-Path -Path $PSCmdlet.MyInvocation.MyCommand.Path -Parent
    $Script:scriptRootParent = Split-Path -Path $ScriptRoot -Parent
    $Script:healthCheckerOutputsMain = Join-Path -Path $ScriptRoot -ChildPath HealthChecker.ps1_Outputs
    $Script:healthCheckerOutputsCurrent = Join-Path -Path $healthCheckerOutputsMain -ChildPath "HealthChecker.ps1_$($dtNow.ToString('yyyy-MM-ddTHH-mm-ss'))"

    # 3. Prepare HealthChecker.ps1 output folder:
    if (-not (Test-Path -Path $healthCheckerOutputsMain)) {
        
        [void](New-Item -Path $healthCheckerOutputsMain -ItemType Directory -ea Stop)
        writeLog @writeLogParams -Message "Created main HealthChecker output folder: $($healthCheckerOutputsMain)"
    }
    [void](New-Item -Path $healthCheckerOutputsCurrent -ItemType Directory -ea Stop)
    writeLog @writeLogParams -Message "Created HealthChecker outputs folder: $($healthCheckerOutputsCurrent)"

    # 4. Locate and update HealthChecker.ps1:
    $Script:healthCheckerPs1 = Join-Path -Path $scriptRoot -ChildPath 'HealthChecker.ps1'
    & $healthCheckerPs1 -ScriptUpdateOnly -OutputFilePath $healthCheckerOutputsCurrent
    writeLog @writeLogParams -Message 'Successfully checked for updated version of HealthChecker.ps1.'

    # 5. Announce/log which Exchange servers will be targeted with HealthChecker.ps1:
    writeLog @writeLogParams -Message "Will run HealthChecker.ps1 again the following servers:`r`n`t$($ExchangeServers -join ', ')"

    #=========#------------------------------#
    #endregion# Initialization and Variables #
    #=========#------------------------------#



    #======#-----------------------#
    #region# Run HealthChecker.ps1 #
    #======#-----------------------#

    # Run HealthChecker.ps1 against all servers:
    writeLog @writeLogParams -Message 'Starting to run HealthChecker.ps1...'
    & $healthCheckerPs1 -Server $ExchangeServers -OutputFilePath $healthCheckerOutputsCurrent
    writeLog @writeLogParams -Message 'Finished running HealthChecker.ps1.'

    # Build HTML report:
    writeLog @writeLogParams -Message 'Generating HTML report.'
    & $healthCheckerPs1 -BuildHtmlServersReport -HtmlReportFile "HealthChecker_Report_$($dtNow.ToString('yyyy-MM-ddTHH-mm-ss')).html" -XMLDirectoryPath $healthCheckerOutputsCurrent -OutputFilePath $scriptRoot

    #=========#-----------------------#
    #endregion# Run HealthChecker.ps1 #
    #=========#-----------------------#
}
catch {
    writeLog @writeLogParams -Message 'Ending script prematurely.' -ErrorRecord $_ -PassThru | Write-Warning
    throw
}
finally {
    writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand.Name) - End"
}

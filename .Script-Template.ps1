<#
    .Synopsis
    This is a template for a PowerShell script which reads its parameters from a separate (PSD1) file.

    .Description
    The script requires that a PSD1 file by the same name resides in the same folder.  Parameters are then included in
    the PSD1 file, which is imported by the script using Import-PowerShellDataFile.  This approach is convenient
    because it lets the script remain fully generic, and easy to run since the parameters are specified through the
    PSD1 file.  It also allows for maintain separate PSD1 files for the script to be run in different environments.

    Example script file / folder:

    <$PSScriptRoot>
        - Get-ADUsersWithNonExpiringPasswords.ps1
        - Get-ADUsersWithNonExpiringPasswords.Params.psd1
        - Get-ADUsersWithNonExpiringPasswords.Params.Dev.psd1

    .Parameter DevMode
    Specifies to import an alternative PSD1 file containing parameters (<ScriptName>.Params.Dev.psd1 instead of just\
    <ScriptName>.Params.psd1).  The intent is for testing in a different environment without needing to make changes in
    the script itself, rather just maintain separate PSD1 files for alternative parameter values.

    .Example

#>
#Requires -Version 5.1
#Requires -Modules ActiveDirectory
#Requires -Modules AzureAD
#Requires -PSEdition Desktop

using namespace System
using namespace System.Diagnostics
using namespace System.Management.Automation

[CmdletBinding(
    SupportsShouldProcess,
    ConfirmImpact = 'High'
)]
param(
    [switch]$DevMode
)

#======#-----------#
#region# Functions #
#======#-----------#
<#
    Functions region goes before Intialization & Variables region so that writeLog function can be available ASAP.
#>
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

#=========#-----------#
#endregion# Functions #
#=========#-----------#

try {

    #======#----------------------------#
    #region# Initialization & Variables #
    #======#----------------------------#

    # 1. writeLog splat and test writeLog:
    $Script:dtNow = [datetime]::Now
    $Script:writeLogParams = @{

        LogName     = "$($PSCmdlet.MyInvocation.MyCommand.Name -replace '\.ps1')"
        LogDateTime = $dtNow
        Folder      = "$($PSCmdlet.MyInvocation.MyCommand.Source -replace '\.ps1')_Logs"
        ErrorAction = 'Stop'
    }
    writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand.Name) - Start"
    writeLog @writeLogParams -Message "MyCommand: $($PSCmdlet.MyInvocation.Line)"

    # 2. Import static data from external PSD1:
    try {
        $Script:extParams = Import-PowerShellDataFile "$($PSCmdlet.MyInvocation.MyCommand.Source -replace '\.ps1').Params.psd1" -ErrorAction Stop
        writeLog @writeLogParams -Message "Imported external parameters: $($PSCmdlet.MyInvocation.MyCommand.Source -replace '\.ps1').Params.psd1"
    }
    catch {
        writeLog @writeLogParams -Message "Failed to import PowerShell data file '$($PSCmdlet.MyInvocation.MyCommand.Source -replace '\.ps1').Params.psd1'." -PassThru |
        Write-Warning
        throw
    }

    #=========#----------------------------#
    #endregion# Initialization & Variables #
    #=========#----------------------------#



    #======#----------------#
    #region# Data Retrieval #
    #======#----------------#

    #=========#----------------#
    #endregion# Data Retrieval #
    #=========#----------------#



    #======#------------#
    #region# Processing #
    #======#------------#

    #=========#------------#
    #endregion# Processing #
    #=========#------------#



    #======#---------------------#
    #region# Reporting & Wrap-Up #
    #======#---------------------#

    #=========#---------------------#
    #endregion# Reporting & Wrap-Up #
    #=========#---------------------#

}
catch {
    writeLog @writeLogParams -Message 'Script-ending problem encountered.' -ErrorRecord $_ -PassThru | Write-Warning
    throw
}
finally {
    writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand.Name) - End"
}

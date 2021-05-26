<#
    .Synopsis
    Get mobile devices and their statistics for specific or all on-premises mailboxes, for use with mailbox migration planning.

    .Example
    .\Get-MDevStats.ps1 -All | Export-Csv "MDevStats_$(Get-Date -Format 'yyyy-MM-dd').csv" -NTI -Encoding UTF8
#>
#Requires -Version 5.1
#Requires -PSEdition Desktop
[CmdletBinding()]
param(
    [Parameter(ParameterSetName='All', Mandatory)]
    [switch]$All,

    [Parameter(ParameterSetName='Identity', Mandatory)]
    [Alias('Guid','PrimarySmtpAddress','UserPrincipalName')]
    [string[]]$Identity
)
########----------------#
#region# Initialization #
########----------------#

$PSSessionsByComputerName = Get-PSSession | Group-Object -Property ComputerName
if (-not (Get-Command Get-MobileDeviceStatistics)) {

    "Command 'Get-MobileDeviceStatistics' is not available.  Make sure to run this script against Exchange 2016 or newer." | Write-Warning
    break
}
elseif ($PSSessionsByComputerName.Name -eq 'outlook.office365.com') {

    'EXO PSSession detected.  This script is intended for use with on-premises Exchange.' | Write-Warning
    break
}

$Start = [datetime]::Now
$Progress = @{

    Activity        = "Get-MDevStats.ps1 - Start time: $($Start)"
    PercentComplete = -1
}

###########----------------#
#endregion# Initialization #
###########----------------#



########----------------#
#region# Data Retrieval #
########----------------#

# 1. Get mailboxes to process:
$Progress['Status'] = 'Getting list of mailboxes to process...'
Write-Progress @Progress
try {
    $Mailboxes = @()
    Set-ADServerSettings -ViewEntireForest $true
    if ($PSCmdlet.ParameterSetName -eq 'All') {

        $Mailboxes += Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue -ErrorAction Stop
    }
    else {
        $Mailboxes += foreach ($id in $Identity) { Get-Mailbox -Identity $id -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction Stop }
    }
    if ($Mailboxes.Count -eq 0) {

        Write-Warning -Message 'Failed to find any mailboxes, yet no errors were encountered \_(;;)_/.'
        break
    }
}
catch { Write-Warning -Message 'Failed on Get-Mailbox step'; throw }


$Progress['Status'] = 'Processing mailboxes (getting devices/statistics)...'
$ProgressCounter = 0
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# 2. Get mobile devices/statistics for each mailbox:
foreach ($mbx in $Mailboxes) {

    $ProgressCounter++
    if ($Stopwatch.Elapsed.Milliseconds -ge 300) {

        $Progress['PercentComplete'] = (($ProgressCounter / $Mailboxes.Count) * 100)
        $Progress['CurrentOperation'] = "Preparing user/device custom object for $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))"
        Write-Progress @Progress

        $Stopwatch.Restart()
    }

    $MDevs = @()
    $MDevs += Get-MobileDeviceStatistics -Mailbox $mbx.Guid.Guid -ErrorAction SilentlyContinue |
    Select-Object -Property @{
        Name       = 'UserDisplayName'
        Expression = { $mbx.DisplayName }
    },
    LastSuccessSync,
    DeviceAccessState,
    DeviceAccessStateReason,
    Status,
    DeviceModel,
    DeviceImei,
    DevicePhoneNumber,
    DeviceOS,
    DeviceType,
    DeviceID,
    DeviceUserAgent,
    DeviceFriendlyName,
    DeviceMobileOperator,
    DevicePolicyApplied,
    DevicePolicyApplicationStatus,
    LastDeviceWipeRequestor,
    ClientVersion,
    NumberOfFoldersSynced,
    ClientType,
    Guid,
    @{
        Name = 'UserGuid'
        Expression = {$mbx.Guid.Guid}
    },
    @{
        Name = 'UserUPN'
        Expression = {$mbx.UserPrincipalName}
    }

    # Output this mailbox' devices/stats:
    $MDevs
}

###########----------------#
#endregion# Data Retrieval #
###########----------------#

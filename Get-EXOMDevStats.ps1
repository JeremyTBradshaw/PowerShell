<#
    .Synopsis
    Get mobile devices and some of their pertinent details/statistics for all mailboxes in the currently connected EXO tenant.

    .Example
    .\Get-EXOMDevStats.ps1 | Export-Csv "EXOMDevStats_$(Get-Date -Format 'yyyy-MM-dd').csv" -NTI -Encoding UTF8
#>
#Requires -Version 5.1
#Requires -Modules ExchangeOnlineManagement
[CmdletBinding()]
param()

#region Initialization
$PSSessionsByComputerName = Get-PSSession | Group-Object -Property ComputerName
if (-not (Get-Command Get-MobileDeviceStatistics)) {

    "Command 'Get-EXOMobileDeviceStatistics' is not available.  Make sure to run this script against Exchange Online using the EXOv2 module " +
    "(Install-Module ExchangeOnlineManagement -Scope CurrentUser; Connect-ExchangeOnline).  Exiting script." | Write-Warning
    break
}
elseif (-not ($PSSessionsByComputerName.Name -eq 'outlook.office365.com')) {

    Write-Warning -Message 'No EXO PSSession detected.  Connect first using Connect-ExchangeOnline.  Exiting script.'
    break
}

$Start = [datetime]::Now
$ProgressProps = @{

    Activity        = "Get-EXOMDevStats.ps1 - Start time: $($Start)"
    Status          = 'Working'
    PercentComplete = -1
}

try {
    Write-Progress @ProgressProps -CurrentOperation 'Get-EXOMailbox (ExchangeOnlineManagement module)'
    $EXOMailboxes = Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
}
catch {
    Write-Warning -Message "Failed on initial Get-EXOMailbox step.  Exiting script.  Error`n`n($_.Exception)"
    break
}
#endregion Initialization

#region Main loop
$ProgressCounter = 0
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($mbx in $EXOMailboxes) {

    $ProgressCounter++
    if ($Stopwatch.Elapsed.Milliseconds -ge 300) {

        $ProgressProps['PercentComplete'] = (($ProgressCounter / $EXOMailboxes.Count) * 100)
        $ProgressProps['CurrentOperation'] = "Preparing user/device combined objects for $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))"
        Write-Progress @ProgressProps

        $Stopwatch.Restart()
    }

    $MDevs = @()
    $MDevs += Get-EXOMobileDeviceStatistics -Mailbox $mbx.Guid.Guid -ErrorAction SilentlyContinue |
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
        Name = 'UserUPN'
        Expression = {$mbx.UserPrincipalName}
    }

    $MDevs
}
#region Main loop

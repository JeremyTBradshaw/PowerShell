<#
    .Synopsis
    Get mobile devices and their statistics for all on-premises mailboxes, for use with mailbox migration planning.

    .Example
    .\Get-MDevStats.ps1 | Export-Csv "MDevStats_$(Get-Date -Format 'yyyy-MM-dd').csv" -NTI -Encoding UTF8
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

#region Initialization
$PSSessionsByComputerName = Get-PSSession | Group-Object -Property ComputerName
if (-not (Get-Command Get-MobileDeviceStatistics)) {

    "Command 'Get-MobileDeviceStatistics' is not available.  Make sure to run this script against Exchange 2016 or newer." |
    Write-Warning
    break
}
elseif ($PSSessionsByComputerName.Name -eq 'outlook.office365.com') {

    'EXO PSSession detected.  This script is intended for use with on-premises Exchange.  Exiting script.' |
    Write-Warning
    break
}

$Start = [datetime]::Now
$ProgressProps = @{

    Activity        = "Get-MDevStats.ps1 - Start time: $($Start)"
    Status          = 'Working'
    PercentComplete = -1
}

try {
    Write-Progress @ProgressProps -CurrentOperation 'Get-Mailbox (on-premises mailboxes)'
    $LocalMailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
    Where-Object { $_.RecipientTypeDetails -ne 'DiscoveryMailbox' -and $_.RecipientTypeDetails -ne 'ArbitrationMailbox' }
}
catch {
    Write-Warning -Message "Failed on initial Get-Mailbox step.  Exiting script.  Error`n`n($_.Exception)"
    break
}
#endregion Initialization

#region Main loop
$ProgressCounter = 0
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($mbx in $LocalMailboxes) {

    $ProgressCounter++
    if ($Stopwatch.Elapsed.Milliseconds -ge 300) {

        $ProgressProps['PercentComplete'] = (($ProgressCounter / $LocalMailboxes.Count) * 100)
        $ProgressProps['CurrentOperation'] = "Preparing user/device custom objects for $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))"
        Write-Progress @ProgressProps

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
        Name = 'UserUPN'
        Expression = {$mbx.UserPrincipalName}
    }

    $MDevs
}
#endregion Main loop

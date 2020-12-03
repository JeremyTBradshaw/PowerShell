<#
    .Synopsis
    Get mobile devices and their statistics for all on-premises mailboxes, for use with mailbox migration planning.
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

begin {
    # Use Exchange 2016 Management Shell or remote PowerShell from Windows 10 to an Exchange 2016 server.
    $PSSessionsByComputerName = Get-PSSession | Group-Object -Property ComputerName
    if (-not (Get-Command Get-MobileDeviceStatistics)) {
        
        Write-Warning -Message "Command 'Get-MobileDeviceStatistics' is not available.  Make sure to run this script against Exchange 2016 or newer."
        break
    }
    elseif ($PSSessionsByComputerName.Name -eq 'outlook.office365.com') {

        Write-Warning -Message 'EXO PSSession detected.  This script is intended for use with on-premises Exchange.  Exiting script.'
    }

    $Start = [datetime]::Now
    $ProgressProps = @{

        Activity        = "Get-MailboxReport - Start time: $($Start)"
        Status          = 'Working'
        PercentComplete = -1
    }

    try {
        Write-Progress @ProgressProps -CurrentOperation 'Get-Mailbox (on-premises mailboxes)'
        $LocalMailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
        Where-Object { $_.RecipientTypeDetails -ne 'DiscoveryMailbox' -and $_.RecipientTypeDetails -ne 'ArbitrationMailbox' }
    }
    catch {
        Write-Warning -Message "Failed on initial data collection step.  Exiting script.  Error`n`n($_.Exception)"
        break
    }
}

process {
    # Start the main loop:

    $ProgressCounter = 0
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($lm in $LocalMailboxes) {

        $ProgressCounter++
        if ($Stopwatch.Elapsed.Milliseconds -ge 300) {

            $ProgressProps['PercentComplete'] = (($ProgressCounter / $LocalMailboxes.Count) * 100)
            $ProgressProps['CurrentOperation'] = "Preparing user/device custom objects for $($lm.DisplayName) ($($lm.PrimarySmtpAddress))"
            Write-Progress @ProgressProps

            $Stopwatch.Restart()
        }

        $MDevs = @()
        $MDevs += Get-MobileDeviceStatistics -Mailbox $lm.Guid.Guid -ErrorAction SilentlyContinue |
        Select-Object -Property @{
            Name       = 'UserDisplayName'
            Expression = { $lm.DisplayName }
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
        @{
            Name       = 'UserGuid'
            Expression = { $lm.Guid.Guid }
        },
        Guid

        if ($MDevs.Count -ge 1) {

            # Output the combined user/device objects:
            $MDevs
        }
    }
}

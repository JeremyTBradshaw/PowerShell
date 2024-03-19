<#
    .Synopsis
    Mailbox report merging output properties from multiple cmdlets into one common output object.

    .Description
    Pulls together some essential details for all on-premises-sourced mailboxes, including migrated (with fewer
    details), and excluding Arbitration (i.e. System) and Discovery mailboxes.

    Sources include:
    - Get-Recipient
    - Get-Mailbox
    - Get-User
    - Get-ADUser
    - Get-MailboxStatistics
    - Get-MobileDeviceStatistics

    Requires an open PSSession to an on-premises Exchange server(2016+).
#>
#Requires -Version 5.1
#Requires -Modules ActiveDirectory
[CmdletBinding()]
param(
    [ValidatePattern('^\w+([-+.'']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$')]
    [string]$AadUPN
)

begin {
    # Use Exchange 2016 Management Shell or remote PowerShell from Windows 10 to an Exchange 2016 server.
    $PSSessionsByComputerName = Get-PSSession | Group-Object -Property ComputerName
    if (-not (Get-Command Get-MobileDeviceStatistics)) {

        Write-Warning -Message "Command 'Get-MobileDeviceStatistics' is not available.  Make sure to run this script against Exchange 2016 or newer."
        break
    }
    elseif ($PSSessionsByComputerName.Name -eq 'outlook.office365.com') {

        Write-Warning -Message 'EXO PSSession detected.  This script is intended for use with on-premises Exchange (and from an AD-joined computer).  Exiting script.'
        break
    }

    # Use a domain-joined computer.
    if ((Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain -eq $false) {

        Write-Warning -Message 'This script must be run from a domain-joined computer.  Exiting script.'
        break
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

        Write-Progress @ProgressProps -CurrentOperation 'Get-Recipient -ResultSize Unlimited (mailboxes local/remote)'
        $MailboxRecipients = Get-Recipient -ResultSize Unlimited -ErrorAction Stop |
        Where-Object { $_.RecipientTypeDetails -match '(^(User)|(Shared)|(Room)|(Equipment)|(Remote).*Mailbox$)' }

        Write-Progress @ProgressProps -CurrentOperation 'Get-User -ResultSize Unlimited (mailboxes local/remote)'
        $MailboxUsers = Get-User -ResultSize Unlimited -ErrorAction Stop |
        Where-Object { $_.RecipientTypeDetails -match '(^(User)|(Shared)|(Room)|(Equipment)|(Remote).*Mailbox$)' }

        Write-Progress @ProgressProps -CurrentOperation 'Get-ADUser (mailboxes local/remote)'
        $ADMailboxUsers = Get-ADUser -Filter "msExchMailboxGuid -like '*'" -Properties msExchMailboxGuid, LastLogonDate -ErrorAction Stop
    }
    catch {
        Write-Warning -Message "Failed on initial data collection step.  Exiting script.  Error`n`n($_.Exception)"
        break
    }
}

process {

    # Prepare lookup tables:
    Write-Progress @ProgressProps -CurrentOperation 'Preparing lookup tables'

    $lmHT = @{}
    foreach ($lm in $LocalMailboxes) {

        $lmHT[$lm.Guid.Guid] = $lm
    }

    $muHT = @{}
    foreach ($mu in $MailboxUsers) {

        $muHT[$mu.Guid.Guid] = $mu
    }

    $admuHT = @{}
    foreach ($admu in $ADMailboxUsers) {

        $admuHT[$admu.ObjectGuid.Guid] = $admu
    }

    # Start the main loop:

    $ProgressCounter = 0
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($mr in $MailboxRecipients) {

        $ProgressCounter++
        if ($Stopwatch.Elapsed.Milliseconds -ge 300) {

            $ProgressProps['PercentComplete'] = (($ProgressCounter / $MailboxRecipients.Count) * 100)
            $ProgressProps['CurrentOperation'] = "Preparing common output object for $($mr.DisplayName) ($($mr.RecipientTypeDetails))"
            Write-Progress @ProgressProps

            $Stopwatch.Restart()
        }

        # Start building the commonized object for this user, using properties available from the initial Get- cmdlets earlier:

        $mrHT = [ordered]@{

            DisplayName                   = $mr.DisplayName
            AccountEnabled                = $admuHT[$mr.Guid.Guid].Enabled
            ADLastLogonDate               = if ($admuHT[$mr.Guid.Guid].LastLogonDate) { $admuHT[$mr.Guid.Guid].LastLogonDate.ToString('yyyy-MM-dd') } else { '' }
            FirstName                     = $mr.FirstName
            Initials                      = $muHT[$mr.Guid.Guid].Initials
            LastName                      = $mr.LastName
            MobilePhone                   = $muHT[$mr.Guid.Guid].MobilePhone
            Phone                         = $mr.Phone
            PrimarySmtpAddress            = $mr.PrimarySmtpAddress
            UserPrincipalName             = $muHT[$mr.Guid.Guid].UserPrincipalName
            PSmtpUpnMatch                 = if ($mr.PrimarySmtpAddress -eq $muHT[$mr.Guid.Guid].UserPrincipalName) { $true } else { $false }
            RecipientTypeDetails          = $mr.RecipientTypeDetails
            EmailAddressPoliciesEnabled   = $mr.EmailAddressPolicyEnabled
            RemoteRoutingAddress          = ''
            HiddenFromAddressListsEnabled = $mr.HiddenFromAddressListsEnabled
            Database                      = $mr.Database
            MailboxSizeGB                 = ''
            MailboxItemCount              = ''
            NewestSentItem                = ''
            ArchiveState                  = $mr.ArchiveState
            ArchiveDatabase               = $mr.ArchiveDatabase
            ArchiveSizeGB                 = ''
            ArchiveItemCount              = ''
            DevicesCount                  = ''
            MostRecentDeviceSuccessSync   = ''
            MostRecentDeviceType          = ''
            MostRecentDeviceId            = ''
            Office                        = $mr.Office
            ManagerId                     = $mr.Manager
            Title                         = $mr.Title
            Department                    = $mr.Department
            Company                       = $mr.Company
            Guid                          = $mr.Guid
            ExchangeGuid                  = $mr.ExchangeGuid
            ArchiveGuid                   = $mr.ArchiveGuid
            OrganizationalUnit            = $mr.OrganizationalUnit
            CanonicalName                 = $mr.Identity
            EmailAddresses                = $mr.EmailAddresses -join ' | '
        }

        $RemoteRoutingAddress = @()
        $RemoteRoutingAddress += $mr.EmailAddresses | Where-Object { $_ -like 'smtp:*@*.mail.onmicrosoft.com' }
        $mrHT['RemoteRoutingAddress'] = $RemoteRoutingAddress[0] -replace 'smtp:'

        if ($mr.RecipientTypeDetails -notmatch '(^Remote.*)') {

            # Only processing local mailboxes (we're not processing remote/migrated mailboxes).

            $MStats = $null
            $MStats = $mr | Get-MailboxStatistics -ErrorAction Continue

            $MailboxSizeGB = try {
                [math]::Round( ([decimal]($MStats.TotalItemSize -replace '(.*\()|(,)|(\s.*)') + [decimal]($MStats.TotalDeletedItemSize -replace '(.*\()|(,)|(\s.*)')) / 1GB, 2 )
            }
            catch { '' }

            $MFSIStats = Get-MailboxFolderStatistics -Identity $mr.Guid.Guid -FolderScope SentItems -IncludeOldestAndNewestItems |
            Sort-Object { $_.Identity -match '(Sent Items$)' }

            $MStatsArchive = $null
            if ($lmHT[$mr.Guid.Guid].ArchiveState -like 'Local') {

                $MStatsArchive = $mr | Get-MailboxStatistics -Archive -ErrorAction Continue
            }
            if ($MStatsArchive) {

                $ArchiveSizeGB = try {
                    [math]::Round( ([decimal]($MStatsArchive.TotalItemSize -replace '(.*\()|(,)|(\s.*)') + [decimal]($MStatsArchive.TotalDeletedItemSize -replace '(.*\()|(,)|(\s.*)')) / 1GB, 2 )
                }
                catch { '' }

                $ArchiveItemCount = $MStatsArchive.ItemCount
            }
            else {
                $ArchiveSizeGB = ''
                $ArchiveItemCount = ''
            }

            $MDevs = @()
            $MDevs += Get-MobileDeviceStatistics -Mailbox $mr.Guid.Guid -ErrorAction Continue

            $RecentMDev = $null
            $RecentMDev = $MDevs | Sort-Object -Property LastSuccessSync | Select-Object -Last 1

            # Add the on-premises mailbox-related properties to the output object:

            $mrHT['MailboxSizeGB'] = $MailboxSizeGB
            $mrHT['MailboxItemCount'] = $MStats.ItemCount
            $mrHT['NewestSentItem'] = $MFSIStats.NewestItemReceivedDate
            $mrHT['ArchiveState'] = $lmHT[$mr.Guid.Guid].ArchiveState
            $mrHT['ArchiveDatabase'] = $lmHT[$mr.Guid.Guid].ArchiveDatabase
            $mrHT['ArchiveSizeGB'] = $ArchiveSizeGB
            $mrHT['ArchiveItemCount'] = $ArchiveItemCount
            $mrHT['DevicesCount'] = $MDevs.Count
            $mrHT['MostRecentDeviceSuccessSync'] = $RecentMDev.LastSuccessSync
            $mrHT['MostRecentDeviceType'] = $RecentMDev.DeviceType
            $mrHT['MostRecentDeviceId'] = $RecentMDev.DeviceId
        }

        Write-Debug -Message 'Stop to inspect $mrHT, $mr, $admuHT[$mr.Guid.Guid], $muHT[$mr.Guid.Guid].'

        # Output the commonized object:
        [PSCustomObject]$mrHT
    }
}

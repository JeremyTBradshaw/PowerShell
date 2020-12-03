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
    - Get-AzureADUser
    - Get-MailboxStatistics
    - Get-MobileDeviceStatistics

    Requires an open PSSession to an on-premises Exchange server(2016+).

    .Parameter AadUPN
    Supply your UserPrincipalName for use with Connect-AzureAD to re-use an existing refresh token (if one exists).
#>
#Requires -Version 5.1
#Requires -Modules ActiveDirectory
#Requires -Modules @{ ModuleName = 'AzureAD'; Guid = 'd60c0004-962d-4dfb-8d28-5707572ffd00'; ModuleVersion = '2.0.2.118'}

[CmdletBinding()]
param(
    [ValidatePattern('^(\w+@)(\w+\.)+\w+$')]
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

        Write-Progress @ProgressProps -CurrentOperation 'Get-AzureADUser -All'
        if ($PSBoundParameters.ContainsKey('AadUPN')) {

            [void](Connect-AzureAD -AccountId $AadUPN -ErrorAction Stop)
        }
        else {
            [void](Connect-AzureAD -ErrorAction Stop)
        }
        $AADUsers = Get-AzureADUser -All $true -ErrorAction Stop
    }
    catch {
        Write-Warning -Message "Failed on initial data collection step.  Exiting script.  Error`n`n($_.Exception)"
        break
    }

    $skuIdHT = @{

        # Build this manually using Get-AzureADSubscribedSku and the Azure AD Portal (Licenses > All Products).

        'f30db892-07e9-47e9-837c-80727f46fd3d' = 'Microsoft Power Automate Free'
        'c5928f49-12ba-48f7-ada3-0d743a3601d5' = 'Visio Plan 2'
        '53818b1b-4a27-454b-8896-0dba576410e6' = 'Project Plan 3'
        '6470687e-a428-4b7a-bef2-8a291ad947c9' = 'Microsoft Store for Business'
        'b05e124f-c7cc-45a0-a6aa-8cf78c946968' = 'Enterprise Mobility + Security E5'
        'a403ebcc-fae0-4ca2-8c8c-7a907fd6c235' = 'Power BI (free)'
        '2b9c8e7c-319c-43a2-a2a0-48c5c6161de7' = 'Azure Active Directory Basic'
        '09015f9f-377f-4538-bbb5-f75ceb09358a' = 'Project Plan 5'
        '4a51bf65-409c-4a91-b845-1121b571cc9d' = 'Power Automate per user plan'
        '6fd2c87f-b296-42f0-b197-1e91e994b900' = 'Office 365 E3'
        '776df282-9fc0-4862-99e2-70e561b9909e' = 'Project Online Essentials'
        '8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b' = 'Rights Management Adhoc'
        '710779e8-3d4a-4c88-adb9-386c958d1fdf' = 'Microsoft Teams Exploratory'
    }
}

process {

    # Prepare lookup tables:
    Write-Progress @ProgressProps -CurrentOperation 'Preparing lookup tables'

    $aadUHT = @{}
    # Only caring about synced AD users, not caring about cloud-only users:
    foreach ($aadU in ($AADUsers | Where-Object { $_.ImmutableId })) {

        $aadUHT["$(([Guid]([Convert]::FromBase64String($aadU.ImmutableId))).Guid)"] = $aadU
    }

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
            O365_E3                       = ''
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
            AssignedLicenses              = ''
            EmailAddresses                = $mr.EmailAddresses -join ' | '
        }

        if ($aadUHT[$mr.Guid.Guid].AssignedLicenses) {

            $mrHT['AssignedLicenses'] = $skuIdHT[$aadUHT[$mr.Guid.Guid].AssignedLicenses.skuId] -join ', '
        }
        if ($mrHT['AssignedLicenses'] -match '(Office 365 E3)') {

            $mrHT['O365_E3'] = $true
        }
        else { $mrHT['O365_E3'] = $false }

        $RemoteRoutingAddress = @()
        $RemoteRoutingAddress += $mr.EmailAddresses | Where-Object { $_ -like 'smtp:*@*.mail.onmicrosoft.com' }
        $mrHT['RemoteRoutingAddress'] = $RemoteRoutingAddress[0] -replace 'smtp:'

        if ($mr.RecipientTypeDetails -notmatch '(^Remote.*)') {

            # Working with a local mailbox.

            $MStats = $null
            $MStats = $mr | Get-MailboxStatistics

            $MStatsArchive = $null
            if ($lmHT[$mr.Guid.Guid].ArchiveState -like 'Local') {
                
                $MStatsArchive = $mr | Get-MailboxStatistics -Archive
            }

            $MDevs = @()
            $MDevs += Get-MobileDeviceStatistics -Mailbox $mr.Guid.Guid
            
            $RecentMDev = $null
            $RecentMDev = $MDevs | Sort-Object -Property LastSuccessSync | Select-Object -Last 1

            $MailboxSizeGB = try {
                [math]::Round( ([decimal]($MStats.TotalItemSize -replace '(.*\()|(,)|(\s.*)') + [decimal]($MStats.TotalDeletedItemSize -replace '(.*\()|(,)|(\s.*)')) / 1GB, 2 )
            }
            catch { '' }

            if ($MStatsArchive) {

                $ArchiveSizeGB = try {
                    [math]::Round( ([decimal]($MStatsArchive.TotalItemSize -replace '(.*\()|(,)|(\s.*)') + [decimal]($MStatsArchive.TotalDeletedItemSize -replace '(.*\()|(,)|(\s.*)')) / 1GB, 2 )
                }
                catch { '' }

                $ArchiveItemCount = $MStatsArchive.ItemCount
            }
            else { $ArchiveSizeGB = ''; $ArchiveItemCount = '' }

            # Add the on-premises mailbox-related properties to the output object:

            $mrHT['MailboxSizeGB'] = $MailboxSizeGB
            $mrHT['MailboxItemCount'] = $MStats.ItemCount
            $mrHT['ArchiveState'] = $lmHT[$mr.Guid.Guid].ArchiveState
            $mrHT['ArchiveDatabase'] = $lmHT[$mr.Guid.Guid].ArchiveDatabase
            $mrHT['ArchiveSizeGB'] = $ArchiveSizeGB
            $mrHT['ArchiveItemCount'] = $ArchiveItemCount
            $mrHT['DevicesCount'] = $MDevs.Count
            $mrHT['MostRecentDeviceSuccessSync'] = $RecentMDev.LastSuccessSync
            $mrHT['MostRecentDeviceType'] = $RecentMDev.DeviceType
            $mrHT['MostRecentDeviceId'] = $RecentMDev.DeviceId
        }

        Write-Debug -Message 'Stop to inspect $mrHT, $mr, $admuHT[$mr.Guid.Guid], $muHT[$mr.Guid.Guid], $aadUHT[$mr.Guid.Guid], $skuIdHT[$aadUHT[$mr.Guid.Guid].AssignedLicenses.skuId].'

        # Output the commonized object:
        [PSCustomObject]$mrHT
    }
}

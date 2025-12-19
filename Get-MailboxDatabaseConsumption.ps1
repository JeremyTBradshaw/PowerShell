<#
    .Synopsis
    Get database size, whitespace, should-be-whitespace (in stubbing environments),
    mailbox count, and consumption breakdown by:
      - Active mailboxes
      - Disconnected mailboxes (disabled, soft deleted)

    .Parameter Database
    Target one of more specific databases, rather than all databases, which is
    the default behavior.

    .Parameter IncludePreExchange2013
    Set this to $false to override the default behavior, which includes
    pre-Exchange 2013 databases.
#>
#Requires -Version 4
[CmdletBinding()]
param(
    [string[]]$Database,
    [bool]$IncludePreExchange2013 = $true
)

Write-Verbose -Message "Determining the connected Exchange environment."

$ExPSSession = @()
$ExPSSession += Get-PSSession |
Where-Object {
    $_.ConfigurationName -eq 'Microsoft.Exchange' -and
    $_.State -eq 'Opened'
}

if ($ExPSSession.Count -eq 1) {
    $Exchange = $null

    # Check if we're in Exchange Online or On-Premises.
    switch ($ExPSSession.ComputerName) {

        outlook.office365.com {
            $Exchange = 'Exchange Online'
        }

        default {
            $Exchange = 'Exchange On-Premises'

            # Set scope to entire forest (important for multi-domain forests).
            Set-ADServerSettings -ViewEntireForest:$true

            # Determine if the connected Exchange server's version is 2010 or 2013 and newer.
            $ExOnPSrv = Get-ExchangeServer -Identity "$($ExPSSession.ComputerName)"

            switch ($ExOnPSrv.AdminDisplayVersion) {

                { $_ -match 'Version 14' } { $LegacyExchange = $true }
                { $_ -match 'Version 15' } { $LegacyExchange = $false }
                default {
                    throw "Unable to determine connect Exchange On-Premises server version.  Only Exchange 2010 and newer are supported by this script."
                }
            }
        }
    }
    Write-Verbose -Message "Connected environment is $($Exchange)."
}
else {
    Write-Warning -Message "Requires a single** active (State: Opened) remote session to an Exchange server."
    break
}

if ($Exchange -eq 'Exchange Online') {
    Write-Warning -Message "The active session is with Exchange Online, which is not supported by this script."
    break
}


if ($PSBoundParameters.ContainsKey('Database')) {

    $Databases = @()
    $Databases +=
    if ($LegacyExchange) { foreach ($DB in $Database) { Get-MailboxDatabase -Identity $DB -Status } }
    else { foreach ($DB in $Database) { Get-MailboxDatabase -Identity $DB -Status -IncludePreExchange2013:$IncludePreExchange2013 } }
}
else {
    $Databases = @()
    $Databases +=
    if ($LegacyExchange) { Get-MailboxDatabase -Status }
    else { Get-MailboxDatabase -Status -IncludePreExchange2013:$IncludePreExchange2013 }
}

if ($Databases.Count -ge 1) {

    $Databases = $Databases | Where-Object { ($_.Recovery -eq $false) -and ($_.Mounted -eq $true) }

    foreach ($DB in $Databases) {

        $MailboxStatistics = Get-MailboxStatistics -Database $DB.Name

        $ActiveMailboxes = @()
        $ActiveMailboxes += $MailboxStatistics | Where-Object { $null -eq $_.DisconnectReason }

        $ArchiveMailboxes = @()
        $ArchiveMailboxes += $MailboxStatistics | Where-Object { $_.IsArchiveMailbox -eq $true }

        $DisconnectedMailboxes = @()
        $DisconnectedMailboxes += $MailboxStatistics | Where-Object { $null -ne $_.DisconnectReason }

        $DisabledMailboxes = @()
        $DisabledMailboxes += $DisconnectedMailboxes | Where-Object { $_.DisconnectReason -like 'Disabled' }

        $SoftDeletedMailboxes = @()
        $SoftDeletedMailboxes += $DisconnectedMailboxes | Where-Object { $_.DisconnectReason -like 'SoftDeleted' }

        $DBConsumption = [PSCustomObject]@{

            Database              = $DB.Name
            DBSize_GB             = [math]::Round(($DB.DatabaseSize -replace '.*\s\(' -replace ',' -replace '\sb.*') / 1024 / 1024 / 1024, 2)
            Whitespace_GB         = [math]::Round(($DB.AvailableNewMailboxSpace -replace '.*\s\(' -replace ',' -replace '\sb.*') / 1024 / 1024 / 1024, 2)

            Mbx_Count             = $MailboxStatistics.Count
            Mbx_GB                = ($MailboxStatistics |
                Select-Object @{
                    Name       = 'MSizeGB'
                    Expression = { (
                            [math]::Round((
                                    [decimal]($_.TotalItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*') +
                                    [decimal]($_.TotalDeletedItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*')
                                ) / 1024 / 1024 / 1024, 2)
                        ) }
                } | Measure-Object -Property MSizeGB -Sum).Sum

            ActiveMbx_Count       = $ActiveMailboxes.Count
            ActiveMbx_GB          = ($ActiveMailboxes |
                Select-Object @{
                    Name       = 'MSizeGB'
                    Expression = { (
                            [math]::Round((
                                    [decimal]($_.TotalItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*') +
                                    [decimal]($_.TotalDeletedItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*')
                                ) / 1024 / 1024 / 1024, 2)
                        ) }
                } | Measure-Object -Property MSizeGB -Sum).Sum

            ArchiveMbx_Count      = $ArchiveMailboxes.Count
            ArchiveMbx_GB         = ($ArchiveMailboxes |
                Select-Object @{
                    Name       = 'MSizeGB'
                    Expression = { (
                            [math]::Round((
                                    [decimal]($_.TotalItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*') +
                                    [decimal]($_.TotalDeletedItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*')
                                ) / 1024 / 1024 / 1024, 2)
                        ) }
                } | Measure-Object -Property MSizeGB -Sum).Sum
    
            DisconnectedMbx_Count = $DisconnectedMailboxes.Count
            DisconnectedMbx_GB    = ($DisconnectedMailboxes |
                Select-Object @{
                    Name       = 'MSizeGB'
                    Expression = { (
                            [math]::Round((
                                    [decimal]($_.TotalItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*') +
                                    [decimal]($_.TotalDeletedItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*')
                                ) / 1024 / 1024 / 1024, 2)
                        ) }
                } | Measure-Object -Property MSizeGB -Sum).Sum

            DisabledMbx_Count     = $DisabledMailboxes.Count
            DisabledMbx_GB        = ($DisabledMailboxes |
                Select-Object @{
                    Name       = 'MSizeGB'
                    Expression = { (
                            [math]::Round((
                                    [decimal]($_.TotalItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*') +
                                    [decimal]($_.TotalDeletedItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*')
                                ) / 1024 / 1024 / 1024, 2)
                        ) }
                } | Measure-Object -Property MSizeGB -Sum).Sum

            SoftDeletedMbx_Count  = $SoftDeletedMailboxes.Count
            SoftDeletedMbx_GB     = ($SoftDeletedMailboxes |
                Select-Object @{
                    Name       = 'MSizeGB'
                    Expression = { (
                            [math]::Round((
                                    [decimal]($_.TotalItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*') +
                                    [decimal]($_.TotalDeletedItemSize -replace '.*\s\(' -replace ',' -replace '\sb.*')
                                ) / 1024 / 1024 / 1024, 2)
                        ) }
                } | Measure-Object -Property MSizeGB -Sum).Sum
        }
        Write-Output -InputObject $DBConsumption
    }
}

<#
    .Synopsis
    Export EXO mailboxes to PST.  Enables folder selection and preserves folder structure.

    .Description
    Intended for interactive use by EXO/SCC administrators with enough access to use the following Cmdlets from the
    ExchangeOnlineManagement PS module (https://www.powershellgallery.com/packages/exchangeonlinemanagement):

        - Exchange Online: Get-EXOMailbox, Get-EXOMailboxStatistics
        - Security/Compliance Center: *-ComplianceSearch, *-ComplianceSearchAction

    .Parameter AdminUPN
    Specifies the UserPrincipalName to be used with Connect-ExchangeOnline and Connect-IPPSSession.  Supplying the UPN
    avoids re-prompting for credentials after the first successful time.

    .Parameter MailboxPSmtp
    Specifies the PrimarySmtpAddress of the mailbox to export.

    .Parameter MailboxSelection
    Provides the choice of 'Primary', 'Archive', or 'Both' (default).  The Out-Gridview folder picker UI will show
    which mailbox the folders reside in, for easy selection of folders from either mailbox location.

    .Parameter InactiveMailbox
    Indicates the mailbox is an Inactive Mailbox (which needs to be indicated to New-ComplianceSearch as
    -ExchangeLocation ".<PrimarySmtpAddress>").

    .Parameter SearchNameOverride
    By default the Content/Compliance Search will be named:
    "Export-EXOMailbox.ps1_<MailboxPSmtp>_<Primary|Archive|Primary+Archive>_<(optionally)InactiveMBX>"

    If an existing search by the same name is found, the script will attempt to delete it before creating a new one.

    .Example
    .\Export-EXOMailbox.ps1 -AdminUPN admin1@contoso.onmicrosoft.com -MailboxPSmtp user1@contoso.com

    .Example
    .\Export-EXOMailbox.ps1 -AdminUPN admin1@contoso.onmicrosoft.com -MailboxPSmtp user1@contoso.com -SearchNameOverride User1Export_2022-01-26

    .Example
    .\Export-EXOMailbox.ps1 -AdminUPN admin1@contoso.onmicrosoft.com -MailboxPSmtp user1@contoso.com -MailboxSelection Primary

    .Example
    .\Export-EXOMailbox.ps1 -AdminUPN admin1@contoso.onmicrosoft.com -MailboxPSmtp user1@contoso.com -MailboxSelection Archive -InactiveMailbox

    .Outputs
    There is some information output and warnings in some scenarios.  For mailbox selection (when there are multiple
    mailboxes found) and for the folder selections, Out-Gridview is used to allow for easy GUI interaction.

    .Notes
    This script was initially created via Save-As from a Microsoft-created script (Export-Folder.ps1). Most of the code
    has been replaced, and a tiny bit remains.  Credits are included where due.

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Export-EXOMailbox.ps1
#>
#Requires -Modules @{ ModuleName = 'ExchangeOnlineManagement'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'; ModuleVersion = '3.9.0' }
[CmdletBinding(
    SupportsShouldProcess,
    ConfirmImpact = 'High'
)]
param (
    [Parameter(Mandatory)]
    [string]$AdminUPN,
    [Parameter(Mandatory)]
    [string]$MailboxPSmtp,
    [ValidateSet('Primary', 'Archive', 'Both')]
    [string]$MailboxSelection = 'Both',
    [switch]$InactiveMailbox,
    [string]$SearchNameOverride
)

if ($WhatIfPreference.IsPresent) {
    "Microsoft does not support -WhatIf for the Security and Compliance Center cmdlets.  " +
    "ShouldProcess support is included in this script to avoid accidentally deleting any compliance searches.  " +
    "Accordingly, -Confirm is still supported, but -WhatIf is not.  Exiting script." | Write-Warning
    break
}
try {
    $progress = @{
        Activity        = "Export-EXOMailbox.ps1 - Start time: $([datetime]::Now)"
        PercentComplete = -1
    }

    ########---------------------------------#
    #region# Find mailbox and select folders #
    ########---------------------------------#

    Write-Progress @progress -Status "Connect-ExchangeOnline"
    Disconnect-ExchangeOnline -Confirm:$false
    Connect-ExchangeOnline -UserPrincipalName $AdminUPN -CommandName Get-Mailbox -ShowBanner:$false -DisableWAM -ErrorAction Stop #<--: CommandName reduces time to connect, auto-includes all *-EXO*** modern Cmdlets.

    Write-Progress @progress -Status 'Get-EXOMailbox'
    $MailboxLookup = @(Get-EXOMailbox $MailboxPSmtp -Properties ExchangeGuid, ArchiveGuid, isInactiveMailbox -InactiveMailboxOnly:$InactiveMailbox -ErrorAction Stop)
    $SelectedMailbox = if ($MailboxLookup.Count -gt 1) {

        $MailboxPicker = @()
        foreach ($mbx in $MailboxLookup) {
            $MailboxPicker += [PSCustomObject]@{
                DisplayName        = $mbx.DisplayName
                PrimarySmtpAddress = $mbx.PrimarySmtpAddress
                WhenCreated        = $mbx.WhenCreated
                WhenMailboxCreated = $mbx.WhenMailboxCreated
                isInactiveMailbox  = $mbx.isInactiveMailbox
                Guid               = $mbx.Guid
            }
        }
        Write-Progress @progress -Status 'Waiting for mailbox selection'
        $MailboxPicker | Out-GridView -OutputMode Single -Title 'Select mailbox to search and export:'
    }
    else { $MailboxLookup[0] }

    $ht_Mailbox = @{
        "$($SelectedMailbox.ExchangeGuid.ToString())" = 'Primary Mailbox'
        "$($SelectedMailbox.ArchiveGuid.ToString())"  = 'Archive Mailbox'
    }

    $FolderStatistics = @(
        if (@('Primary', 'Both') -contains $MailboxSelection) {
            Write-Progress @progress -Status 'Get-EXOMailboxFolderStatistics (primary mailbox)'
            Get-EXOMailboxFolderStatistics $SelectedMailbox.ExchangeGuid.ToString() -IncludeSoftDeletedRecipients:$InactiveMailbox -ErrorAction Stop
        }
        if (@('Archive', 'Both') -contains $MailboxSelection) {
            if ($SelectedMailbox.ArchiveGuid.ToString() -like '000*') {
                Write-Warning -Message "Mailbox $($MailboxPSmtp) is not enabled with an Archive mailbox."
            }
            else {
                Write-Progress @progress -Status 'Get-EXOMailboxFolderStatistics (archive mailbox)'
                Get-EXOMailboxFolderStatistics $SelectedMailbox.ArchiveGuid.ToString() -IncludeSoftDeletedRecipients:$InactiveMailbox -Archive -ErrorAction Stop
            }
        }
    )

    $FolderPicker = @()
    $fCounter = 0
    foreach ($fStat in $FolderStatistics) {
        $fCounter++
        $progress['PercentComplete'] = ($fCounter / $FolderStatistics.Count) * 100
        Write-Progress @progress -Status 'Parsing/process folder statistics'

        # Borrowed code (start)
        $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
        $nibbler = $encoding.GetBytes("0123456789ABCDEF")
        $folderIdBytes = [Convert]::FromBase64String($fStat.FolderId)
        $indexIdBytes = New-Object byte[] 48
        $indexIdIdx = 0
        $folderIdBytes | Select-Object -skip 23 -First 24 | ForEach-Object {
            $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]
            $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF]
        }
        # Borrowed code (end)

        $FolderPicker += [PSCustomObject]@{
            Mailbox       = $ht_Mailbox["$($fStat.ContentMailboxGuid.ToString())"]
            FolderPath    = $fStat.FolderPath
            Foldersize    = $fStat.FolderSize
            ItemsInFolder = $fStat.ItemsInFolder
            FolderQuery   = "folderid:$($encoding.GetString($indexIdBytes))" #<--: Borrowed code.
        }
    }

    $progress['PercentComplete'] = -1
    Write-Progress @progress -Status 'Waiting for folder selections'
    $SelectedFolders = @($FolderPicker | Out-GridView -OutputMode Multiple -Title 'Select folders to include in the export:')
    if ($SelectedFolders.Count -lt 1) {
        Write-Warning -Message 'No folders were selected.  Exiting script prematurely.'
        break
    }

    ###########---------------------------------#
    #endregion# Find mailbox and select folders #
    ###########---------------------------------#



    ########----------------------------------#
    #region# Create and run compliance search #
    ########----------------------------------#

    Write-Progress @progress -Status 'Connect-IPPSSession'
    $Script:connectIPPMaxRetries = 2
    $Script:connectIPPRetries = 0
    function connectIPPSSession ([switch]$EnableSearchOnlySession) {
        try {
            $connectIPPSSessionParams = @{
                UserPrincipalName = $AdminUPN
                DisableWAM        = $true
                ShowBanner        = $false
                CommandName       = @('Get-ComplianceSearch', 'New-ComplianceSearch', 'Remove-ComplianceSearch', 'Start-ComplianceSearch', 'Get-ComplianceSearchAction', 'New-ComplianceSearchAction')
                ErrorAction       = 'Stop'
                WarningAction     = 'SilentlyContinue'
            }
            if ($EnableSearchOnlySession) { $connectIPPSSessionParams['EnableSearchOnlySession'] = $true }
            Disconnect-ExchangeOnline -Confirm:$false
            Connect-IPPSSession @connectIPPSSessionParams
        }
        catch {
            # Connect-IPPSSession fails often, and often it is for no good reason...
            if ($Script:connectIPPRetries -lt $Script:connectIPPMaxRetries) {
                "Failed on Connect-IPPSSession.  Will retry maximum $($Script:connectIPPMaxRetries) times, pausing 10 seconds between attempts." |
                Write-Warning
                $Script:connectIPPRetries++
                Start-Sleep -Seconds 10
                connectIPPSSession
            }
            else { throw }
        }
    }
    connectIPPSSession

    $Script:SearchName = if ($PSBoundParameters.ContainsKey('SearchNameOverride')) { $SearchNameOverride } else {
        "Export-EXOMailbox.ps1_$($MailboxPSmtp)_$($MailboxSelection -replace 'Both','Primary+Archive')$(if ($InactiveMailbox) { '_InactiveMBX' })"
    }

    Write-Progress @progress -Status "New-ComplianceSearch (-Name '$($Script:SearchName)')"
    $ComplianceSearchParams = @{
        Name                                  = $Script:SearchName
        ContentMatchQuery                     = "$($SelectedFolders.FolderQuery -join ' OR ')"
        ExchangeLocation                      = "$(switch ($InactiveMailbox) {$true {'.'}})$($MailboxPSmtp)"
        AllowNotFoundExchangeLocationsEnabled = $InactiveMailbox
    }
    # In case we're re-trying the script (i.e., starting over from scratch), we'll need to ensure a unique search name.
    # (as of November 2025, recreating a search by the same name in PowerShell results in the new search being invisible/unmanagement in the Purview Compliance Center website):
    $Script:ComplianceSearch = Get-ComplianceSearch $Script:SearchName -ErrorAction SilentlyContinue
    if ($Script:ComplianceSearch) {
        if ($PSCmdlet.ShouldProcess("Message", 'Press Y or Enter to provide a new search name or N to cancel and delete the conflicting search manually before re-running this script.', "'$($Script:SearchName)' already exists")) {
            $Script:SearchName = $null; $Script:SearchName = Read-Host -Prompt "Enter new and unique name for the search" -ErrorAction Stop
        }
        else {
            Write-Warning -Message "Ending script prematurely due to existing compliance search found with conflicting name '$($Script:SearchName)'."
            break
        }
    }
    if ($null -ne $Script:SearchName) {
        $Script:ComplianceSearchParams['Name'] = $Script:SearchName
        $Script:ComplianceSearch = New-ComplianceSearch @ComplianceSearchParams -ErrorAction Stop -Confirm:$false
    }
    else { break }

    # Starting with ExchangeOnlineManagement v3.0.9, we need to connect with the -EnableSearchOnlySession before we can start searches or perform search actions:
    Disconnect-ExchangeOnline -Confirm:$false; connectIPPSSession -EnableSearchOnlySession

    Write-Progress @progress -Status "Start-ComplianceSearch (-Name '$($Script:SearchName)')"
    Start-ComplianceSearch $Script:SearchName -ErrorAction Stop
    do {
        Write-Progress @progress -Status "Waiting for compliance search to complete (search name: '$($Script:SearchName)')"
        Start-Sleep -Seconds 5
        $Script:ComplianceSearch = Get-ComplianceSearch $Script:SearchName -ErrorAction Stop
    }
    while ($Script:ComplianceSearch.Status -ne 'Completed')

    if ($Script:ComplianceSearch.Items -gt 0) {

        Write-Progress @progress -Status 'New-ComplianceSearchAction (-Preview)'
        $ComplianceSearchPreview = New-ComplianceSearchAction -SearchName $Script:SearchName -Preview -ErrorAction Stop -Confirm:$false
        do {
            Write-Progress @progress -Status "Waiting for preview of compliance search results (search name: '$($Script:SearchName)')"
            Start-Sleep -Seconds 5
            $ComplianceSearchPreview = Get-ComplianceSearchAction "$($Script:SearchName)_Preview" -ErrorAction Stop
        }
        while ($ComplianceSearchPreview.Status -ne 'Completed')

        Write-Progress @progress -Status 'Get-ComplianceSearch, parsing/processing search results'
        $Script:ComplianceSearch = Get-ComplianceSearch $Script:SearchName -ErrorAction Stop
        [PSCustomObject]@{
            SearchName     = $Script:ComplianceSearch.Name
            Status         = $Script:ComplianceSearch.Status
            SuccessResults = $Script:ComplianceSearch.SuccessResults
            Items          = $Script:ComplianceSearch.Items
            SizeMB         = [math]::Round($Script:ComplianceSearch.Size / 1MB, 2)
            SizeGB         = [math]::Round($Script:ComplianceSearch.Size / 1GB, 2)
            ExportPSTUrl   = 'https://compliance.microsoft.com/contentsearchv2?viewid=search'
        }
    }
    else {
        Write-Warning -Message "The compliance search ('$($Script:SearchName)') didn't return any results."
    }
    ###########----------------------------------#
    #endregion# Create and run compliance search #
    ###########----------------------------------#
}
catch { throw }
finally { Disconnect-ExchangeOnline -Confirm:$false }

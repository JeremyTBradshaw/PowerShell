<#
    .Synopsis
    Export EXO mailboxes to PST.  Enables folder selection and preserves folder structure.

    .Description
    Intended for interactive use by EXO/SCC administrators with enough access to use the following Cmdlets from the
    ExchangeOnlineManagement PS module (a.k.a., EXO V2; https://www.powershellgallery.com/packages/exchangeonlinemanagement):

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

    .Example
    .\Export-EXOMailbox.ps1 -AdminUPN admin1@contoso.onmicrosoft.com -MailboxPSmtp user1@contoso.com

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
#Requires -PSEdition Desktop
#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'ExchangeOnlineManagement'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'; ModuleVersion = '2.0.5' }
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$AdminUPN,

    [Parameter(Mandatory)]
    [string]$MailboxPSmtp,

    [ValidateSet('Primary', 'Archive', 'Both')]
    [string]$MailboxSelection = 'Both',

    [switch]$InactiveMailbox
)

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
    Connect-ExchangeOnline -UserPrincipalName $AdminUPN -CommandName Get-Mailbox -ShowBanner:$false -ErrorAction Stop #<--: CommandName reduces time to connect, auto-includes all *-EXO*** modern Cmdlets.

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
            Get-EXOMailboxFolderStatistics $SelectedMailbox.Guid.ToString() -IncludeSoftDeletedRecipients:$InactiveMailbox -ErrorAction Stop
        }
        if (@('Archive', 'Both') -contains $MailboxSelection) {

            if ($SelectedMailbox.ArchiveGuid.ToString() -like '000*') {

                Write-Warning -Message "Mailbox $($MailboxPSmtp) is not enabled with an Archive mailbox."
            }
            else {
                Write-Progress @progress -Status 'Get-EXOMailboxFolderStatistics (archive mailbox)'
                Get-EXOMailboxFolderStatistics $SelectedMailbox.Guid.ToString() -IncludeSoftDeletedRecipients:$InactiveMailbox -Archive -ErrorAction Stop
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
    function connectIPPSSession {
        try {
            Disconnect-ExchangeOnline -Confirm:$false
            Connect-IPPSSession -UserPrincipalName $AdminUPN -CommandName *-ComplianceSearch, *-ComplianceSearchAction -ErrorAction Stop -WarningAction SilentlyContinue
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

    $SearchName = "Mailbox-Search_$($MailboxPSmtp)"

    Write-Progress @progress -Status "New-ComplianceSearch (-Name '$($SearchName))"
    $ComplianceSearchParams = @{

        Name                                  = $SearchName
        ContentMatchQuery                     = "$($SelectedFolders.FolderQuery -join ' OR ')"
        ExchangeLocation                      = "$(switch ($InactiveMailbox) {$true {'.'}})$($MailboxPSmtp)"
        AllowNotFoundExchangeLocationsEnabled = $InactiveMailbox
    }
    # In case we're re-trying the script (i.e., starting over from scratch), we'll first try to delete the Compliance Search (if it already exists):
    $ComplianceSearch = Get-ComplianceSearch $SearchName -ErrorAction SilentlyContinue
    if ($ComplianceSearch) {

        Write-Progress @progress -Status "Removing pre-existing compliance search '$($SearchName)' before creating a new one by the same name.  Then sleeping 30 seconds before proceeding."
        Remove-ComplianceSearch $SearchName -Confirm:$false -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 30
    }
    $ComplianceSearch = New-ComplianceSearch @ComplianceSearchParams -ErrorAction Stop

    Write-Progress @progress -Status "Start-ComplianceSearch (-Name '$($SearchName))"
    Start-ComplianceSearch $SearchName -ErrorAction Stop
    do {
        Write-Progress @progress -Status "Waiting for compliance search to complete (search name: '$($SearchName)')"
        Start-Sleep -Seconds 5
        $ComplianceSearch = Get-ComplianceSearch $SearchName -ErrorAction Stop
    }
    while ($ComplianceSearch.Status -ne 'Completed')

    if ($ComplianceSearch.Items -gt 0) {

        Write-Progress @progress -Status 'New-ComplianceSearchAction (-Preview)'
        $ComplianceSearchPreview = New-ComplianceSearchAction -SearchName $SearchName -Preview -ErrorAction Stop
        do {
            Write-Progress @progress -Status "Waiting for preview of compliance search results (search name: '$($SearchName)')"
            Start-Sleep -Seconds 5
            $ComplianceSearchPreview = Get-ComplianceSearchAction "$($SearchName)_Preview" -ErrorAction Stop
        }
        while ($ComplianceSearchPreview.Status -ne 'Completed')

        Write-Progress @progress -Status 'Get-ComplianceSearch, parsing/processing search results'
        $ComplianceSearch = Get-ComplianceSearch $SearchName -ErrorAction Stop
        [PSCustomObject]@{
            SearchName     = $ComplianceSearch.Name
            Status         = $ComplianceSearch.Status
            SuccessResults = $ComplianceSearch.SuccessResults
            Items          = $ComplianceSearch.Items
            SizeMB         = [math]::Round($SearchResultBytes / 1MB, 2)
            SizeGB         = [math]::Round($SearchResultBytes / 1GB, 2)
            ExportPSTUrl   = 'https://compliance.microsoft.com/contentsearchv2?viewid=search'
        }
    }
    else {
        Write-Warning -Message "The compliance search ('$($SearchName)') didn't return any results."
    }
    ###########----------------------------------#
    #endregion# Create and run compliance search #
    ###########----------------------------------#
}
catch { throw }
finally { Disconnect-ExchangeOnline -Confirm:$false }

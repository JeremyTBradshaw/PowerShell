<#
    .Synopsis
    Generate KeyQL query for Purview Content Search to search/export entire folders from a mailbox.

    .Description
    Via MC1238428, in March 2026, Microsoft stopped synchronizing Compliance Searches in PowerShell with the Purview UI.  As a
    result, this script has been reduced down to a mere KeyQL generator for searches that you now do in the UI instead.

    .Parameter MailboxPSmtp
    Specifies the PrimarySmtpAddress of the mailbox to export.

    .Parameter MailboxSelection
    Provides the choice of 'Primary', 'Archive', or 'Both' (default).  The Out-Gridview folder picker UI will show
    which mailbox the folders reside in, for easy selection of folders from either mailbox location.

    .Parameter InactiveMailbox
    Indicates the mailbox is an Inactive Mailbox.

    .Example
    .\Export-EXOMailbox.ps1 -MailboxPSmtp user1@contoso.com

    .Example
    .\Export-EXOMailbox.ps1 -MailboxPSmtp user1@contoso.com -SearchNameOverride User1Export_2022-01-26

    .Example
    .\Export-EXOMailbox.ps1 -MailboxPSmtp user1@contoso.com -MailboxSelection Primary

    .Example
    .\Export-EXOMailbox.ps1 -MailboxPSmtp user1@contoso.com -MailboxSelection Archive -InactiveMailbox

    .OUTPUTS
    Sets a global variable $Global:KeyQLForContentSearch.  Also outputs to a PSCustomObject.

    .NOTES
    V1.0 (2026-04-21): Initial version (repurposed now-deprecated Export-EXOMailbox.ps1).
#>
#Requires -Modules @{ ModuleName = 'ExchangeOnlineManagement'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'; ModuleVersion = '3.9.0' }
[CmdletBinding()]
param (
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

    if (-not (Get-ConnectionInformation -ErrorAction:SilentlyContinue).ConnectionUri -eq 'https://outlook.office365.com') {
        Write-Warning -Message "Run Connect-ExchangeOnline before calling this script."; break
    }

    ########---------------------------------#
    #region# Find mailbox and select folders #
    ########---------------------------------#

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

    $keyQL = [PSCustomObject]@{
        KeyQL = "$($SelectedFolders.FolderQuery -join ' OR ')"
    }
    $Global:KeyQLForContentSearch = $keyQL
    Write-Host -ForegroundColor Green -Object 'The query has been saved to $KeyQLForContentSearch.  Try $KeyQLForContentSearch | fl or $KeyQLForContentSearch.KeyQL.'
    $keyQL
}
catch { throw }
finally {}

<#
    .SYNOPSIS
    Get folderid in 'Guid format' for Exchange mailbox folders for use with Microsoft Purview Compliance Searches.

    .PARAMETER Identity
    Specifies the mailbox to list folders for.  Pass-through parameter for Get-MailboxFolderStatistics.

    .PARAMETER Archive
    Indicates to target the Archive mailbox instead of the Primary.  Pass-through parameter for Get-MailboxFolderStatistics.

    .PARAMETER FolderScope
    Narrows the scope of which folders to list.  Pass-through parameter for Get-MailboxFolderStatistics.
Get
    .PARAMETER IncludeSoftDeletedRecipients
    Pass-through parameter for Get-MailboxFolderStatistics.

    .PARAMETER UseFolderPicker
    Invoke's Out-Gridview to enable specifc folder selections.

    .PARAMETER ConvertId
    Use this option to supply one or more FolderId values, in the encoded format as returned in results from
    Get-MailboxFolderStatistics.  This option is handy if encoded FolderId's have already been obtained and just need
    to be converted.

    .EXAMPLE
    .\Get-MailboxFolderId -Identity 'Jeremy Bradshaw'

    .EXAMPLE
    .\Get-MailboxFolderId -Identity 'Jeremy Bradshaw' -Archive -FolderScope Calendar -UseFolderPicker

    .EXAMPLE
    .\Get-MailboxFolderId -ConvertId 'LgAAAACLKa67viTPS4YIWrWNWo1kAQDLk1Jju/W9TYZ8G+Dk/ONhAAAAAICvAAAB', 'LgAAAACLKa67viTPS4YIWrWNWo1kAQDLk1Jju/W9TYZ8G+Dk/ONhAAAAAICaAAAB'

    .EXAMPLE
    $Mailboxes = Get-Mailbox -ResultSize Unlimited
    $MailboxFolders = @()
    $MailboxFolders = foreach ($mailbox in $Mailboxes) {Get-MailboxFolderStatistics $mailbox.Guid.ToString() -FolderScope RecoverableItems | Where-Object {$_.FolderPath -notlike '*Deletions*'}}
    $FolderIds = .\Get-MailboxFolderId -ConvertId $MailboxFolders.FolderId
    $ContentMatchQuery = "(c:c)(kind=meetings)(subjecttitle=""Dancy Party for Lunch"" NOT ((folderid:" + $($FolderIds.FolderId -join ') OR (folderid:') + "))"
    # ^^ Produces a query to exclude all RecoverableItems folders (minus Deletions folder), similar to the following:
    # (c:c)(kind=meetings)(subjecttitle="Dancy Party for Lunch" NOT ((folderid:cb935263bbf5bd4d867c1be0e4fce36100000000010d0000) OR (folderid:cb935263bbf5bd4d867c1be0e4fce36100001323a6b00000) OR (folderid:cb935263bbf5bd4d867c1be0e4fce36100001323a6b10000))
#>
#Requires -Version 4.0
[CmdletBinding(DefaultParameterSetName = 'FindFolders')]
param (
    [Parameter(Mandatory, ParameterSetName = 'FindFolders', HelpMessage = 'ExchangeGuid,DistinguishedName are good choices here, especially with SoftDeleted/Inactive mailboxes.')]
    [string]$Identity,

    [Parameter(ParameterSetName = 'FindFolders')]
    [switch]$Archive,

    [Parameter(ParameterSetName = 'FindFolders')]
    [ValidateSet('All', 'Archive', 'Calendar', 'Contacts', 'ConversationHistory', 'DeletedItems', 'Drafts', 'Inbox', 'JunkEmail', 'Journal', 'LegacyArchiveJournals',
        'ManagedCustomFolder', 'NonIpmRoot', 'Notes', 'Outbox', 'Personal', 'RecoverableItems', 'RssSubscriptions', 'SentItems', 'SyncIssues', 'Tasks')]
    [string]$FolderScope = 'All',

    [Parameter(ParameterSetName = 'FindFolders')]
    [switch]$IncludeSoftDeletedRecipients,

    [Parameter(ParameterSetName = 'FindFolders')]
    [switch]$UseFolderPicker,

    [Parameter(ParameterSetName = 'ConvertId')]
    [ValidateScript(
        { if ($_.Length -ne 64) { throw "FolderId '$($_)' is invalid.  Supply one or more FolderId values as returned from Get-MailboxFolderStatistics." } else { $true } }
    )]
    [string[]]$ConvertId
)

#======#----------------------------#
#region# Initialization & Variables #
#======#----------------------------#

# Verify required commands are available:
$_requiredCmdlets = @('Get-MailboxFolderStatistics')
$_missingCmdlets = @()
foreach ($_cmdlet in $_requiredCmdlets) {

    if (-not (Get-Command $_cmdlet -ErrorAction SilentlyContinue)) { $_missingCmdlets += $_cmdlet }
}
if ($_missingCmdlets.Count -ge 1) {

    throw "Missing cmdlets: $($_missingCmdlets -join ', ').  Required cmdlets: $($_requiredCmdlets -join ', ')."
}

#=========#----------------------------#
#endregion# Initialization & Variables #
#=========#----------------------------#



#======#-----------#
#region# Functions #
#======#-----------#

function getFolderId ([string]$EncodedFolderId) {

    # Borrowed code (start)
    $encoding = [System.Text.Encoding]::GetEncoding('us-ascii')
    $nibbler = $encoding.GetBytes('0123456789abcdef')
    $folderIdBytes = [Convert]::FromBase64String($EncodedFolderId)
    $indexIdBytes = New-Object byte[] 48
    $indexIdIdx = 0
    $folderIdBytes | Select-Object -skip 23 -First 24 | ForEach-Object {

        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF]
    }
    # Borrowed code (end)

    [PSCustomObject]@{ FolderId = $encoding.GetString($indexIdBytes) }
}

#=========#-----------#
#endregion# Functions #
#=========#-----------#



#======#------------------------#
#region# Scenario: Find Folders #
#======#------------------------#

if ($PSCmdlet.ParameterSetName -eq 'FindFolders') {
    try {
        $FolderStatistics = Get-MailboxFolderStatistics $Identity -FolderScope $FolderScope -Archive:$Archive -IncludeSoftDeletedRecipients:$IncludeSoftDeletedRecipients -ErrorAction Stop
        $SelectedFolders = if ($UseFolderPicker) {

            $FolderStatistics | Select-Object FolderPath, FolderAndSubFolderSize, ItemsInFolderAndSubFolders, FolderId |
            Out-GridView -OutputMode Multiple -Title 'Select folders to obtain FolderId values for:' -ErrorAction Stop

        }
        else { $FolderStatistics }

        foreach ($folder in $SelectedFolders) {

            getFolderId -EncodedFolderId $folder.FolderId | Select-Object @{Name = 'FolderPath'; Expression = { $folder.FolderPath } }, FolderId
        }
    }
    catch { throw }
}

#=========#------------------------#
#endregion# Scenario: Find Folders #
#=========#------------------------#



#======#---------------------------------------#
#region# Scenario: Convert encoded FolderId(s) #
#======#---------------------------------------#

if ($PSCmdlet.ParameterSetName -eq 'ConvertId') {

    foreach ($id in $ConvertId) {
        try {
            getFolderId -EncodedFolderId $id | Select-Object @{Name = 'EncodedFolderId'; Expression = { $id } }, FolderId
        }
        catch {
            Write-Warning "Failed to convert ID '$($id)'.`nError: $($_.Exception)"
        }
    }
}

#=========#---------------------------------------#
#endregion# Scenario: Convert encoded FolderId(s) #
#=========#---------------------------------------#

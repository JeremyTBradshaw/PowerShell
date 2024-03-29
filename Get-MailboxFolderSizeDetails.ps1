<#
    .SYNOPSIS
    A wrapper for Get-MailboxFolderStatistics with output mainly focused on numeric values for size properties.

    .PARAMETER Identity
    Passthrough for Get-MailboxFolderStatistics' -Identity parameter.

    .PARAMETER Archive
    Passthrough for Get-MailboxFolderStatistics' -Archive parameter.

    .PARAMETER FolderScope
    Passthrough for Get-MailboxFolderStatistics' -FolderScope parameter.

    .PARAMETER IncludeOldestAndNewestItems
    Passthrough for Get-MailboxFolderStatistics' -IncludeOldestAndNewestItems parameter.

    .PARAMETER IncludeSoftDeletedRecipients
    Passthrough for Get-MailboxFolderStatistics' -IncludeSoftDeletedRecipients parameter.
#>
#Requires -Version 5.1
[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [string[]]$Identity,
    [switch]$Archive,
    [ValidateSet(
        'All', 'Archive', 'Calendar', 'Contacts', 'ConversationHistory', 'DeletedItems', 'Drafts', 'Inbox', 'JunkEmail',
        'Journal', 'LegacyArchiveJournals', 'ManagedCustomFolder', 'NonIpmRoot', 'Notes', 'Outbox', 'Personal',
        'RecoverableItems', 'RssSubscriptions', 'SentItems', 'SyncIssues', 'Tasks'
    )]
    [string]$FolderScope = 'All',
    [switch]$IncludeOldestAndNewestItems
)
begin {
    if ((Get-Command Get-MailboxFolderStatistics, Get-Recipient -ErrorAction SilentlyContinue).Count -ne 2) {

        throw 'An active Exchange PowerShell session is required, along with access to the Get-CalendarProcessing and Get-Recipient cmdlets.'
    }
    $Script:startTime = [datetime]::Now
    $Script:stopwatchMain = [System.Diagnostics.Stopwatch]::StartNew()
    $Script:stopwatchPipeline = [System.Diagnostics.Stopwatch]::new()
    $Script:progress = @{
        Id              = 0
        Activity        = "$($PSCmdlet.MyInvocation.MyCommand.Name)"
        Status          = "Start time: $($startTime.ToString('yyyy-MM-ddTHH:mm:ss')) | Elapsed: $($stopWatchMain.Elapsed.ToString('hh\:mm\:ss'))"
        PercentComplete = -1
    }
    Write-Progress @progress
    $stopWatchPipeline.Start()
}
process {
    try {
        $Script:progress.Status = "Start time: $($startTime.ToString('yyyy-MM-ddTHH:mm:ss')) | Elapsed: $($stopWatchMain.Elapsed.ToString('hh\:mm\:ss'))"
        $Script:progress.CurrentOperation = "Resource: $($Identity[0]) - Get-MailboxStatistics..."
        Write-Progress @progress

        $getMailboxFolderStatsParams = @{
            Identity                    = $Identity[0]
            FolderScope                 = $FolderScope
            Archive                     = $Archive
            IncludeAnalysis             = $true
            IncludeOldestAndNewestItems = $IncludeOldestAndNewestItems
            ErrorAction                 = 'Stop'
        }
        $mailboxFolderStats = Get-MailboxFolderStatistics @getMailboxFolderStatsParams

        $folderCounter = 0
        foreach ($folder in $mailboxFolderStats) {

            $folderCounter++
            if ($stopWatchPipeline.ElapsedMilliseconds -ge 200) {
                $Script:progress.Status = "Start time: $($startTime.ToString('yyyy-MM-ddTHH:mm:ss')) | Elapsed: $($stopWatchMain.Elapsed.ToString('hh\:mm\:ss'))"
                $Script:progress.CurrentOperation = "Resource: $($Identity[0]) - Processing..."
                Write-Progress @progress
                Write-Progress -Activity Processing -Id 1 -ParentId 0 -PercentComplete (($folderCounter / $mailboxFolderStats.Count) * 100)
                $stopWatchPipeline.Restart()
            }

            [int64]$FolderSize = $folder.FolderSize -replace '(.*\()|,|(\s.*)'
            [int64]$FolderAndSubfolderSize = $folder.FolderAndSubfolderSize -replace '(.*\()|,|(\s.*)'
            [int64]$TopSubjectSize = $folder.TopSubjectSize -replace '(.*\()|,|(\s.*)'

            $folderOutput = [PSCustomObject]@{
                Name                       = $folder.Name
                FolderPath                 = $folder.FolderPath
                FolderSizeGB               = [math]::Round(($FolderSize / 1GB), 2)
                ItemsInFolder              = $folder.ItemsInFolder
                AvgItemSizeMB              = [math]::Round((($FolderSize / [math]::Max(1, $folder.ItemsInFolder)) / 1MB), 1)
                ItemsInFolderAndSubfolders = $folder.ItemsInFolderAndSubfolders
                FolderAndSubfolderSizeGB   = [math]::Round(($FolderAndSubfolderSize / 1GB), 2)
                TopSubjectSizeMB           = [math]::Round((($TopSubjectSize / [math]::Max(1, $folder.TopSubjectCount)) / 1MB), 1)
                TopSubjectTotalSizeGB      = [math]::Round(($TopSubjectSize / 1GB), 2)
                TopSubjectClass            = $folder.TopSubjectClass
                TopSubjectCount            = $folder.TopSubjectCount
                TopSubject                 = $folder.TopSubject
                RecoverableItemsFolder     = $folder.RecoverableItemsFolder
                FolderId                   = $folder.FolderId
            }
            if ($IncludeOldestAndNewestItems) {
                $folderOutput |
                Add-Member -NotePropertyName NewestItemLastModifiedDate -NotePropertyValue $folder.NewestItemLastModifiedDate -PassThru |
                Add-Member -NotePropertyName NewestItemReceivedDate -NotePropertyValue $folder.NewestItemReceivedDate -PassThru |
                Add-Member -NotePropertyName OldestItemLastModifiedDate -NotePropertyValue $folder.OldestItemLastModifiedDate -PassThru |
                Add-Member -NotePropertyName OldestItemReceivedDate -NotePropertyValue $folder.OldestItemReceivedDate -PassThru |
                Add-Member -NotePropertyName NewestDeletedItemLastModifiedDate -NotePropertyValue $folder.NewestDeletedItemLastModifiedDate -PassThru |
                Add-Member -NotePropertyName NewestDeletedItemReceivedDate -NotePropertyValue $folder.NewestDeletedItemReceivedDate -PassThru |
                Add-Member -NotePropertyName OldestDeletedItemLastModifiedDate -NotePropertyValue $folder.OldestDeletedItemLastModifiedDate -PassThru |
                Add-Member -NotePropertyName OldestDeletedItemReceivedDate -NotePropertyValue $folder.OldestDeletedItemReceivedDate
            }
            $folderOutput
        }
    }
    catch {
        Write-Warning -Message "Failed on Identity: $($Identity[0])"; throw
    }
}
end {
    Write-Progress @progress -Completed
}

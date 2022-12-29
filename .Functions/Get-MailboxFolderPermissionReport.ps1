function Get-MailboxFolderPermissionReport {
    <#
        .Synopsis
        Get a report of assigned mailbox folder permissions for a given mailbox.

        .Parameter Identity
        Treat this parameter exactly like Get-MailboxFolderStatistics cmdlet's -Identity parameter.

        .Parameter ReportMode
        Summary mode outputs one permission object per folder, with columns for each permission role.  Detailed mode,
        the default, outputs one permission object for every access right entry for every folder.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Identity,

        [ValidateSet('Summary', 'Detailed')]
        [string]$ReportMode = 'Detailed'
    )
    begin {
        if (-not (Get-Command Get-MailboxFolderStatistics, Get-MailboxFolderPermission -ErrorAction SilentlyContinue)) {

            throw 'This script requires an active session to Exchange/EXO and access to the Get-MailboxFolderStatistics and Get-MailboxFolderPermission cmdlets.'
        }
        $dtNow = [datetime]::Now
        $Progress = @{

            Activity        = "$($PSCmdlet.MyInvocation.MyCommand.Name) -Start time: $($dtNow)"
            PercentComplete = -1
        }
    }
    process {
        try {
            # The following line gets around an issue with ExchangeOnlineManagement PS module v3.0.0 where $ProgressPreference is repeatedly set to 'SilentlyContinue'
            $ProgressPreference = 'Continue'
            Write-Progress @Progress -CurrentOperation "Get-MailboxFolderStatistics -Identity '$($Identity)'"

            $MailboxFolderStatistics = Get-MailboxFolderStatistics -Identity $Identity -ErrorAction Stop

            $_folderCounter = 0
            foreach ($folder in $MailboxFolderStatistics) {
                try {
                    $_folderCounter++

                    $Progress['PercentComplete'] = (($_folderCounter / $MailboxFolderStatistics.Count) * 100) 
                    # The following line gets around an issue with ExchangeOnlineManagement PS module v3.0.0 where $ProgressPreference is repeatedly set to 'SilentlyContinue'
                    $ProgressPreference = 'Continue'
                    Write-Progress @Progress -Status 'Getting permissions...' -CurrentOperation "Current folder: $($folder.FolderPath)"
                    
                    $Permissions = Get-MailboxFolderPermission -Identity "$($Identity):$($folder.FolderId)" -ErrorAction Stop
                }
                catch {
                    Write-Warning -Message "Failed to get permissions for FolderPath: $($folder.FolderPath)"
                }

                switch ($ReportMode) {

                    'Summary' {
                        [PSCustomObject] @{
                            'Folder'           = $folder.FolderPath
                            'None'             = ($Permissions | Where-Object { $folder.AccessRights -eq 'None' }).User -join ', '
                            'Custom'           = ($Permissions | Where-Object { ($folder.AccessRights -like '*') -and ($folder.AccessRights -ne 'None') -and ($folder.AccessRights -ne 'Owner') -and ($folder.AccessRights -ne 'PublishingEditor') -and ($folder.AccessRights -ne 'Editor') -and ($folder.AccessRights -ne 'PublishingAuthor') -and ($folder.AccessRights -ne 'Author') -and ($folder.AccessRights -ne 'NonEditingAuthor') -and ($folder.AccessRights -ne 'Reviewer') -and ($folder.AccessRights -ne 'Contributor') }).User -join ', '
                            'Owner'            = ($Permissions | Where-Object { $folder.AccessRights -eq 'Owner' }).User -join ', '
                            'PublishingEditor' = ($Permissions | Where-Object { $folder.AccessRights -eq 'PublishingEditor' }).User -join ', '
                            'Editor'           = ($Permissions | Where-Object { $folder.AccessRights -eq 'Editor' }).User -join ', '
                            'PublishingAuthor' = ($Permissions | Where-Object { $folder.AccessRights -eq 'PublishingAuthor' }).User -join ', '
                            'Author'           = ($Permissions | Where-Object { $folder.AccessRights -eq 'Author' }).User -join ', '
                            'NonEditingAuthor' = ($Permissions | Where-Object { $folder.AccessRights -eq 'NonEditingAuthor' }).User -join ', '
                            'Reviewer'         = ($Permissions | Where-Object { $folder.AccessRights -eq 'Reviewer' }).User -join ', '
                            'Contributor'      = ($Permissions | Where-Object { $folder.AccessRights -eq 'Contributor' }).User -join ', '
                            'MailboxIdentity'  = $Identity
                            'FolderId'         = $folder.FolderId
                        }
                    }
                    'Detailed' {
                        foreach ($perm in $Permissions) {

                            [PSCustomObject]@{
                                Folder          = $folder.FolderPath
                                User            = $perm.User
                                AccessRights    = $perm.AccessRights
                                SharingFlags    = $perm.SharingFlags -join ', '
                                MailboxIdentity = $Identity
                                FolderId        = $folder.FolderId
                            }
                        }
                    }
                }
            }
        }
        catch { throw }
    }
    end { Write-Progress @Progress -Completed }
}

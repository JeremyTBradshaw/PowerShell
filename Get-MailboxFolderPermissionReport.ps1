<#
    .SYNOPSIS
    Get a report of assigned mailbox folder permissions for a given mailbox.

    .PARAMETER Identity
    Treat this parameter exactly like Get-MailboxFolderStatistics cmdlet's -Identity parameter.
d
    .PARAMETER ReportMode
    Summary mode outputs one permission object per folder, with columns for each permission role.  Detailed mode,
    the default, outputs one permission object for every access right entry for every folder.

    .PARAMETER ResultSize
    This allows for overriding the default behavior (-ResultSize Unlimited), if/when need (e.g., very large mailboxes).

    .PARAMETER ExportCSVs
    Exports 2 CSV files (1 for each report type (Summary/Detailed).

    .PARAMETER OutputFolderPath
    Specify a folder to override the default location of the CSV files, which is $PWD.

    .PARAMETER ShowIEQDuplicateFolderIds
    FolderIds are encoded in a case-sensitive format.  If we disregard case, there are often many duplicate FolderIds.
    This switch will simply trigger a warning in PowerShell to point these folders out.  They should be targeted by
    Folder name/path rather than by FolderId when using Get-/Set-/Add-/Remove-MailboxFolderPermission, or unexpected
    information/changes may be produced.
#>
[CmdletBinding(DefaultParameterSetName = 'NoExport')]
Param (
    [Parameter(ParameterSetName = 'NoExport', Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(ParameterSetName = 'ExportCSVs', Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [string]$Identity,

    [Parameter(ParameterSetName = 'NoExport')]
    [ValidateSet('Summary', 'Detailed')]
    [string]$ReportMode = 'Detailed',

    [Parameter(ParameterSetName = 'NoExport')]
    [Parameter(ParameterSetName = 'ExportCSVs')]
    [Object]$ResultSize = 'Unlimited',

    [Parameter(ParameterSetName = 'ExportCSVs')]
    [switch]$ExportCSVs,

    [Parameter(ParameterSetName = 'ExportCSVs')]
    [ValidateScript(
        {
            if (Test-Path -Path $_ -PathType Container) { $true }
            else { throw "'$($_)' does not appear to be valid folder path (absolute or relative)." }
        }
    )]
    [System.IO.FileInfo]$OutputFolderPath,

    [bool]$ShowIEQDuplicateFolderIds = $true
)
begin {
    if (-not (Get-Command Get-MailboxFolderStatistics, Get-MailboxFolderPermission -ea SilentlyContinue)) {

        throw 'This script requires an active session to Exchange/EXO and access to the Get-MailboxFolderStatistics and Get-MailboxFolderPermission cmdlets.'
    }
    $dtNow = [datetime]::Now
    $Progress = @{

        Activity        = "$($PSCmdlet.MyInvocation.MyCommand.Name) -Start time: $($dtNow)"
        PercentComplete = -1
    }

    $_outputFolder = if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('OutputFolderPath')) { $OutputFolderPath }
    else { $PWD }
}
process {
    try {
        # The following line gets around an issue with ExchangeOnlineManagement PS module v3.0.0 where $ProgressPreference is repeatedly set to 'SilentlyContinue'
        $ProgressPreference = 'Continue'
        Write-Progress @Progress -CurrentOperation "Get-MailboxFolderStatistics -Identity '$($Identity)'"


        $MailboxFolderStatistics = Get-MailboxFolderStatistics -Identity $Identity -ResultSize:$ResultSize -ea Stop

        if ($ShowIEQDuplicateFolderIds) {
            # Warn on case-insensitive duplidate FolderId's which can be problematic if setting permissions using FolderId for the -Identity parameter:
            $_dupeFolderIds = @($MailboxFolderStatistics | Group-Object -Property FolderId | Where-Object { $_.Count -ge 2 })
            if ($_dupeFolderIds.Count -ge 1) {

                "The following folders have duplicate (matching) FolderId's, which means they should only be referenced by folder path, not FolderId, " +
                'for the -Identity parameter of Get-/Set-/Add-/Remove-MailboxFolderPermission cmdlets: ' | Write-Warning

                foreach ($_dupe in ($_dupeFolderIds | Sort-Object -Property Name)) {

                    Write-Warning "$($_dupe.Name) ($($_dupe.Count) folders sharing same FolderId (case-insensitive))."
                }
            }
        }

        if ($PSCmdlet.ParameterSetName -eq 'ExportCSVs') {
            $SummaryCSVFile = New-Item -Path $_outputFolder -Name "MailboxFolderPermissionReport_Summary_$([datetime]::Now.ToString('yyyy-MM-dd_hhmmttzz')).csv" -ItemType File -ea Stop
            $DetailedCSVFile = New-Item -Path $_outputFolder -Name "MailboxFolderPermissionReport_Detailed_$([datetime]::Now.ToString('yyyy-MM-dd_hhmmttzz')).csv" -ItemType File -ea Stop
        }

        $_folderCounter = 0
        foreach ($folder in $MailboxFolderStatistics) {
            try {
                $_folderCounter++

                $Progress['PercentComplete'] = (($_folderCounter / $MailboxFolderStatistics.Count) * 100)
                # The following line gets around an issue with ExchangeOnlineManagement PS module v3.0.0 where $ProgressPreference is repeatedly set to 'SilentlyContinue'
                $ProgressPreference = 'Continue'
                Write-Progress @Progress -Status 'Getting permissions...' -CurrentOperation "Current folder: $($folder.FolderPath)"

                $Permissions = Get-MailboxFolderPermission -Identity "$($Identity):$($folder.FolderId)" -ea Stop
            }
            catch {
                Write-Warning -Message "Failed to get permissions for FolderPath: $($folder.FolderPath)"
            }

            if (($ReportMode -eq 'Summary') -or ($PSCmdlet.ParameterSetName -eq 'ExportCSVs')) {
                $Summary = [PSCustomObject] @{
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
                if ($ReportMode -eq 'Summary') { $Summary }
                else { $Summary | Export-Csv -Path $SummaryCSVFile -Append -NTI -Encoding utf8 -ea Stop }
            }

            if (($ReportMode -eq 'Detailed') -or ($PSCmdlet.ParameterSetName -eq 'ExportCSVs')) {
                foreach ($perm in $Permissions) {
                    [PSCustomObject]@{
                        Folder          = $folder.FolderPath
                        User            = $perm.User
                        AccessRights    = $perm.AccessRights
                        SharingFlags    = $perm.SharingFlags -join ', '
                        MailboxIdentity = $Identity
                        FolderId        = $folder.FolderId
                    } |
                    Tee-Object -Variable _detailedOutput
                    if ($PSCmdlet.ParameterSetName -eq 'ExportCSVs') { $_detailedOutput | Export-Csv -Path $DetailedCSVFile -Append -NTI -Encoding utf8 -ea Stop }
                }
            }
        }
    }
    catch { throw }
}
end { Write-Progress @Progress -Completed }

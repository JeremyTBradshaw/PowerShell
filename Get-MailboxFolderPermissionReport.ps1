<#
    .SYNOPSIS
    Get a report of assigned mailbox folder permissions for a given mailbox.

    .PARAMETER Identity
    Treat this parameter exactly like Get-MailboxFolderStatistics cmdlet's -Identity parameter.
d
    .PARAMETER ResultSize
    This allows for overriding the default behavior (-ResultSize Unlimited), if/when needed (e.g., very large mailboxes).

    .PARAMETER OutputFolderPath
    Specify a folder to override the default location of the CSV files, which is $PWD.

    .PARAMETER ShowIEQDuplicateFolderIds
    FolderIds are encoded in a case-sensitive format.  If we disregard case, there are often many duplicate FolderIds.
    This switch will simply trigger a warning in PowerShell to point these folders out.  They should be targeted by
    Folder name/path rather than by FolderId when using Get-/Set-/Add-/Remove-MailboxFolderPermission, or unexpected
    information/changes may be produced.

    .PARAMETER Force
    Sometimes Microsoft breaks Get-MailboxFolderPermission and a folder's permissions CANNOT be retrieved until Microsoft later
    update the ExchangeOnlineManagement module.  -Force will make script keep going down through the list of folders even if any
    fail on the command.
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [string]$Identity,
    [Object]$ResultSize = 'Unlimited',
    [ValidateScript(
        {
            if (Test-Path -Path $_ -PathType Container) { $true }
            else { throw "'$($_)' does not appear to be valid folder path (absolute or relative)." }
        }
    )]
    [System.IO.FileInfo]$OutputFolderPath,
    [bool]$ShowIEQDuplicateFolderIds = $true,
    [switch]$Force
)
begin {
    function Get-RequiredCommands ([string[]]$Commands) {
        $missingCommands = @(foreach ($c in $Commands) { if (-not (Get-Command $c -ErrorAction:SilentlyContinue)) { $c } })
        if ($missingCommands.Count -ge 1) { $false }
        else { $true }
    }
    if (-not (Get-RequiredCommands Get-MailboxFolderStatistics, Get-MailboxFolderPermission)) {
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

        $mbxFoldersListCSV = New-Item -Path $_outputFolder -Name "MailboxFolders_$([datetime]::Now.ToString('yyyy-MM-ddTHH-mm-sszz')).csv" -ItemType File -ea Stop
        $MailboxFolderStatistics |
        Select-Object @{Name = 'MailboxIdentity'; Expression = { $_.ContentMailboxGuid } }, FolderPath, ContainerClass, Foldertype, CreationTime, FolderId |
        Export-Csv -Path $mbxFoldersListCSV -NoTypeInformation -ErrorAction Stop

        $mbxFolderPermissionsCSV = New-Item -Path $_outputFolder -Name "MailboxFolderPermissions_$([datetime]::Now.ToString('yyyy-MM-ddTHH-mm-sszz')).csv" -ItemType File -ea Stop

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
                if (-not $Force) { throw }
            }

            foreach ($perm in $Permissions) {
                $Details = [PSCustomObject]@{
                    FolderPath      = $folder.FolderPath
                    User            = $perm.User
                    AccessRights    = $perm.AccessRights
                    SharingFlags    = $perm.SharingFlags -join ', '
                    MailboxIdentity = $MailboxFolderStatistics[0].ContentMailboxGuid
                    FolderId        = $folder.FolderId
                }
                $Details | Export-Csv -Path $mbxFolderPermissionsCSV -Append -NTI -Encoding utf8 -ea Stop
            }
        }
    }
    catch { throw }
}
end { Write-Progress @Progress -Completed }

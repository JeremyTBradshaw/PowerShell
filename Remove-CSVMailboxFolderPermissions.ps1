<#
    # Remove-CSVMailboxFolderPermissions
    .SYNOPSIS
    Remove mailbox folder permissions in bulk from a CSV file.

    .PARAMETER CsvInputFilePath
    Specifies the full or relative file path to the CSV containing the FolderPath, FolderId, MailboxIdentity, and User columns.

    .NOTES
    In the input CSV file:
        - FolderPath should be in same format as ouputted by Get-MailboxFolderPermission (i.e., forward slashes).
        - FolderId should be exactly as what is outputted by Get-MailboxFolderStatistics
        - MailboxIdentity should be whatever would work with Add-/Get-/Remove-/Set-MailboxFolderPermission's -Identity parameter.
        - User should be whatever would work with Add-/Get-/Remove-/Set-MailboxFolderPermission's -User parameter.

    Version History:
        - v0.0.1 (2025-06-11): Created initial script.
        - v1.0.0 (2025-TBD..): ...
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) { throw "Couldn't find the file '$($_)'." } else { $true }
        })]
    $CsvInputFilePath
)

function Get-RequiredCommands ([string[]]$Commands) {
    $missingCommands = @(foreach ($c in $Commands) { if (-not (Get-Command $c -ErrorAction:SilentlyContinue)) { $c } })
    if ($missingCommands.Count -ge 1) { $false }
    else { $true }
}
if (-not (Get-RequiredCommands Remove-MailboxFolderPermission)) {
    throw 'This script requires an active session to Exchange/EXO and access to the Remove-MailboxFolderPermission cmdlet.'
}

#region CSV validation
try { $importedCSV = Import-Csv -Path $CsvInputFilePath -ErrorAction Stop }
catch { throw }
$requiredColumns = @('FolderPath', 'FolderId', 'MailboxIdentity', 'User')
$includedColumns = ($importedCSV | Get-Member -MemberType NoteProperty)
$missingColumns = @(); foreach ($c in $requiredColumns) { if ($includedColumns.Name -notcontains $c) { $missingColumns += $c } }
if ($missingColumns.Count -ge 1) {
    throw "CSV file is missing the following required columns: $($missingColumns -join ', ')."
}
#endregion CSV validation

#region Add permissions
if ($PSCmdlet.ShouldProcess("all folders in the CSV", "remove permission(s)")) {
    foreach ($f in $importedCSV) {
        try {
            $commandParams = @{
                User        = $f.User
                ErrorAction = 'Stop'
                Confirm     = $false
            }
            # Attempt by FolderPath 1st (to avoid case-insensitivity matching issues that can be encountered with FolderId).
            [void](Remove-MailboxFolderPermission "$($f.MailboxIdentity):$($f.FolderPath -replace '\/', '\')" @commandParams)
            Write-Host -ForegroundColor Green "Success (Remove by Path): User:$($f.User) | Folder:$($f.FolderPath)"
        }
        catch {
            Write-Host -ForegroundColor Red "Failure (Remove by Path): User:$($f.User) | Folder:$($f.FolderPath)`nError: $($_.Exception.Message)"
            if ($_.Exception.Message -like '*There is no existing permission entry found for user*') { continue }
            if ($_.Exception.Message -like '*matches multiple entries.') { Write-Host -ForegroundColor Red "Ending script prematurely."; break }
            try {
                # If FolderPath fails (likely due to special characters in a folder name), attempt by FolderId.
                [void](Remove-MailboxFolderPermission "$($f.MailboxIdentity):$($f.FolderId)" @commandParams)
                Write-Host -ForegroundColor Green "Success (Remove by FolderId): User:$($f.User) | Folder:$($f.FolderPath)"
            }
            catch {
                Write-Host -ForegroundColor Red "Failure (Remove by FolderId): User:$($f.User) | Folder:$($f.FolderPath)`nError: $($_.Exception.Message)"
                if ($_.Exception.Message -like '*There is no existing permission entry found for user*') { continue }
                if ($_.Exception.Message -like '*matches multiple entries.') { Write-Host -ForegroundColor Red "Ending script prematurely."; break }
                try {
                    # If still uncessful, it might be an ex-user / stale ACE, try with -Force, trying 1st by FolderPath:
                    [void](Remove-MailboxFolderPermission "$($f.MailboxIdentity):$($f.FolderPath -replace '\/', '\')" @commandParams -Force)
                    Write-Host -ForegroundColor Green "Success (Remove by Path, with -Force): User:$($f.User) | Folder:$($f.FolderPath)"
                }
                catch {
                    Write-Host -ForegroundColor Red "Failure (Remove by Path, with -Force): User:$($f.User) | Folder:$($f.FolderPath)`nError: $($_.Exception.Message)"
                    try {
                        # Try -Force by FolderId as a last resort:
                        [void](Remove-MailboxFolderPermission "$($f.MailboxIdentity):$($f.FolderId)" @commandParams -Force)
                        Write-Host -ForegroundColor Green "Success (Remove by FolderId, with -Force): User:$($f.User) | Folder:$($f.FolderPath)"
                    }
                    catch {
                        # Report that it failed and move on:
                        Write-Host -ForegroundColor Red "Failure (Remove by FolderId, with -Force): User:$($f.User) | Folder:$($f.FolderPath)`nError: $($_.Exception.Message)"
                        continue
                    }
                }
            }
        }
    }
}
#endregion Add permissions

<#
    # Add-CSVMailboxFolderPermissions
    .SYNOPSIS
    Add mailbox folder permissions in bulk from a CSV file.

    .DESCRIPTION
    Uses Add-MailboxFolderPermission and if that fails due to existing permission found for the specified user, will
    automatically try with Set-MailboxFolderPermission instead.  Also, will try using the FolderPath first, and if
    unsuccessful, typically due to special characters in the folder path, FolderId will be attempted instead.  If still unable,
    an error message will be written (Write-Host (red), not a thrown error).

    .PARAMETER CsvInputFilePath
    Specifies the full or relative file path to the CSV containing the FolderPath, FolderId, MailboxIdentity, User, and
    AccessRights columns.

    .NOTES
    In the input CSV file:
        - FolderPath should be in same format as ouputted by Get-MailboxFolderPermission (i.e., forward slashes).
        - FolderId should be exactly as what is outputted by Get-MailboxFolderStatistics
        - MailboxIdentity should be whatever would work with Add-/Get-/Remove-/Set-MailboxFolderPermission's -Identity parameter.
        - User should be whatever would work with Add-/Get-/Remove-/Set-MailboxFolderPermission's -User parameter.
        - AccessRights should be again, whatever would work with Add-/Get-/Remove-/Set-MailboxFolderPermission's -AccessRights parameter.

    Version History:
        - 2025-06-10: Created initial script.
        - 2025-TBD..: v1.0.0
#>
[CmdletBinding()]
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
if (-not (Get-RequiredCommands Add-MailboxFolderPermission, Set-MailboxFolderPermission)) {
    throw 'This script requires an active session to Exchange/EXO and access to the Add- and Set-MailboxFolderPermission cmdlets.'
}

#region CSV validation
try { $importedCSV = Import-Csv -Path $CsvInputFilePath -ErrorAction Stop }
catch { throw }
$requiredColumns = @('FolderPath', 'FolderId', 'MailboxIdentity', 'User', 'AccessRights')
$includedColumns = ($importedCSV | Get-Member -MemberType NoteProperty)
$missingColumns = @(); foreach ($c in $requiredColumns) { if ($includedColumns.Name -notcontains $c) { $missingColumns += $c } }
if ($missingColumns.Count -ge 1) {
    throw "CSV file is missing the following required columns: $($missingColumns -join ', ')."
}
#endregion CSV validation

#region Add permissions
foreach ($f in $importedCSV) {
    try {
        $commandParams = @{
            User         = $f.User
            AccessRights = $f.AccessRights
            ErrorAction  = 'Stop'
        }
        if ($f.AccessRights -like '*,*') { $commandParams['AccessRights'] = $f.AccessRights -replace '\s' -split ',' }

        # Attempt by FolderPath 1st (to avoid case-insensitivity matching issues that can be encountered with FolderId).
        [void](Add-MailboxFolderPermission "$($f.MailboxIdentity):$($f.FolderPath -replace '\/', '\')" @commandParams)
        Write-Host -ForegroundColor Green "Success (Add by Path): User:$($f.User) | AccessRights:$($f.AccessRights) | Folder:$($f.FolderPath)"
    }
    catch {
        Write-Host -ForegroundColor Red "Failure (Add by Path): User:$($f.User) | AccessRights:$($f.AccessRights) | Folder:$($f.FolderPath)`nError: $($_.Exception.Message)"
        if ($_.Exception.Message -like "*isn't valid to use for permissions*") { Write-Host -ForegroundColor Red "Ending script prematurely."; break }
        try {
            # If FolderPath fails (likely due to special characters in a folder name), attempt by FolderId.
            [void](Add-MailboxFolderPermission "$($f.MailboxIdentity):$($f.FolderId)" @commandParams)
            Write-Host -ForegroundColor Green "Success (Add by FolderId): User:$($f.User) | AccessRights:$($f.AccessRights) | Folder:$($f.FolderPath)"
        }
        catch {
            Write-Host -ForegroundColor Red "Failure (Add by FolderId): User:$($f.User) | AccessRights:$($f.AccessRights) | Folder:$($f.FolderPath)`nError: $($_.Exception.Message)"
            if ($_.Exception.Message -like "*isn't valid to use for permissions*") { Write-Host -ForegroundColor Red "Ending script prematurely."; break }
            try {
                # If Add- fails, try Set- in case there's an existing permission, trying 1st by FolderPath:
                [void](Set-MailboxFolderPermission "$($f.MailboxIdentity):$($f.FolderPath -replace '\/', '\')" @commandParams)
                Write-Host -ForegroundColor Green "Success (Set by Path): User:$($f.User) | AccessRights:$($f.AccessRights) | Folder:$($f.FolderPath)"
            }
            catch {
                Write-Host -ForegroundColor Red "Failure (Set by Path): User:$($f.User) | AccessRights:$($f.AccessRights) | Folder:$($f.FolderPath)`nError: $($_.Exception.Message)"
                try {
                    # Try Set- by FolderId as a last resort:
                    [void](Set-MailboxFolderPermission "$($f.MailboxIdentity):$($f.FolderId)" @commandParams)
                    Write-Host -ForegroundColor Green "Success (Set by FolderId): User:$($f.User) | AccessRights:$($f.AccessRights) | Folder:$($f.FolderPath)"
                }
                catch {
                    # Report that it failed and move on:
                    Write-Host -ForegroundColor Red "Failure (Set by FolderId): User:$($f.User) | AccessRights:$($f.AccessRights) | Folder:$($f.FolderPath)`nError: $($_.Exception.Message)"; continue
                }
            }
        }
    }
}
#endregion Add permissions

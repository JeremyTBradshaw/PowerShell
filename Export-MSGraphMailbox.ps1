<#
    .SYNOPSIS
    Export an Exchange Online mailbox to OneDrive, SharePoint (Online), M365 Group/Team, or local file system.

    .DESCRIPTION
    The goal of this script is to use the official Microsoft Graph PowerShell SDK modules to export mailbox contents,
    including folder structure, to OneDrive.  The exact details are TBD as I eventually work my way through it.
    I expect many options will be possible to offer over time, such as:

        - Exporting to OneDrive, SharePoint, local file system, etc.
        - Individual emails/items, specific folders, or entire mailbox (primary/archive).
        - Assuming EML files.
        - Maybe headers only option.
        - Open to suggestions (please use GitHub Discussions).

    .NOTES
    Links to the Microsoft Graph PowerShell SDK modules:
        https://www.powershellgallery.com/packages/Microsoft.Graph.Authentication
        https://www.powershellgallery.com/packages/Microsoft.Graph.Files
        https://www.powershellgallery.com/packages/Microsoft.Graph.Mail
        https://www.powershellgallery.com/packages/Microsoft.Graph.Sites

    Links to the Microsoft Graph PowerShell SDK documentation:
        https://docs.microsoft.com/en-us/graph/powershell/installation
        https://docs.microsoft.com/en-us/graph/powershell/get-started
        https://docs.microsoft.com/en-us/graph/powershell/module/overview?view=graph-powershell-latest
        https://docs.microsoft.com/en-us/graph/powershell/v2-authentication-and-authorization

    Links to some related Microsoft Graph API documentation, specific to this script:
        https://docs.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0
        https://docs.microsoft.com/en-us/graph/api/user-get-mailfolders?view=graph-rest-1.0
        https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
        https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.files/new-mgdriveitemuploadsession?view=graph-powershell-1.0
        https://docs.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0
#>
#Requires -Version 7.3.4
#Requires -Module @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.0.0'; Guid = '883916f2-9184-46ee-b1f8-b6a2fb784cee' }
#Requires -Module @{ ModuleName = 'Microsoft.Graph.Mail'; ModuleVersion = '2.0.0'; Guid = '6e4d36b5-7ff2-454b-8572-674b3ab0362b' }
[CmdletBinding(DefaultParameterSetName = 'ExportMailbox')]
param (
    [Parameter(ParameterSetName = 'OneDriveUserId')]
    [Parameter(ParameterSetName = 'DriveId')]
    [Parameter(ParameterSetName = 'SiteId')]
    [Parameter(ParameterSetName = 'FileSystem')]
    [ValidateScript(
        {
            if ($_ -is [string]) {
                if ($_ -match "^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$") { $true }
                elseif ([guid]::TryParse($_, [ref]$null)) { $true }
                else { throw 'Supply a valid Microsoft Account (MSA) email address or Azure AD User UserPrincipalName or GUID (ObjectId).' }
            }
            elseif ($_ -is [guid]) { $true }
            else { throw 'Supply a valid Microsoft Account (MSA) email address or Azure AD User UserPrincipalName or GUID (ObjectId).' }
        }
    )]
    [Object]$MailboxUserId,

    [Parameter(ParameterSetName = 'OneDriveUserId', Mandatory)]
    [ValidateScript(
        {
            if ($_ -is [string]) {
                if ($_ -match "^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$") { $true }
                elseif ([guid]::TryParse($_, [ref]$null)) { $true }
                else { throw 'Supply a valid Microsoft Account (MSA) email address or Azure AD User UserPrincipalName or GUID (ObjectId).' }
            }
            elseif ($_ -is [guid]) { $true }
            else { throw 'Supply a valid Microsoft Account (MSA) email address or Azure AD User UserPrincipalName or GUID (ObjectId).' }
        }
    )]
    [Object]$OneDriveUserId,

    [Parameter(ParameterSetName = 'DriveId', Mandatory)]
    [ValidateScript(
        {
            if (($_ -is [string]) -and ([guid]::TryParse($_, [ref]$null))) { $true }
            elseif ($_ -is [guid]) { $true }
            else { throw 'Supply a valid GUID.' }
        }
    )]
    [Object]$DriveId,

    [Parameter(ParameterSetName = 'SiteId', Mandatory)]
    [ValidateScript(
        {
            if (($_ -is [string]) -and ([guid]::TryParse($_, [ref]$null))) { $true }
            elseif ($_ -is [guid]) { $true }
            else { throw 'Supply a valid GUID.' }
        }
    )]
    [Object]$SiteId,

    [Parameter(ParameterSetName = 'FileSystem')]
    [ValidateScript(
        {
            if (Test-Path -Path $_ -PathType Container) { $true }
            else { throw 'Supply a valid path to a folder.' }
        }
    )]
    [System.IO.FileInfo]$FilePath
)
begin {
    $mgContext = Get-MgContext
    if ($mgContext.Scopes -notlike '*Mail.ReadWrite*') {
        throw 'Before calling this script, you must connect to Microsoft Graph with (at minimum) either of the Mail.ReadWrite or Mail.ReadWrite.All scopes.'
    }
    else {
        switch -Regex ($PSCmdlet.ParameterSetName) {
            '(Drive)|(Group)' {
                $Script:requiredScopes = 'Files.ReadWrite', 'Files.ReadWrite.All'
                $Script:requiredModules = 'Microsoft.Graph.Files'
            }
            '(Site)' {
                $Script:requiredScopes = 'Sites.ReadWrite.All', 'Sites.Selected'
                $Script:requiredModules = 'Microsoft.Graph.Sites'
            }
            default { $Script:requiredScopes = $null; $Script:requiredModules = $null }
        }

        foreach ($requiredScope in $requiredScopes) {
            if ($mgContext.Scopes -notcontains $requiredScope) { $missingScopes += $requiredScope }
        }
        if ($missingScopes) {
            throw "For the chosen export target, your session with Microsoft Graph requires the following scopes " +
            "(in addition to either of the Mail.ReadWrite or Mail.ReadWrite.All scopes): $($missingScopes -join ', ')."
        }
        if ($requiredModules) {
            # Insist on PowerShellGet version 2.0.0 or later, for the -AllowPrerelease parameter:
            if (-not (Get-Module PowerShellGet -ListAvailable | Where-Object { $_.Version.Major -ge 2 })) {
                throw "This script requires PowerShellGet version 2.0.0 or later.  Please update PowerShellGet and try again. " +
                "To update, try: Install-Module PowerShellGet -Scope CurrentUser -Force -AllowClobber.  Will need to restart PowerShell."
            }
            foreach ($requiredModule in $requiredModules) {
                if (-not (Get-Module $requiredModule -ListAvailable | Where-Object { $_.Version.Major -ge 2 })) {
                    Install-Module $requiredModule -Scope CurrentUser -Force -AllowClobber -AllowPrerelease -MinimumVersion 2.0.0-rc3 -ErrorAction Stop
                }
            }
        }
    }
}
process {}
end {}

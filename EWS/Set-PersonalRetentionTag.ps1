<#
    .Synopsis
    Assign a personal retention tag (previously created in EXO) to one or more mailbox folders.

    .Description
    The starting use case for this script was that of excluding Calendar, Tasks, and Notes folders from a default
    (i.e., whole-mailbox) retention tag of 'Move to archive'.  The way to accomplish this is to create a personal tag
    that uses the 'Move to archive' action but has 'Never' (or 0 days) set for the retention period, and then A) add
    that tag to an MRM retention policy, and B) assign that tag to the folders of choice.

    .Parameter EwsManagedApiDllFilePath
    Specifies the path the the Microsoft.Exchange.WebServices.dll file.  Requires product/file version 15.00.0913.015.
    Defaults to 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll', and
    doesn't try to verify otherwise that the installable EWS Managed API has been installed, it just needs access to
    the single DLL file, wherever it may be.

    .Parameter AccessToken
    Supply $Token where the token was obtained using MSGraphPSEssentials (PS module). For example:
    $Token = New-MSGraphAccessToken -ApplicationId <App.Id Guid> -TenantId <Tenant>.onmicrosoft.com -ExoEwsAppOnlyScope

    An access token obtain using a public client flow is also supported.  For example:
    $Token = New-MSGraphAccessToken -ApplicationId <App.Id Guid> -Scopes Ews.AccessAsUser.All

    .Parameter PersonalRetentionTagGuid
    Find the Guid of the intended Personal Tag using EXO PowerShell: Get-RetentionPolicyTag | select Name, Guid

    .Parameter RetentionPeriodInDaysOfTag
    This should match the retention age set in the retention tag specified in the -PersonalRetentionTagGuid parameter.

    .Parameter MailboxPSMTPs
    Specify one or more mailboxes' PrimarySmtpAddress to be processed.

    .Parameter FolderDisplayNames
    Specifies which folders to apply the personal tag to. E.g., -FolderDisplayNames Calendar, Tasks, Notes

    .Notes
    Setting personal retention tags on folders in mailboxes via EWS is not a Microsoft-supported approach.  The
    original articles which I've linked in this help section were intended for use with Exchange 2010 / Outlook 2010. I
    have observed in October 2022 that in Exchange Online and with the current version of Outlook (from M365 Apps for
    Enterprise), the personal tag is set successfully, however Outlook nor OWA recognize it.  Meanwhile the personal
    tag does take effect, and can be confirmed using MFCMAPI.  This is again not Microsoft-supported so it may not be
    worth going through with this operation since it could lead to confusion for users and there is no solution for
    this lack of client visibility of the assigned personal tag.

    Access Token: $Token1 = New-MSGraphAccessToken -ApplicationId 37d171ae-c6bc-4485-be24-74ff28057485 -TenantId <Tenant>.onmicrosoft.com -Certificate (Get-ChildItem cert:\LocalMachine\My\<Thumbprint_of_Cert>)

    .Link
    https://learn.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.wellknownfoldername?view=exchange-ews-api

    .Link
    https://learn.microsoft.com/en-us/archive/blogs/akashb/stamping-retention-policy-tag-using-ews-managed-api-1-1-from-powershellexchange-2010

    .Example
    .\Set-PersonalRetention.ps1 `
        -FolderDisplayNames Calendar, Tasks, Notes `
        -PersonalRetentionTagGuid 41352f92-b179-4e42-bf4b-ea807b495a0b `
        -RetentionPeriodInDaysOfTag 0 `
        -MailboxPSMTPs user1@contoso.com, user2@contoso.com, user3@contoso.com `
        -EwsManagedApiDllFilePath .\Microsoft.Exchange.WebServices.dll `
        -AccessToken $Token
#>
#Requires -Version 5.1 -PSEdition Desktop
#Requires -Modules @{ModuleName = 'MSGraphPSEssentials'; Guid = '7394f3f8-a172-4e18-8e40-e41295131e0b'; RequiredVersion = '0.6.0'}
#Requires -Modules @{ModuleName = 'ExchangeOnlineManagement'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'; RequiredVersion = '3.0.0'}

#using namespace System.Management.Automation
#using namespace Microsoft.Exchange.WebServices.Data

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [System.IO.FileInfo]$EwsManagedApiDllFilePath,

    [Parameter(Mandatory)]
    [Object]$AccessToken,

    [Parameter(Mandatory)]
    [guid]$PersonalRetentionTagGuid,

    [Parameter(Mandatory, HelpMessage = 'This should match the setting in the Personal Tag.')]
    [ValidateRange(0, 24855)]
    [int]$RetentionPeriodInDaysOfTag,

    [Parameter(Mandatory)]
    [string[]]$MailboxPSMTPs,

    [Parameter(Mandatory, HelpMessage = '-FolderDisplayNames Calendar, Tasks, Notes')]
    [string[]]$FolderDisplayNames
)
try {
    Import-Module $EwsManagedApiDllFilePath -ErrorAction Stop

    $ExSvc = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
    $ExSvc.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]::new($AccessToken.access_token)
    $ExSvc.Url = 'https://outlook.office365.com/ews/exchange.asmx'
    $ExSvc.UserAgent = 'MSGraphPSEssentials/0.6.0'
    $ExSvc.Timeout = 150000

    function Update-FolderWithPersonalTag ($ExSvc, $Mailbox, $Folder) {

        $FolderId = if ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName].GetEnumNames() -contains $Folder) { $Folder } else {

            (Get-Folder -ExSvc $ExSvc -Mailbox $Mailbox -FolderDisplayName $Folder).Id
        }

        #PR_POLICY_TAG 0x3019
        $PRPolicyTag = [Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition]::new(0x3019, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)

        #PR_RETENTION_FLAGS 0x301D
        $PRRetentionFlags = [Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition]::new(0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)

        #PR_RETENTION_PERIOD 0x301A
        $PRRetentionPeriod = [Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition]::new(0x301A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)

        #Bind to the folder and update it with retention-related extended properties:
        $BoundFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExSvc, $FolderId)
        $BoundFolder.SetExtendedProperty($PRRetentionFlags, 137)
        $BoundFolder.SetExtendedProperty($PRRetentionPeriod, $RetentionPeriodInDaysOfTag)
        $BoundFolder.SetExtendedProperty($PRPolicyTag, $PersonalRetentionTagGuid.ToByteArray())
        $BoundFolder.Update()
    }

    function Get-Folder ($ExSvc, $Mailbox, $FolderDisplayName, [switch]$Archive) {

        $ParentFolder = if ($Archive) { 'ArchiveRoot' } else { ' Root' }

        $FolderView = [Microsoft.Exchange.WebServices.Data.FolderView]::new(1)
        $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep

        $SearchFilterCollection = [Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection]::new([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        $SearchFilterCollection.Add([Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo]::new([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $FolderDisplayName))

        $Folder = $null
        $Folder = $ExSvc.FindFolders(

            [Microsoft.Exchange.WebServices.Data.FolderId]::new($ParentFolder, $Mailbox),
            $SearchFilterCollection,
            $FolderView
        )

        $Folder
    }

    $Counter = 0
    foreach ($Mailbox in $MailboxPSMTPs) {
        $Counter++
        Write-Progress -Activity "Processing $($Mailbox)" -PercentComplete (($Counter / $MailboxPSMTPs.count) * 100) -Status "Working"

        $ExSvc.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new(

        [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox
    )

    # https://docs.microsoft.com/en-us/archive/blogs/webdav_101/best-practices-ews-authentication-and-access-issues
    $ExSvc.HttpHeaders['X-AnchorMailbox'] = $Mailbox

        foreach ($Folder in $FolderDisplayNames) {

            Update-FolderWithPersonalTag -ExSvc $ExSvc -Mailbox $Mailbox -Folder $Folder
        }
    }
}
catch { throw }

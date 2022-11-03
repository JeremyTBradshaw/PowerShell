<#
    .Synopsis
    Assign a personal retention tag (previously created in EXO) to one or more mailbox folders.

    .Description
    The starting use case for this script was that of excluding Calendar, Tasks, and Notes folders from a default
    (i.e., whole-mailbox) retention tag of 'Move to archive'.  The way to accomplish this is to create a personal tag
    that uses the 'Move to archive' action but has 'Never' (or 0 days) set for the retention period, and then A) add
    that tag to an MRM retention policy, and B) assign that tag to the folders of choice.

    .Notes
    Setting personal retention tags on folders in mailboxes via EWS is not a Microsoft-supported approach.  The
    original articles which I've linked in this help section were intended for use with Exchange 2010 / Outlook 2010. I
    have observed in October 2022 that in Exchange Online and with the current version of Outlook (from M365 Apps for
    Enterprise), the personal tag is set successfully, however Outlook nor OWA recognize it.  Meanwhile the personal
    tag does take effect, and can be confirmed using MFCMAPI.  This is again not Microsoft-supported so it may not be
    worth going through with this operation since it could lead to confusion for users and there is no solution for
    this lack of client visibility of the assigned personal tag.

    .Link
    https://learn.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.wellknownfoldername?view=exchange-ews-api

    .Link
    https://learn.microsoft.com/en-us/archive/blogs/akashb/stamping-retention-policy-tag-using-ews-managed-api-1-1-from-powershellexchange-2010

    .Example
    .\Set-PersonalRetentionTagViaEWS.ps1 `
        -FolderDisplayNames Calendar, Tasks, Notes `
        -PersonalRetentionTagGuid 41352f92-b179-4e42-bf4b-ea807b495a0b `
        -RetentionPeriodInDaysOfTag 0 `
        -MailboxPSMTPs user1@contoso.com, user2@contoso.com, user3@contoso.com `
        -EwsManagedApiDllFilePath .\Microsoft.Exchange.WebServices.dll `
        -AADRegisteredApplicationId 92795d50-1691-46f5-8026-07dc6ef33261
#>
#Requires -Version 5.1 -PSEdition Desktop
#Requires -Modules @{ModuleName = 'MSGraphPSEssentials'; Guid = '7394f3f8-a172-4e18-8e40-e41295131e0b'; RequiredVersion = '0.6.0'}
using namespace System.Management.Automation
using namespace Microsoft.Exchange.WebServices.Data

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    # [ValidateSet('Calendar', 'Tasks', 'Notes')]
    [string[]]$FolderDisplayNames,

    [Parameter(Mandatory)]
    [guid]$PersonalRetentionTagGuid,

    [Parameter(Mandatory)]
    [ValidateRange(0, 24855)]
    [int]$RetentionPeriodInDaysOfTag,

    [Parameter(Mandatory)]
    [string[]]$MailboxPSMTPs,

    [Parameter(Mandatory)]
    [System.IO.FileInfo]$EwsManagedApiDllFilePath,

    [Parameter(Mandatory)]
    [guid]$AADRegisteredApplicationId
)

Import-Module $EwsManagedApiDllFilePath

$NewToken = New-MSGraphAccessToken -ApplicationId $AADRegisteredApplicationId -Scopes Ews.AccessAsUser.All
$ExSvc = [ExchangeService]::new([ExchangeVersion]::Exchange2010_SP1)
$ExSvc.Credentials = [OAuthCredentials]::new($NewToken.access_token)
$ExSvc.Url = 'https://outlook.office365.com/ews/exchange.asmx'
$ExSvc.UserAgent = 'MSGraphPSEssentials/0.6.0'
$ExSvc.Timeout = 150000

function Update-FolderWithPersonalTag ($ExSvc, $Mailbox, $Folder) {

    $FolderId = if ([WellKnownFolderName].GetEnumNames() -contains $Folder) { $Folder } else {

        (Get-Folder -ExSvc $ExSvc -Mailbox $Mailbox -FolderDisplayName $Folder).Id
    }

    #PR_POLICY_TAG 0x3019
    $PRPolicyTag = [ExtendedPropertyDefinition]::new(0x3019, [MapiPropertyType]::Binary)

    #PR_RETENTION_FLAGS 0x301D
    $PRRetentionFlags = [ExtendedPropertyDefinition]::new(0x301D, [MapiPropertyType]::Integer)

    #PR_RETENTION_PERIOD 0x301A
    $PRRetentionPeriod = [ExtendedPropertyDefinition]::new(0x301A, [MapiPropertyType]::Integer)

    #Bind to the folder and update it with retention-related extended properties:
    $BoundFolder = [Folder]::Bind($ExSvc, $FolderId)
    $BoundFolder.SetExtendedProperty($PRRetentionFlags, 137)
    $BoundFolder.SetExtendedProperty($PRRetentionPeriod, $RetentionPeriodInDaysOfTag)
    $BoundFolder.SetExtendedProperty($PRPolicyTag, $PersonalRetentionTagGuid.ToByteArray())
    $BoundFolder.Update()
}

function Get-Folder ($ExSvc, $Mailbox, $FolderDisplayName, [switch]$Archive) {

    $ParentFolder = if ($Archive) { 'ArchiveRoot' } else { ' Root' }

    $FolderView = [FolderView]::new(1)
    $FolderView.Traversal = [FolderTraversal]::Deep

    $SearchFilterCollection = [SearchFilter+SearchFilterCollection]::new([LogicalOperator]::And)
    $SearchFilterCollection.Add([SearchFilter+IsEqualTo]::new([FolderSchema]::DisplayName, $FolderDisplayName))

    $Folder = $null
    $Folder = $ExSvc.FindFolders(

        [FolderId]::new($ParentFolder, $Mailbox),
        $SearchFilterCollection,
        $FolderView
    )

    $Folder
}

Write-Debug "Start working in live EWS here"

foreach ($Mailbox in $MailboxPSMTPs) {
    foreach ($Folder in $FolderDisplayNames) {

        Update-FolderWithPersonalTag -ExSvc $ExSvc -Mailbox $Mailbox -Folder $Folder
    }
}

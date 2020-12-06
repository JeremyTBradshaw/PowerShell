<#
    .Synopsis
    Find large items in mailbox(es) using EWS Managed API 2.2.

    .Description
    Find all 'large' items using the hidden 'AllItems' search folder.  This method is much faster than enumerating
    folders and items to accomplish the same thing (faster as in hours, even days, depending on the number and size of
    mailboxes being searched).

    When using -MailboxListCSV, a logs folder will be created in the same directory as the script, and so will a CSV
    output file (even if there are no large items found).

    The account used for -Credential parameter needs to be assigned the ApplicationImpersonation RBAC role.  The
    application used for the -AccessToken parameter needs to be setup in Azure AD as an App Registration, configured
    for app-only authentication (see .Links section).

    If the AllItems search folder is not found, the script will attempt to ceate it.  For this reason, the script
    supports ShouldProcess (i.e. -WhatIf / -Confirm).  Specifying -WhatIf will forego logging, outputting to
    CSV, and creating a new AllItems search folder, if one exists.  It won't however forego searching for an existing
    AllItems folder, nor searching that (if found) for large items.  The real intent of this implementation is to
    enable Confirmation if/when a mailbox is encountered which doesn't have the AllItems folder.  There is little risk
    in creating it, but depending on which mailbox, it may not be desired.  Simply supply -Confirm:$false to avoid
    being prompted for confirmation and allow the folder to be created if it doesn't already exist.

    .Notes
    Logging/outputting CSV to a OneDrive-synced folder may result in encountering the following error:
    "System.IO.IOException: The cloud operation was not completed before the time-out period expired.".  To avoid this,
    either place the script in a non-OneDrive-synced folder, or pause OneDrive syncing while the script is running.
    This is not a concern when using -MailboxSmtpAddress, which foregoes logging and only outputs to the host.

    .Parameter AccessToken
    Specifies an access token object (e.g. from New-EwsAccessToken (EwsOAuthAppOnlyEssentials PS module)) for the
    Azure AD application/app registration to be used for connecting to EWS using OAuth.

    .Parameter Credential
    Specifies a PSCredential object for the account to be used for connecting to EWS using Basic Authentication.

    .Parameter EwsManagedApiDllPath
    Specifies the path the the Exchange.WebServices.dll file.  Requires product/file version 15.00.0913.015.
    Defaults to 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll', and doesn't
    try to verify otherwise that the installable EWS Managed API has been installed, it just needs access to the single
    DLL file, wherever it may be.

    .Parameter LargeItemSizeMB
    Sets the value from which you will measure items against (in MB).
    Default is 150MB, the threshold for Hybrid/Remove Moves migrations to/from/between Exchange Online tenants (as of
    December 2020).

    .Parameter MailboxListCSV
    Specifies the source CSV file containing mailboxes to search through. There must be an "SmtpAddress" column header.

    .Parameter MailboxSmtpAddress
    Specifies one or more mailboxes (by SMTP address (primary/aliases) to search.

    .Parameter Archive
    Indicates to search the archive mailbox rather than the primary.

    .Parameter EwsUrl
    Specifies the URL for the Exchange Web Services endpoint.  Required when using -Credential paramter (i.e.
    Basic authentication) and regardless of whether connecting to Exchange on-premises or Exchange Online.  If using
    -AccessToken (i.e. OAuth), Exchange Online's EWS URL is automatically used instead.

    .Example
    $EwsToken = New-EwsAccessToken -TenantId 832c3217-760c-4d87-9386-efcbb4a965e5 `
                                   -ApplicationId 40d8fc2b-c0e6-4b7b-9234-d377a64e86ed `
                                   -CertificateStorePath Cert:\CurrentUser\My\51258EAF3F6EC72A7E412B239FFF39A3159D59CD

    # Get an EWS OAuth access token (New-EwsAccessToken is available in PS Module 'EwsOAuthAppOnlyEssentials').

    .Example
    .\Get-MailboxLargeItems.ps1 -AccessToken $EwsToken -MailboxSmtpAddress Larry.Iceberg@jb365.ca

    # Search a single mailbox for large items (150MB+ (default)), using OAuth (exclusively with EXO).

    .Example
    .\Get-MailboxLargeItems.ps1 -EwsUrl https://mail.contoso.com/ews/exchange.asmx -Credential $Creds -MailboxSmtpAddress Larry.Iceberg@jb365.ca -LargeItemSizeMB 10

    # Search a single mailbox for large items (10MB+).

    .Example
    .\Get-MailboxLargeItems.ps1 -EwsUrl https://mail.contoso.com/ews/exchange.asmx -Credential $Creds -MailboxSmtpAddress Larry.Iceberg@jb365.ca -Archive

    # Search a single *archive* mailbox for large items (150MB+ (default)).

    .Example
    .\Get-MailboxLargeItems.ps1 -EwsUrl https://mail.contoso.com/ews/exchange.asmx -Credential $Creds -MailboxListCsv .\Users.csv

    # Create CSV report of mailboxes with large items (150MB+ (default)).

    .Example
    .\Get-MailboxLargeItems.ps1 -EwsUrl https://mail.contoso.com/ews/exchange.asmx -Credential $Creds -MailboxListCsv .\Users.csv -Archive

    # Create CSV report of *archive* mailboxes with large items (150MB+ (default)).

    .Inputs
    # Sample CSV file for use with -MailboxListCSV parameter:
    Users.csv:
        "SmtpAddress"
        "Larry.Iceberg@jb365.ca"
        "Louis.Isaacson@jb365.ca"
        "Levy.Ingram@jb365.ca"

    .Outputs
    # Output object (will be exported to CSV when using -MailboxListCSV, only to the console when using -MailboxSmtpAddress):
        Mailbox         : Larry.Iceberg@jb365.ca
        MailboxLocation : Primary Mailbox
        ItemClass       : IPM.Note
        Subject         : My favorite photos from 2020.
        SizeMB          : 420
        DateTimeSent    : 11/30/2020 2:39:58 PM
        FolderPath      : Inbox\Personal Emails

    # Sample output CSV (when using -MailboxListCSV):
    MailboxLargeItems_2020-11-30_14-14-30.csv:
        "Mailbox","MailboxLocation","ItemClass","Subject","SizeMB","DateTimeSent","FolderPath"
        "Larry.Iceberg@jb365.ca","Primary Mailbox","IPM.Note","Company Policy PDF downloads","190","1/22/2014 10:52:39 AM","Sent Items"
        "Larry.Iceberg@jb365.ca","Primary Mailbox","IPM.Note","Photo Album 1999,"190","1/22/2014 10:49:14 AM","Sent Items"
        "Larry.Iceberg@jb365.ca","Primary Mailbox","IPM.Note","Meeting Notes (with attachments)","262","10/12/2012 1:25:52 PM","Sent Items"
        "Larry.Iceberg@jb365.ca","Primary Mailbox","IPM.Note","Wedding Pictures","232","7/8/2013 8:49:27 AM","Inbox"
        "Larry.Iceberg@jb365.ca","Primary Mailbox","IPM.Note","Study Guide for Exam 123","174","5/4/2012 3:16:22 PM","Inbox\Personal\Study"
        "Louis.Isaacson@jb365.ca","Primary Mailbox","IPM.Note","Company Event Photos","232","7/8/2013 8:49:27 AM","Deleted Items\1998"
        "Louis.Isaacson@jb365.ca","Primary Mailbox","IPM.Note","Movie Collection 2020","192","6/3/2013 11:32:05 AM","Inbox"
        "Louis.Isaacson@jb365.ca","Primary Mailbox","IPM.Note","Meeting notes (with attachments)","192","6/3/2013 11:54:54 AM","Sent Items"
        "Louis.Isaacson@jb365.ca","Primary Mailbox","IPM.Note","Wedding pictures","192","6/3/2013 11:32:05 AM","Sent Items"
        "Louis.Isaacson@jb365.ca","Primary Mailbox","IPM.Note","Phone screenshots (attached)","192","6/3/2013 11:07:01 AM","Inbox"

    # Sample log file:
    Get-MailboxLargeItems_2020-11-30_14-14-30.log:
        [ 2020-12-03 02:14:30 PM ] Get-MailoxLargeItems.ps1 - Script begin.
        [ 2020-12-03 02:14:30 PM ] PSScriptRoot: C:\Users\ExAdmin123\
        [ 2020-12-03 02:14:30 PM ] Command: .\Get-MailboxLargeItems.ps1 -AccessToken $EwsToken -MailboxListCSV .\Desktop\users.csv
        [ 2020-12-03 02:14:30 PM ] Authentication: OAuth (Exchange Online)
        [ 2020-12-03 02:14:30 PM ] LargeItemsSizeMB set to 150 MB.
        [ 2020-12-03 02:14:31 PM ] Searching Primary mailboxes (-Archive switch parameter was not used).
        [ 2020-12-03 02:14:31 PM ] Successfully imported mailbox list CSV '.\Desktop\users.csv'.
        [ 2020-12-03 02:14:31 PM ] Will process 420 mailboxes.
        [ 2020-12-03 02:14:31 PM ] Created (empty shell) output CSV file (to ensure it's avaiable for Export-Csv of any larged items that are found).
        [ 2020-12-03 02:14:31 PM ] Output CSV: C:\Users\ExAdmin123\Get-MailboxLargeItems_Outputs\MailboxLargeItems_2020-11-30_14-14-30.csv
        [ 2020-12-03 02:14:31 PM ] Successfully verified version and imported EWS Managed API 2.2 DLL (with Import-Module).
        [ 2020-11-30 02:14:31 PM ] Mailbox: 1 of 420
        [ 2020-11-30 02:14:32 PM ] Mailbox: Larry.Iceberg@jb365.ca | Found 'AllItems' search folder.  Searching it...
        [ 2020-11-30 02:14:32 PM ] Mailbox: Larry.Iceberg@jb365.ca | Found 5 large items.
        [ 2020-11-30 02:14:32 PM ] Mailbox: Larry.Iceberg@jb365.ca | Writing large items to output CSV.
        [ 2020-11-30 02:14:53 PM ] Mailbox: 2 of 420
        [ 2020-11-30 02:14:53 PM ] Mailbox: Louis.Isaacson@jb365.ca | Found 'AllItems' search folder.  Searching it...
        [ 2020-11-30 02:14:31 PM ] Mailbox: Louis.Isaacson@jb365.ca | Found 27 large items.
        [ 2020-11-30 02:14:52 PM ] Mailbox: Louis.Isaacson@jb365.ca | Writing large items to output CSV.
        ...
        ...
        [ 2020-11-22 04:20:00 PM ] Get-MailoxLargeItems.ps1 - Script end.

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxLargeItems.ps1

    .Link
    https://www.microsoft.com/en-us/download/details.aspx?id=42951 (EWS Managed API 2.2 download)

    .Link
    https://github.com/JeremyTBradshaw/EwsOAuthAppOnlyEssentials (PS module for easy access tokens)

    .Link
    https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth
#>
#Requires -Version 5.1 -PSEdition Desktop
using namespace System.Management.Automation
using namespace Microsoft.Exchange.WebServices.Data

[CmdletBinding(
    DefaultParameterSetName = 'OAuth_SmtpAddress',
    SupportsShouldProcess = $true,
    ConfirmImpact = 'High'
)]
param(
    [Parameter(Mandatory, ParameterSetName = 'OAuth_SmtpAddress')]
    [Parameter(Mandatory, ParameterSetName = 'OAuth_CSV')]
    [ValidateScript(
        {
            if ($_.token_type -eq 'Bearer' -and $_.access_token -match '^[-\w]+\.[-\w]+\.[-\w]+$') { $true } else {

                throw 'Invalid access token.  For best results, supply $AccessToken where: $AccessToken = New-EwsAccessToken ...'
            }
        }
    )]
    [Object]$AccessToken,

    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_SmtpAddress')]
    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_CSV')]
    [PSCredential]$Credential,

    [switch]$UseImpersonation,

    [ValidateScript(
        {
            if (Test-Path -Path $_) { $true } else {

                throw "Could not find EWS Managed API 2.2 DLL file $($_)"
            }
        }
    )]
    [System.IO.FileInfo]$EwsManagedApiDllPath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll',

    [ValidateRange(1, 999)]
    [int16]$LargeItemSizeMB = 150,

    [Parameter(Mandatory, ParameterSetName = 'OAuth_CSV')]
    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_CSV')]
    [ValidateScript(
        {
            if (
                (Test-Path -Path $_) -and
                ((Get-Content $_ -First 1) -replace '"' -replace "'" -match '(^SmtpAddress$)') -and
                ((Get-Content $_ -First 2).Count -eq 2)
            ) {
                $true
            }
            else { throw "CSV file failed validation.  Ensure the path is valid, there is an 'SmtpAddress' column header, and there is at least one entry/line not including the header." }
        }
    )]
    [System.IO.FileInfo]$MailboxListCSV,

    [Parameter(Mandatory, ParameterSetName = 'OAuth_SmtpAddress')]
    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_SmtpAddress')]
    [ValidatePattern('(^.*\@.*\..*$)')]
    [string[]]$MailboxSmtpAddress,

    [switch]$Archive,

    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_SmtpAddress')]
    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_CSV')]
    [uri]$EwsUrl
)

if (-not $UseImpersonation) {

    $RequireImpersonation = $false

    if ($PSCmdlet.ParameterSetName -like '*_CSV' -and (Get-Content -Path $MailboxListCSV -First 3).Count -eq 3) {

        $RequireImpersonation = $true
    }
    elseif ($MailboxSmtpAddress.Count -gt 1) {

        $RequireImpersonation = $true
    }
    if ($RequireImpersonation) {

        throw 'To process more than one mailbox, use the -UseImpersonation switch.' 
    }
}

#region Functions
function writeLog {
    param(
        [Parameter(Mandatory)]
        [string]$LogName,

        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [System.IO.FileInfo]$Folder,

        [ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
        [datetime]$LogDateTime = [datetime]::Now,

        [switch]$DisableLogging
    )

    if (-not $DisableLogging -and -not $WhatIfPreference.IsPresent) {

        # Check for current log file and if necessary create it.
        $LogFile = Join-Path -Path $Folder -ChildPath "$($LogName)_$($LogDateTime.ToString('yyyy-MM-dd_HH-mm-ss')).log"
        if (-not (Test-Path $LogFile)) {
            try {
                [void](New-Item -Path $LogFile -ItemType:File -ErrorAction Stop)
            }
            catch {
                throw "Unable to create log file $($LogFile).  Unable to write to log."
            }
        }

        [string]$Date = Get-Date -Format 'yyyy-MM-dd hh:mm:ss tt'

        # Write message to log file:
        $MessageText = "[ $($Date) ] $($Message)"
        switch ($SectionStart) {

            $true { $MessageText = "`r`n" + $MessageText }
        }
        $MessageText | Out-File -FilePath $LogFile -Append -Encoding UTF8

        # If an error was supplied, write it to the log as well.
        if ($PSBoundParameters.ErrorRecord) {

            # Format the error as it would be displayed in the PS console.
            $ErrorForLog = "$($ErrorRecord.Exception)`r`n" +
            "$($ErrorRecord.InvocationInfo.PositionMessage)`r`n" +
            "`t+ CategoryInfo: " +
            "$($ErrorRecord.CategoryInfo.Category): " +
            "($($ErrorRecord.CategoryInfo.TargetName):$($ErrorRecord.CategoryInfo.TargetType))" +
            "[$($ErrorRecord.CategoryInfo.Activity)], " +
            "$($ErrorRecord.CategoryInfo.Reason)`r`n" +
            "`t+ FullyQualifiedErrorId: $($ErrorRecord.FullyQualifiedErrorId)"

            "[ $($Date) ][Error] $($ErrorForLog)" | Out-File -FilePath $LogFile -Append
        }
    }
}

function New-EwsBinding ($AccessToken, $Url, [PSCredential]$Credential, $Mailbox) {

    # Going with Exchange2010_SP1 because it is the earliest version of the EWS schema that does what we need, per:
    # https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/ews-schema-versions-in-exchange#designing-your-application-with-schema-version-in-mind
    $ExSvc = [ExchangeService]::new(

        [ExchangeVersion]::Exchange2010_SP1
    )

    if ($PSCmdlet.ParameterSetName -like 'OAuth*') {

        $ExSvc.Credentials = [OAuthCredentials]::new($AccessToken.access_token)
        $ExSvc.Url = 'https://outlook.office365.com/ews/exchange.asmx'
    }
    else {
        $ExSvc.Credentials = [System.Net.NetworkCredential]($Credential)
        $ExSvc.Url = $Url.AbsoluteUri
    }

    if ($Script:UseImpersonation) {

        $ExSvc.ImpersonatedUserId = [ImpersonatedUserId]::new(

            [ConnectingIdType]::SmtpAddress, $Mailbox
        )

        # https://docs.microsoft.com/en-us/archive/blogs/webdav_101/best-practices-ews-authentication-and-access-issues
        $ExSvc.HttpHeaders['X-AnchorMailbox'] = $Mailbox
    }

    $ExSvc.UserAgent = 'Get-MailboxLargeItems.ps1'

    # Increase the timeout by 50% (default is 100,000) to cater to large mailboxes:
    $ExSvc.Timeout = 150000
    $ExSvc
}

function Get-AllItemsSearchFolder ($ExSvc, $Mailbox, [switch]$Archive) {

    $AllItemsParentFolder = if ($Archive) { 'ArchiveRoot' } else { ' Root' }

    $FolderView = [FolderView]::new(1)
    $FolderView.Traversal = [FolderTraversal]::Shallow

    $SearchFilterCollection = [SearchFilter+SearchFilterCollection]::new([LogicalOperator]::And)
    $SearchFilterCollection.Add([SearchFilter+IsEqualTo]::new([FolderSchema]::DisplayName, 'AllItems'))
    $SearchFilterCollection.Add([SearchFilter+IsEqualTo]::new(

            # ExtendedPropertyDefinition for MAPI property PR_FOLDER_TYPE:
            [ExtendedPropertyDefinition]::new(
                13825, #<--: Tag (Int32 value)
                [MapiPropertyType]::Integer
            ),
            2 #<--: PR_FOLDER_TYPE = 2 (a.k.a. Search Folder (vs regular folder))
        ))

    $AllItemsSearchFolder = $null
    $AllItemsSearchFolder = $ExSvc.FindFolders(

        [FolderId]::new($AllItemsParentFolder, $Mailbox),
        $SearchFilterCollection,
        $FolderView
    )

    $AllItemsSearchFolder
}

function New-AllItemsSearchFolder ($ExSvc, $Mailbox, [switch]$Archive) {

    $AllItemsParentFolder = if ($Archive) { 'ArchiveRoot' } else { 'Root' }
    $SearhRootFolder = if ($Archive) { 'ArchiveMsgFolderRoot' } else { 'MsgFolderRoot' }

    $AllItemsSearchFolder = [SearchFolder]::new($ExSvc)
    $AllItemsSearchFolder.SearchParameters.Traversal = [SearchFolderTraversal]::Deep
    $AllItemsSearchFolder.SearchParameters.SearchFilter = [SearchFilter+Exists]([ItemSchema]::ItemClass)
    $AllItemsSearchFolder.SearchParameters.RootFolderIds.Add($SearhRootFolder)
    $AllItemsSearchFolder.DisplayName = 'AllItems'

    $AllItemsSearchFolder.Save(

        [FolderId]::new($AllItemsParentFolder, $Mailbox)
    )
}

function Get-LargeItems ($ExSvc, $Mailbox, $FolderId, $LargeItemSizeMB, [switch]$Archive) {

    $ProgressParams = @{

        Id              = 1
        ParentId        = 0
        Activity        = 'Get-LargeItems'
        Status          = 'Searching all items'
        PercentComplete = -1
    }
    Write-Progress @ProgressParams

    $SearchFilter = [SearchFilter+IsGreaterThanOrEqualTo]::new(

        [ItemSchema]::Size, ($LargeItemSizeMB * 1KB)
    )

    $LargeItems = @()
    $PageSize = 1000
    $Offset = 0
    $MoreAvailable = $true

    do {
        $ItemView = [ItemView]::new(

            $PageSize,
            $Offset,
            [OffsetBasePoint]::Beginning
        )
        $ItemView.PropertySet = [PropertySet]::new(

            [BasePropertySet]::IdOnly,
            [ItemSchema]::ParentFolderId,
            [ItemSchema]::Subject,
            [ItemSchema]::Size,
            [ItemSchema]::DateTimeSent,
            [ItemSchema]::DateTimeCreated,
            [ItemSchema]::ItemClass
        )

        $FindItemsResult = $ExSvc.FindItems($FolderId, $SearchFilter, $ItemView)
        $LargeItems += $FindItemsResult.Items

        if ($FindItemsResult.MoreAvailable) {

            $Offset = $FindItemsResult.NextPageOffset 
        }
        else { $MoreAvailable = $false }
    }
    while ($MoreAvailable)

    $ItemCounter = 0
    $KnownFolderPaths = @{}
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($Item in $LargeItems) {

        $ItemCounter++
        if ($Stopwatch.Elapsed.Milliseconds -ge 300) {

            $ProgressParams['CurrentOperation'] = "Processing $($LargeItems.Count) large items"
            $ProgressParams['PercentComplete'] = (($ItemCounter / $LargeItems.Count) * 100)
            Write-Progress @ProgressParams
            $Stopwatch.Restart()
        }

        $FolderPath = $null

        if ($KnownFolderPaths.ContainsKey($Item.ParentFolderId.UniqueId)) {

            $FolderPath = $KnownFolderPaths["$($Item.ParentFolderId.UniqueId)"]
        }
        else {
            $FolderPath = Get-FolderPath -ExSvc $ExSvc -FolderId $Item.ParentFolderId -Archive:$Archive
            $KnownFolderPaths["$($Item.ParentFolderId.UniqueId)"] = $FolderPath
        }

        [PSCustomObject]@{

            Mailbox         = $Mailbox
            MailboxLocation = if ($Archive) { 'Archive Mailbox' } else { 'Primary Mailbox' }
            ItemClass       = $Item.ItemClass
            Subject         = $Item.Subject
            SizeMB          = [math]::Ceiling($Item.Size / 1MB)
            DateTimeSent    = $Item.DateTimeSent
            FolderPath      = $FolderPath
        }
    }
}

function Get-FolderPath ($ExSvc, $FolderId, [switch]$Archive) {

    $MsgFolderRoot = if ($Archive) { 'ArchiveMsgFolderRoot' } else { 'MsgFolderRoot' }

    $TopOfInformationStore = [Folder]::Bind($ExSvc, $MsgFolderRoot)
    $FolderPath = @()
    $nextFolderId = $FolderId

    do {
        $thisFolder = $null
        $thisFolder = [Folder]::Bind($ExSvc, $nextFolderId)
        $FolderPath += $thisFolder.DisplayName
        $nextFolderId = $thisFolder.ParentFolderId
    }
    while ($nextFolderId -ne $TopOfInformationStore.Id)

    [System.Array]::Reverse($FolderPath)
    $FolderPath -join '\' -replace 'Top of Information Store'
}

function Get-OAuthUserSmtpAddress ($ExSvc) {

    $ExSvc.ConvertId(
        [AlternateId]::New(
            'EwsId',
            ([Folder]::Bind($ExSvc, 'Inbox')).Id.UniqueId, 
            'OAuthUserSmtpFinder@LargeItems.ps1'
        ),
        'EwsId'

    ).Mailbox 
}
#endregion Functions

#region Main Script
try {
    #region Initialization
    $dtNow = [datetime]::Now

    $writeLogParams = @{

        LogName     = "$($MyInvocation.MyCommand.Name -replace '\.ps1')"
        Folder      = "$($PSScriptRoot)\$($MyInvocation.MyCommand.Name -replace '\.ps1')_Outputs"
        LogDateTime = $dtNow
        ErrorAction = 'Stop'
    }

    $Mailboxes = @()

    if ($PSCmdlet.ParameterSetName -like '*_CSV') {

        # Check for and if necessary create logs folder:
        if (-not (Test-Path -Path "$($writeLogParams['Folder'])")) {

            [void](New-Item -Path "$($writeLogParams['Folder'])" -ItemType Directory -ErrorAction Stop)
        }

        writeLog @writeLogParams -Message 'Get-MailoxLargeItems.ps1 - Script begin.'
        writeLog @writeLogParams -Message "PSScriptRoot: $($PSScriptRoot)"
        writeLog @writeLogParams -Message "Command: $($PSCmdlet.MyInvocation.Line)"

        if ($PSCmdlet.ParameterSetName -like 'OAuth*') {

            writeLog @writeLogParams -Message 'Authentication: OAuth (Exchange Online)'
        }
        else {
            writeLog @writeLogParams -Message "Authentication: Basic ($($Credential.UserName))"
            writeLog @writeLogParams -Message "EWS URL: $($EwsUrl)"
        }

        writeLog @writeLogParams -Message "LargeItemsSizeMB set to $($LargeItemSizeMB) MB."

        if ($PSBoundParameters.ContainsKey('Archive')) {

            writeLog @writeLogParams -Message 'Searching Archive mailboxes (-Archive switch parameter was used).'
        }
        else {
            writeLog @writeLogParams -Message 'Searching Primary mailboxes (-Archive switch parameter was not used).'
        }

        $Mailboxes += Import-Csv $MailboxListCSV -ErrorAction Stop

        writeLog @writeLogParams -Message "Successfully imported mailbox list CSV '$($MailboxListCSV)'."
        writeLog @writeLogParams -Message "Will process $($Mailboxes.Count) mailboxes."

        $OutputCSV = "$($writeLogParams['Folder'])\MailboxLargeItems_$($dtNow.ToString('yyyy-MM-dd_HH-mm-ss')).csv"
        [void](New-Item -Path $OutputCSV -ItemType File -ErrorAction Stop)

        writeLog @writeLogParams "Created (empty shell) output CSV file (to ensure it's avaiable for Export-Csv of any larged items that are found)."
        writeLog @writeLogParams "Output CSV: $($OutputCSV)"
    }
    else {
        # Disable logging and forego output CSV.
        $writeLogParams['DisableLogging'] = $true

        foreach ($sA in $MailboxSmtpAddress) {

            $Mailboxes += [PSCustomObject]@{ SmtpAddress = $sA }
        }
    }

    $EwsManagedApiDll = Get-ChildItem -Path $EwsManagedApiDllPath -ErrorAction Stop

    if ($EwsManagedApiDll.VersionInfo.FileVersion -ne '15.00.0913.015') {

        throw "EWS Managed API 2.2 is required, specifically product/file version 15.00.0913.015.`r`n" +
        "Download: https://www.microsoft.com/en-us/download/details.aspx?id=42951"
    }
    Import-Module $EwsManagedApiDll -ErrorAction Stop

    writeLog @writeLogParams 'Successfully verified version and imported EWS Managed API 2.2 DLL (with Import-Module).'
    #endregion Initialization

    #region Mailbox Loop
    $MainProgressParams = @{

        Id       = 0
        Activity = "Get-MailboxLargeItem.ps1 (Primary Mailboxes) - Start time: $($dtNow)"
    }

    $MailboxCounter = 0

    foreach ($Mailbox in $Mailboxes.SmtpAddress) {

        $MailboxCounter++

        $MainProgressParams['Status'] = "Finding large items ($($LargeItemSizeMB)+ MB) | Mailbox $($MailboxCounter) of $($Mailboxes.Count)"
        $MainProgressParams['PercentComplete'] = (($MailboxCounter / $Mailboxes.Count) * 100)

        if ($Archive) {

            $MainProgressParams['Activity'] = $MainProgressParams['Activity'] -replace 'Primary', 'Archive'
        }

        Write-Progress @MainProgressParams

        try {
            writeLog @writeLogParams -Message "Mailbox: $($MailboxCounter) of $($Mailboxes.Count)"

            $ExSvcParams = @{ Mailbox = $Mailbox }

            if ($PSCmdlet.ParameterSetName -like 'OAuth*') {

                $ExSvcParams['AccessToken'] = $AccessToken
            }
            else {
                $ExSvcParams['Url'] = $EwsUrl
                $ExSvcParams['Credential'] = $Credential
            }

            $ExSvc = New-EwsBinding @ExSvcParams

            # In case the supplied SmtpAddress is not that of the actual OAuth-authenticated user:
            if (-not $UseImpersonation -and $PSCmdlet.ParameterSetName -like 'OAuth*') {

                $Mailbox = Get-OAuthUserSmtpAddress -ExSvc $ExSvc
            }

            $MainProgressParams['CurrentOperation'] = "Current mailbox: $($Mailbox)"

            $AllItemsSearchFolder = $null
            $AllItemsSearchFolder = Get-AllItemsSearchFolder -ExSvc $ExSvc -Mailbox $Mailbox -Archive:$Archive

            if (-not $AllItemsSearchFolder) {

                $currentMsg = $null
                $currentMsg = "Mailbox: $($Mailbox) | No 'AllItems' hidden search folder found."
                writeLog @writeLogParams -Message $currentMsg

                if ($PSCmdlet.ShouldProcess(

                        "Mailbox: $($Mailbox) | Creating new 'AllItems' hidden search folder.",
                        "Are you sure you want to create a new 'AllItems' hidden search folder?",
                        $currentMsg
                    )) {
                    [void](New-AllItemsSearchFolder -ExSvc $ExSvc -Mailbox $Mailbox -Archive:$Archive)

                    writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Created new 'AllItems' hidden search folder."

                    Start-Sleep -Seconds 3 #<--: Not expected often so 3 seconds is acceptable.

                    $AllItemsSearchFolder = Get-AllItemsSearchFolder -ExSvc $ExSvc -Mailbox $Mailbox -Archive:$Archive

                    if (-not $AllItemsSearchFolder) { throw 90210 }
                }
            }

            if ($AllItemsSearchFolder) {

                writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Found 'AllItems' search folder.  Searching it..."

                $getLargeItemsParams = @{

                    ExSvc           = $ExSvc
                    Mailbox         = $Mailbox
                    FolderId        = $AllItemsSearchFolder.Id
                    LargeItemSizeMB = $LargeItemSizeMB
                    Archive         = $Archive
                }
                $LargeItems = @()
                $LargeItems += Get-LargeItems @getLargeItemsParams

                writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Found $($LargeItems.Count) large items."

                if ($LargeItems.Count -ge 1) {

                    writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Writing large items to output CSV."

                    if ($PSCmdlet.ParameterSetName -like '*_CSV') {

                        $LargeItems | Export-Csv -Path $OutputCSV -Append -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
                    }
                    else { $LargeItems }
                }
            }
        }
        catch {
            if ($_ -match '(90210)') {
                $currentMsg = "Mailbox $($Mailbox) | Newly created 'AllItems' folder is still not availalbe.  Try again later."

                Write-Warning -Message $currentMsg
                writeLog @writeLogParams -Message $currentMsg
            }
            elseif (
                ($_.ToString().Contains('The SMTP address has no mailbox associated with it.')) -or
                ($_.Exception.InnerException -match 'No mailbox with such guid.') -or
                ((-not $Archive) -and $_.Exception.InnerException -match '(The element at position 0 is invalid.*\nParameter name: parentFolderIds)')
            ) {
                $currentMsg = $null
                $currentMsg = "Mailbox: $($Mailbox) | A mailbox was not found for this user."

                Write-Warning -Message $currentMsg
                writeLog @writeLogParams -Message $currentMsg
            }
            elseif (
                ($PSBoundParameters.ContainsKey('Archive')) -and
                (
                    ($_.ToString().Contains('The specified folder could not be found in the store.')) -or
                    ($_.Exception.InnerException -match '(The element at position 0 is invalid.*\nParameter name: parentFolderIds)')
                )
            ) {
                $currentMsg = $null
                $currentMsg = "Mailbox: $($Mailbox) | There is no archive mailbox for this user."

                Write-Warning -Message $currentMsg
                writeLog @writeLogParams -Message $currentMsg
            }
            elseif (
                ($PSBoundParameters.ContainsKey('Archive')) -and
                ($_.Exception.InnerException -match "The user's remote archive is disabled.")
            ) {
                $currentMsg = $null
                $currentMsg = "Mailbox: $($Mailbox) | There is no local archive mailbox for this user, although there may be one in EXO."

                Write-Warning -Message $currentMsg
                writeLog @writeLogParams -Message $currentMsg
            }
            else {
                $currentMsg = $null
                $currentMsg = "Mailbox $($Mailbox) | Script-ending error: $($_.Exception.Message)"

                Write-Warning -Message $currentMsg
                writeLog @writeLogParams -Message $currentMsg -ErrorRecord $_
                Write-Error $_
                break
            }
            Write-Debug -Message 'Suspend the script here to investigate the error. Otherwise, unless Halted, the script will continue.'
        }	
    }
    #endregion Mailbox Loop
}
catch {
    $currentMsg = $null
    $currentMsg = "Script-ending failure: $($_.Exception.Message)"

    Write-Warning -Message $currentMsg
    writeLog @writeLogParams -Message $currentMsg
    throw $_
}

finally { writeLog @writeLogParams -Message 'Get-MailoxLargeItems.ps1 - Script end.' }
#endRegion Main Script

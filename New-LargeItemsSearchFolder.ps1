<#
    .Synopsis
    Create 'Large Items (###MB+)' Search Folder in mailboxes using EWS Managed API 2.2.

    .Description
    Create 'Large Items (###MB+)' Search Folder in mailboxes to help users find their large items.  This task commonly
    comes up when planning a migration to/from/between Exchange Online tenants.  A good idea is to use the sibling
    script - Get-MailboxLargeItems.ps1 - to determine which mailboxes have large items in them.  Alternatively, just
    provide users with this new search folder and advise them to backup the items it finds so they don't lose them
    during their migration.

    When using -MailboxListCSV, a logs folder will be created in the same directory as the script, and so will a CSV
    output file (even if there are no large items found).  When using either -MailboxListCSV or -MailboxSmtpAddress
    parameters, impersonation is implied and the account used for -Credential parameter needs to be assigned the
    ApplicationImpersonation RBAC role, at least for the scope of the mailboxes being searched.  Similarly, If
    -MailboxListCSV or -MailboxSmtpAddress are used with the -AccessToken parameter, the application used for
    the -AccessToken parameter needs to be setup in Azure AD as an App Registration, and, if the access token is an
    App-Only token, the application must be configured for app-only authentication (see .Links section), or if the 
    token is a delegated token, the user of the token must have the ApplicationImpersonation RBAC role assigned, at
    least for the scope of the mailboxes being searched.

    .Parameter AccessToken
    Specifies an access token object (e.g. from New-EwsAccessToken (EwsOAuthAppOnlyEssentials PS module)) for the
    Azure AD application/app registration to be used for connecting to EWS using OAuth.  Delegated OAuth tokens are
    supported and will be documented here in more detail soon.

    .Parameter Credential
    Specifies a PSCredential object for the account to be used for connecting to EWS using Basic Authentication.

    .Parameter EwsManagedApiDllPath
    Specifies the path the the Exchange.WebServices.dll file.  Requires product/file version 15.00.0913.015.
    Defaults to 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll', and doesn't
    try to verify otherwise that the installable EWS Managed API has been installed, it just needs access to the single
    DLL file, wherever it may be.

    .Parameter LargeItemSizeMB
    Sets the size (in MB) of the large items to search for (in the search folder's criteria).
    Default is 150MB, the threshold for Hybrid/Remove Moves migrations to/from/between Exchange Online tenants (as of
    December 2020).

    .Parameter MailboxListCSV
    Specifies the source CSV file containing mailboxes to create the new Search Folder in. There must be an "SmtpAddress" column header.

    .Parameter MailboxSmtpAddress
    Specifies one or more mailboxes (by SMTP address (primary/aliases) to create the search folder in.

    .Parameter Archive
    Indicates to create the Search Folder in the archive mailbox (if one exists), rather than the primary mailbox.

    .Parameter EwsUrl
    Specifies the URL for the Exchange Web Services endpoint.  Required when using -Credential paramter (i.e.
    Basic authentication) and regardless of whether connecting to Exchange on-premises or Exchange Online.  If using
    -AccessToken (i.e. OAuth), Exchange Online's EWS URL is automatically used instead.

    .Example
    .\New-LargeItemsSearchFolder.ps1 -AccessToken $DelegatedAccessToken

    # Create large items search folder in the user of the delegated access token's mailbox.

    .Example
    $EwsToken = New-EwsAccessToken -TenantId 832c3217-760c-4d87-9386-efcbb4a965e5 `
                                   -ApplicationId 40d8fc2b-c0e6-4b7b-9234-d377a64e86ed `
                                   -CertificateStorePath Cert:\CurrentUser\My\51258EAF3F6EC72A7E412B239FFF39A3159D59CD

    # Get an EWS OAuth access token (New-EwsAccessToken is available in PS Module 'EwsOAuthAppOnlyEssentials').

    .Example
    New-LargeItemsSearchFolder.ps1 -EwsUrl https://ex2016.contoso.com/ews/exchange.asmx -Credential <PSCredential> -MailboxListCSV .\LIUsers.csv

    # Create large items search folder, in mailboxes supplied in the CSV file.

    .Example
    New-LargeItemsSearchFolder.ps1 -EwsUrl https://ex2016.contoso.com/ews/exchange.asmx -Credential <PSCredential> -MailboxListCSV .\LIUsers.csv -LargeItemSizeMB <Value in MB>

    # Create large items search folder, in mailboxes supplied in the CSV file (using a large item definition other than 150MB).

    .Example
    New-LargeItemsSearchFolder.ps1 -EwsUrl https://ex2016.contoso.com/ews/exchange.asmx -Credential <PSCredential> -LargeItemSizeMB <Value in MB> -Archive -MailboxListCSV .\LIUsers.csv

    # Create large items search folder, both in the primary and archive mailbox, in mailboxes supplied in the CSV file.

    .Outputs
    # Sample log file (when using -MailboxListCSV):
    New-LargeItemsSearchFolder_2020-12-03_20-20-24.log:
        [ 2020-12-03 10:20:24 PM ] New-LargeItemsSearchFolder.ps1 - Script begin.
        [ 2020-12-03 10:20:24 PM ] PSScriptRoot: C:\Users\ExAdmin123
        [ 2020-12-03 10:20:24 PM ] Command: .\New-LargeItemsSearchFolder.ps1 -AccessToken $EwsToken -MailboxListCSV .\Desktop\users.csv
        [ 2020-12-03 10:20:24 PM ] Authentication: OAuth (Exchange Online)
        [ 2020-12-03 10:20:24 PM ] LargeItemsSizeMB set to 150 MB.
        [ 2020-12-03 10:20:24 PM ] Targeting Primary mailboxes (-Archive switch parameter was not used).
        [ 2020-12-03 10:20:24 PM ] Successfully imported mailbox list CSV '.\Desktop\users.csv'.
        [ 2020-12-03 10:20:24 PM ] Will process 4 mailboxes.
        [ 2020-12-03 10:20:24 PM ] Successfully verified version and imported EWS Managed API 2.2 DLL (with Import-Module).
        [ 2020-12-03 10:20:24 PM ] Mailbox: 1 of 4
        [ 2020-12-03 10:20:24 PM ] Mailbox: HandledFailure@example.123 | A mailbox was not found for this user.
        [ 2020-12-03 10:20:25 PM ] Mailbox: 2 of 4
        [ 2020-12-03 10:20:25 PM ] Mailbox: Larry.Iceberg@jb365.ca | Created search folder 'Large Items (150MB+)'.
        [ 2020-12-03 10:20:25 PM ] Mailbox: 3 of 4
        [ 2020-12-03 10:20:26 PM ] Mailbox: Louis.Isaacson@jb365.ca | Created search folder 'Large Items (150MB+)'.
        [ 2020-12-03 10:20:26 PM ] Mailbox: 4 of 4
        [ 2020-12-03 10:20:26 PM ] Mailbox: Levy.Ingram@jb365.ca | Created search folder 'Large Items (150MB+)'.
        [ 2020-12-03 10:20:26 PM ] New-LargeItemsSearchFolder.ps1 - Script end.

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/edit/main/New-LargeItemsSearchFolder.ps1

    .Link
    https://www.microsoft.com/en-us/download/details.aspx?id=42951 (EWS Managed API 2.2 download)

    .Link
    https://github.com/JeremyTBradshaw/EwsOAuthAppOnlyEssentials (PS module for easy access tokens)

    .Link
    https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth
#>
#Requires -Version 5.1
using namespace System.Management.Automation
using namespace Microsoft.Exchange.WebServices.Data

[CmdletBinding(
    DefaultParameterSetName = 'Basic_NoImpersonation',
    SupportsShouldProcess = $true,
    ConfirmImpact = 'High'
)]
param(
    [Parameter(Mandatory, ParameterSetName = 'OAuth_NoImpersonation')]
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

    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_NoImpersonation')]
    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_SmtpAddress')]
    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_CSV')]
    [PSCredential]$Credential,

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

    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_NoImpersonation')]
    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_SmtpAddress')]
    [Parameter(Mandatory, ParameterSetName = 'BasicAuth_CSV')]
    [ValidateScript(
        {
            if ($_.AbsoluteUri) { $true } else { throw "$($_) is not a valid URL." }
        }
    )]
    [uri]$EwsUrl
)

#region Functions
function writeLog {
    param (
        [Parameter(Mandatory)][string]$LogName,
        [Parameter(Mandatory)][System.IO.FileInfo]$Folder,
        [Parameter(Mandatory, ValueFromPipeline)][string]$Message,
        [ErrorRecord]$ErrorRecord,
        [datetime]$LogDateTime = [datetime]::Now,
        [switch]$DisableLogging,
        [switch]$SectionStart,
        [switch]$PassThru
    )

    if (-not $DisableLogging -and -not $WhatIfPreference.IsPresent) {
        try {
            if (-not (Test-Path -Path $Folder)) {

                [void](New-Item -Path $Folder -ItemType Directory -ErrorAction Stop)
            }
            $LogFile = Join-Path -Path $Folder -ChildPath "$($LogName)_$($LogDateTime.ToString('yyyy-MM-dd_HH-mm-ss')).log"
            if (-not (Test-Path $LogFile)) {

                [void](New-Item -Path $LogFile -ItemType:File -ErrorAction Stop)
            }
        }
        catch { throw $_ }

        $Date = Get-Date -Format 'yyyy-MM-dd hh:mm:ss tt'
        $MessageText = "[ $($Date) ] $($Message)"
        switch ($SectionStart) {

            $true { $MessageText = "`r`n" + $MessageText }
        }
        $MessageText | Out-File -FilePath $LogFile -Append

        if ($PSBoundParameters.ErrorRecord) {

            # Format the error as it would be displayed in the PS console.
            "[ $($Date) ][Error] $($ErrorRecord.Exception.Message)`r`n" +
            "$($ErrorRecord.InvocationInfo.PositionMessage)`r`n" +
            "`t+ CategoryInfo: $($ErrorRecord.CategoryInfo.Category): " +
            "($($ErrorRecord.CategoryInfo.TargetName):$($ErrorRecord.CategoryInfo.TargetType))" +
            "[$($ErrorRecord.CategoryInfo.Activity)], $($ErrorRecord.CategoryInfo.Reason)`r`n" +
            "`t+ FullyQualifiedErrorId: $($ErrorRecord.FullyQualifiedErrorId)`r`n" |
            Out-File -FilePath $LogFile -Append
        }
    }
    if ($PassThru) { $Message }
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

    if (-not ($PSCmdlet.ParameterSetName -like '*NoImpersonation')) {

        $ExSvc.ImpersonatedUserId = [ImpersonatedUserId]::new(

            [ConnectingIdType]::SmtpAddress, $Mailbox
        )

        # https://docs.microsoft.com/en-us/archive/blogs/webdav_101/best-practices-ews-authentication-and-access-issues
        $ExSvc.HttpHeaders['X-AnchorMailbox'] = $Mailbox
    }

    $ExSvc.UserAgent = $PSCmdlet.MyInvocation.MyCommand

    $ExSvc
}

function New-SearchFolder ($ExSvc, $LargeItemSizeMB, [switch]$Archive) {

    $SearchFilter = [SearchFilter+IsGreaterThanOrEqualTo]::new([ItemSchema]::Size, ($LargeItemSizeMB * 1MB))
    $SearchFolder = [SearchFolder]::new($ExSvc)
    $SearchFolder.DisplayName = "Large Items ($($LargeItemSizeMB)MB+)"
    $SearchFolder.SearchParameters.Traversal = 'Deep'
    $SearchFolder.SearchParameters.SearchFilter = $SearchFilter

    if ($Archive) {

        # There is no 'ArchiveSearchFolders' in [WellKnownFolderName] enum, so we need to find its Id instead.
        $ArchiveSearchFolders = $ExSvc.FindFolders(

            'ArchiveRoot',
            [SearchFilter+IsEqualTo]::new([FolderSchema]::DisplayName, 'Finder'), [FolderView]::new(1)
        )
        if ($ArchiveSearchFolders) {

            $SearchFolder.SearchParameters.RootFolderIds.Add('ArchiveMsgFolderRoot')
            $SearchFolder.Save($ArchiveSearchFolders.Folders.Id)
        }
        else {
            throw "Finder folder (a.k.a. 'Search Folders') wasn't found in the Archive mailbox."
        }
    }
    else {
        $SearchFolder.SearchParameters.RootFolderIds.Add('MsgFolderRoot')
        $SearchFolder.Save('SearchFolders')
    }
}

function Get-ConnectingUserSmtpAddress ($ExSvc) {

    $ExSvc.ConvertId(
        [AlternateId]::New(
            'EwsId',
            ([Folder]::Bind($ExSvc, 'Root')).Id.UniqueId, 
            "ConnectingUser@$($PSCmdlet.MyInvocation.MyCommand)"
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
        Folder      = "$($PSScriptRoot)\$($MyInvocation.MyCommand.Name -replace '\.ps1')_Logs"
        LogDateTime = $dtNow
        ErrorAction = 'Stop'
    }

    $Mailboxes = @()

    if ($PSCmdlet.ParameterSetName -like '*_CSV') {

        # Check for and if necessary create logs folder:
        if (-not (Test-Path -Path "$($writeLogParams['Folder'])")) {

            [void](New-Item -Path "$($writeLogParams['Folder'])" -ItemType Directory -ErrorAction Stop)
        }

        writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand) - Script begin."
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

            writeLog @writeLogParams -Message 'Targeting Archive mailboxes (-Archive switch parameter was used).'
        }
        else {
            writeLog @writeLogParams -Message 'Targeting Primary mailboxes (-Archive switch parameter was not used).'
        }

        $Mailboxes += Import-Csv $MailboxListCSV -ErrorAction Stop

        writeLog @writeLogParams -Message "Successfully imported mailbox list CSV '$($MailboxListCSV)'."
        writeLog @writeLogParams -Message "Will process $($Mailboxes.Count) mailboxes."
    }
    else {
        # Disable logging.
        $writeLogParams['DisableLogging'] = $true

        if ($PSCmdlet.ParameterSetName -like '*_SmtpAddress') {

            foreach ($sA in $MailboxSmtpAddress) {

                $Mailboxes += [PSCustomObject]@{ SmtpAddress = $sA }
            }
        }
        elseif ($PSCmdlet.ParameterSetName -like 'Basic*') {

            # Setting as placeholder.  Will determine SmtpAddress later:
            $Mailboxes += [PSCustomObject]@{ SmtpAddress = $Credential.UserName }
        }
        else {
            # Set sa placeholder SmtpAddress for the user of the -AccessToken:
            $Mailboxes += [PSCustomObject]@{ SmtpAddress = "ConnectingUser@$($PSCmdlet.MyInvocation.MyCommand)" }
        }
    }

    $EwsManagedApiDll = Get-ChildItem -Path $EwsManagedApiDllPath -ErrorAction Stop

    if ($EwsManagedApiDll.VersionInfo.FileVersion -ne '15.00.0913.015') {

        $errorMessage = "EWS Managed API 2.2 is required, specifically product/file version 15.00.0913.015.`r`n" +
        "Download: https://www.microsoft.com/en-us/download/details.aspx?id=42951"

        throw $errorMessage
    }
    Import-Module $EwsManagedApiDll -ErrorAction Stop

    writeLog @writeLogParams 'Successfully verified version and imported EWS Managed API 2.2 DLL (with Import-Module).'
    #endregion Initialization

    #region Mailbox Loop
    $MainProgressParams = @{

        Id       = 0
        Activity = "New-LargeItemsSearchFolder.ps1 - Start time: $($dtNow)"
    }

    $MailboxCounter = 0

    foreach ($Mailbox in $Mailboxes.SmtpAddress) {

        $MailboxCounter++

        $MainProgressParams['Status'] = "Creating 'Large Items ($($LargeItemSizeMB)MB+)' folder | Mailbox $($MailboxCounter) of $($Mailboxes.Count)"
        $MainProgressParams['PercentComplete'] = (($MailboxCounter / $Mailboxes.Count) * 100)

        if ($Archive) {

            $MainProgressParams['CurrentOperation'] = $MainProgressParams['CurrentOperation'] -replace 'mailbox:', 'mailbox (Archive):'
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

            # Attempt to find SmtpAddress connecting/authenticating user:
            if ($PSCmdlet.ParameterSetName -like '*NoImpersonation') {

                $Mailbox = Get-ConnectingUserSmtpAddress -ExSvc $ExSvc
            }

            $MainProgressParams['CurrentOperation'] = "Current mailbox: $($Mailbox)"

            if ($PSCmdlet.ShouldProcess(

                    "Mailbox: $($Mailbox) | Creating a new 'Large Items ($($LargeItemSizeMB)MB+)' search folder.",
                    "Are you sure you want to create a new 'Large Items ($($LargeItemSizeMB)MB+)' search folder?",
                    "Mailbox: $($Mailbox)"
                )
            ) {
                [void](New-SearchFolder -ExSvc $ExSvc -LargeItemSizeMB $LargeItemSizeMB -Archive:$Archive)
                writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Created search folder 'Large Items ($($LargeItemSizeMB)MB+)'."
            }
            else {
                writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Folder creation cancelled (via Confirm prompt)."
            }
        }
        catch {
            # Depends on PSVersion 5.1, or for PSVersions 6+ - $DebugPreference set to 'Inquire':
            $debugHelpMessage = 'Suspend the script here to investigate the error (e.g. check $ExSvc.HttpResponseHeaders). ' +
            'Otherwise, unless Halted, or if the error is script-ending, the script will continue.'

            if ($_.Exception.InnerException -match 'A folder with the specified name already exists') {

                "Mailbox: $($Mailbox) | The search folder 'Large Items ($($LargeItemSizeMB)MB+)' already exists." |
                writeLog @writeLogParams -PassThru | Write-Warning
            }
            elseif (
                ($_.ToString().Contains('The SMTP address has no mailbox associated with it.')) -or
                ($_.Exception.InnerException -match 'No mailbox with such guid\.') -or
                ((-not $Archive) -and $_.Exception.InnerException -match '(The element at position 0 is invalid.*\nParameter name: parentFolderIds)')
            ) {
                "Mailbox: $($Mailbox) | A mailbox was not found for this user." |
                writeLog @writeLogParams -PassThru | Write-Warning
            }
            elseif (
                ($PSBoundParameters.ContainsKey('Archive')) -and
                (
                    ($_.ToString().Contains('The specified folder could not be found in the store.')) -or
                    ($_.Exception.InnerException -match '(The element at position 0 is invalid.*\nParameter name: parentFolderIds)')
                )
            ) {
                "Mailbox: $($Mailbox) | There is no archive mailbox for this user." |
                writeLog @writeLogParams -PassThru | Write-Warning
            }
            elseif (
                ($PSBoundParameters.ContainsKey('Archive')) -and
                ($_.Exception.InnerException -match "The user's remote archive is disabled\.")
            ) {
                "Mailbox: $($Mailbox) | There is no local archive mailbox for this user, although there may be one in EXO." |
                writeLog @writeLogParams -PassThru | Write-Warning
            }
            elseif ($_.Exception.Message -eq "Finder folder (a.k.a. 'Search Folders') wasn't found in the Archive mailbox.") {

                "Mailbox: ($Mailbox) | Finder folder (a.k.a. 'Search Folders') wasn't found in the Archive mailbox.  " +
                'Unable to create the new search folder.' | writeLog @writeLogParams -PassThru | Write-Warning
            }
            elseif ($_.Exception.InnerException -match '(ExchangeImpersonation SOAP header must be present for this type of OAuth token\.)') {

                'The supplied access token appears to be for app-only authentication, and therefore impersonation must be used.  ' +
                'For this kind of access token, supply either -MailboxListCSV or -MailboxSmtpAddress.  ' +
                'Otherwise, provide a delegated authorization access token.' | Write-Warning
                break
            }
            else {
                "Mailbox $($Mailbox) | Script-ending error: $($_.Exception.Message)" |
                writeLog @writeLogParams -ErrorRecord $_ -PassThru | Write-Warning
                Write-Error $_
                Write-Debug -Message $debugHelpMessage
                break
            }
            Write-Debug -Message $debugHelpMessage
        }	
    }
    #endregion Mailbox Loop
}
catch {
    "Script-ending failure: $($_.Exception.Message)" |
    writeLog @writeLogParams -PassThru | Write-Warning
    Write-Debug -Message $debugHelpMessage
    throw $_
}

finally { writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand) - Script end." }
#endRegion Main Script

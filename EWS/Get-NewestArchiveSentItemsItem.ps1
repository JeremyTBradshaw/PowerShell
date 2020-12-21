<#
    .Synopsis
    Find newest item in Archive mailbox's Sent Items folder, using EWS Managed API 2.2.

    .Description
    Finding the newest item in an Archive mailbox is something that you'd think could be done using the Exchange
    PowerShell Cmdlet Get-MailboxFolderStatistics -Archive -IncludeOldestAndNewestItems.  Unfortunately, when the
    -Archive switch parameter is used, the -IncludeOldestAndNewestItems switch parameter seems to go dormant.  Since
    it was already 2020 when I realized this, and no Google/Bing search results reveal that anyone else on the planet
    cares, I figure I'll close this gap with this script instead of bothering the Exchange team to repair the Cmdlet.

    When using -MailboxListCSV, a logs folder will be created in the same directory as the script, and so will a CSV
    output file (even if there are no large items found).  When using either -MailboxListCSV or -MailboxSmtpAddress
    parameters, impersonation is implied and the account used for -Credential parameter needs to be assigned the
    ApplicationImpersonation RBAC role, at least for the scope of the mailboxes being searched.  Similarly, If
    -MailboxListCSV or -MailboxSmtpAddress are used with the -AccessToken parameter, the application used for
    the -AccessToken parameter needs to be setup in Azure AD as an App Registration, and, if the access token is an
    App-Only token, the application must be configured for app-only authentication (see .Links section), or if the
    token is a delegated token, the user of the token must have the ApplicationImpersonation RBAC role assigned, at
    least for the scope of the mailboxes being searched.

    .Notes
    Logging/outputting CSV to a OneDrive-synced folder may result in encountering the following error:
    "System.IO.IOException: The cloud operation was not completed before the time-out period expired.".  To avoid this,
    either place the script in a non-OneDrive-synced folder, or pause OneDrive syncing while the script is running.
    This is not a concern when using -MailboxSmtpAddress, which foregoes logging and only outputs to the host.

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

    .Parameter MailboxListCSV
    Specifies the source CSV file containing mailboxes to search through. There must be an "SmtpAddress" column header.

    .Parameter MailboxSmtpAddress
    Specifies one or more mailboxes (by SMTP address (primary/aliases) to search.

    .Parameter EwsUrl
    Specifies the URL for the Exchange Web Services endpoint.  Required when using -Credential paramter (i.e.
    Basic authentication) and regardless of whether connecting to Exchange on-premises or Exchange Online.  If using
    -AccessToken (i.e. OAuth), Exchange Online's EWS URL is automatically used instead.
#>
#Requires -Version 5.1 -PSEdition Desktop
using namespace System.Management.Automation
using namespace Microsoft.Exchange.WebServices.Data

[CmdletBinding(
    DefaultParameterSetName = 'BasicAuth_NoImpersonation',
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
    [CmdletBinding()]
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
                Out-File -FilePath $LogFile -Append -ErrorAction Stop
            }
        }
        catch { throw $_ }
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

function Get-ArchiveSentItemsFolder ($ExSvc, $Mailbox) {

    $FolderView = [FolderView]::new(1)
    $FolderView.Traversal = [FolderTraversal]::Shallow

    $SearchFilter = [SearchFilter+IsEqualTo]::new([FolderSchema]::DisplayName, 'Sent Items')
    $ArchiveSentItems = $null
    $ArchiveSentItems = $ExSvc.FindFolders(

        [FolderId]::new('ArchiveMsgFolderRoot', $Mailbox),
        $SearchFilter,
        $FolderView
    )

    $ArchiveSentItems
}

function Get-NewestItem ($ExSvc, $Mailbox, $FolderId) {

    $ItemView = [ItemView]::new(1)
    $ItemView.PropertySet = [PropertySet]::new(

        [BasePropertySet]::IdOnly,
        [ItemSchema]::Subject,
        [ItemSchema]::Size,
        [ItemSchema]::DateTimeSent,
        [ItemSchema]::ItemClass
    )

    $NewestItem = $ExSvc.FindItems($FolderId, $ItemView)

    [PSCustomObject]@{

        Mailbox         = $Mailbox
        ItemClass       = $NewestItem.ItemClass
        Subject         = $NewestItem.Subject
        SizeMB          = [math]::Round(($NewestItem.Size / 1MB), 2)
        DateTimeSent    = $NewestItem.DateTimeSent
    }
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

        throw "EWS Managed API 2.2 is required, specifically product/file version 15.00.0913.015.`r`n" +
        "Download: https://www.microsoft.com/en-us/download/details.aspx?id=42951"
    }
    Import-Module $EwsManagedApiDll -ErrorAction Stop

    writeLog @writeLogParams 'Successfully verified version and imported EWS Managed API 2.2 DLL (with Import-Module).'
    #endregion Initialization

    #region Mailbox Loop
    $MainProgressParams = @{

        Id       = 0
        Activity = "$($PSCmdlet.MyInvocation.MyCommand) - Start time: $($dtNow)"
    }

    $MailboxCounter = 0

    foreach ($Mailbox in $Mailboxes.SmtpAddress) {

        $MailboxCounter++

        $MainProgressParams['Status'] = "Finding newest item in (Archive) Sent Items folder | Mailbox $($MailboxCounter) of $($Mailboxes.Count)"
        $MainProgressParams['PercentComplete'] = (($MailboxCounter / $Mailboxes.Count) * 100)

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

            $ArchiveSentItemsFolder = $null
            $ArchiveSentItemsFolder = Get-ArchiveSentItemsFolder -ExSvc $ExSvc -Mailbox $Mailbox

            if ($ArchiveSentItemsFolder) {

                writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Found 'Sent Items' in root of archive mailbox.  Searching it..."

                $NewestItem = $null
                $NewestItem = Get-NewestItem -ExSvc $ExSvc -FolderId $ArchiveSentItemsFolder.Id -Mailbox $Mailbox

                if ($NewestItem) {

                    writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Found newest item."

                    if ($PSCmdlet.ParameterSetName -like '*_CSV') {

                        writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Writing item to output CSV."
                        $Newest | Export-Csv -Path $OutputCSV -Append -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
                    }
                    else { $NewestItem }
                }
                else {
                    writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | No items found."
                }
            }
        }
        catch {
            # Depends on PSVersion 5.1, or for PSVersions 6+ - $DebugPreference set to 'Inquire':
            $debugHelpMessage = 'Suspend the script here to investigate the error (e.g. check $ExSvc.HttpResponseHeaders). ' +
            'Otherwise, unless Halted, or if the error is script-ending, the script will continue.'

            if ($_ -match '(90210)') {

                "Mailbox $($Mailbox) | Newly created 'AllItems' folder is still not availalbe.  Try again later." |
                writeLog @writeLogParams -PassThru | Write-Warning
            }
            elseif (
                ($_.ToString().Contains('The SMTP address has no mailbox associated with it.')) -or
                ($_.Exception.InnerException -match 'No mailbox with such guid.') -or
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
                ($_.Exception.InnerException -match "The user's remote archive is disabled.")
            ) {
                "Mailbox: $($Mailbox) | There is no local archive mailbox for this user, although there may be one in EXO." |
                writeLog @writeLogParams -PassThru | Write-Warning
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
    "Script-ending failure: $($_.Exception.Message)" | writeLog @writeLogParams
    throw $_
}

finally { writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand) - Script end." }
#endRegion Main Script

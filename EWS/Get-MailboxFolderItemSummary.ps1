<#
    .Synopsis
    Summarize the items within a specified mailbox folder by age.

    .Notes
    I'm using Glen Scales' blog post as my guide for this script:
    https://gsexdev.blogspot.com/2013/06/ewspowershell-recoverable-items-age.html

    Particularly the parts concerning FindItems part of the process.  The rest of the script borrows my base amount of
    EWS Managed API stuff which can be seen in the rest of my EWS-focused scripts.
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

    $ExSvc = [ExchangeService]::new()

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

    # Increase the timeout by 50% (default is 100,000) to cater to large mailboxes:
    $ExSvc.Timeout = 150000
    $ExSvc
}

function Get-EwsFolder ($ExSvc, $Mailbox, $FolderDisplayName, [switch]$Archive) {

    $TargetRootFolder = if ($Archive) { 'ArchiveRoot' } else { ' Root' }

    $FolderView = [FolderView]::new(1)
    $FolderView.Traversal = [FolderTraversal]::Deep

    $SearchFilterCollection = [SearchFilter+SearchFilterCollection]::new([LogicalOperator]::And)
    $SearchFilterCollection.Add([SearchFilter+IsEqualTo]::new([FolderSchema]::DisplayName, $FolderDisplayName))

    $EwsFolder = $null
    $EwsFolder = $ExSvc.FindFolders(

        [FolderId]::new($TargetRootFolder, $Mailbox),
        $SearchFilterCollection,
        $FolderView
    )

    $EwsFolder
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

        [ItemSchema]::Size, ($LargeItemSizeMB * 1MB)
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

function Get-FolderItemSummary ($ExSvc, $Mailbox, $FolderId, [switch]$Archive) {

    $PR_MESSAGE_SIZE_EXTENDED = [ExtendedPropertyDefinition]::new(3592,[MapiPropertyType]::Long)
    $FolderPropertySet = [PropertySet]::new([BasePropertySet]::FirstClassProperties)
    $FolderPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED)

    $Folder = [Folder]::Bind($ExSvc,$FolderId,$FolderPropertySet)

    Write-Debug "STOP here"
    $FolderSummary = [PSCustomObject]@{

        Folder = $Folder.DisplayName
        AgeDays0To14 = [INT64]0
        AgeDays0To14Size = [INT64]0
        AgeDays15To30 = [INT64]0
        AgeDays15To30Size = [INT64]0
        AgeMonths1to11 = [INT64]0
        AgeMonths1to11Size = [INT64]0
        AgeYears1to2 = [INT64]0
        AgeYears1to2Size = [INT64]0
        AgeYears2to5 = [INT64]0
        AgeYears2to5Size = [INT64]0
        AgeYears5andOlder = [INT64]0
        AgeYears5andOlderSize = [INT64]0
    }

    $folderSizeVal = $null;
    if($Folder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED,[ref]$folderSizeVal)){
        $rptObject.TotalSize = [Math]::Round([Int64]$folderSizeVal / 1mb,2)
    }
    #Define ItemView to retrive just 1000 Items
    $ItemView =  [ItemView]::new(1000)
    $ItemPropset = [PropertySet]::new([BasePropertySet]::IdOnly)
    $ItemPropset.Add([ItemSchema]::DateTimeReceived)
    $ItemPropset.Add([ItemSchema]::DateTimeCreated)
    $ItemPropset.Add([ItemSchema]::LastModifiedTime)
    $ItemPropset.Add([ItemSchema]::Size)

    $Items = @()
    $PageSize = 1000
    $Offset = 0
    $MoreAvailable = $true

    do {
        $FindItemsResult = $ExSvc.FindItems($FolderId, $ItemView)
        $Items += $FindItemsResult.Items

        if ($FindItemsResult.MoreAvailable) {

            $Offset = $FindItemsResult.NextPageOffset
        }
        else { $MoreAvailable = $false }
    }
    while ($MoreAvailable)

    do{
        $fiItems = $service.FindItems($FolderId,$ItemView)
        #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)
        foreach($Item in $fiItems.Items){
            $Notadded = $true
            $dateVal = $null
            if($Item.TryGetProperty([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[ref]$dateVal )-eq $false){
                $dateVal = $Item.DateTimeCreated
            }
            $modDateVal = $null
            if($Item.TryGetProperty([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime,[ref]$modDateVal)){
                if($modDateVal -gt (Get-Date).AddDays(-7))
                {
                    $rptObject.DeletedLessThan7days++
                    $rptObject.DeletedLessThan7daysSize += $Item.Size
                    $Notadded = $false
                }
                if($modDateVal -le (Get-Date).AddDays(-7) -band $modDateVal -gt (Get-Date).AddDays(-30))
                {
                    $rptObject.Deleted7To30Days++
                    $rptObject.Deleted7To30DaysSize += $Item.Size
                    $Notadded = $false
                }
                if($modDateVal -le (Get-Date).AddDays(-30) -band $modDateVal -gt (Get-Date).AddMonths(-6))
                {
                    $rptObject.Deleted1to6Months++
                    $rptObject.Deleted1to6MonthsSize += $Item.Size
                    $Notadded = $false
                }
                if($modDateVal -le (Get-Date).AddMonths(-6) -band $modDateVal -gt (Get-Date).AddMonths(-12))
                {
                    $rptObject.Deleted6To12Months++
                    $rptObject.Deleted6To12MonthsSize += $Item.Size
                    $Notadded = $false
                }
                if($modDateVal -le (Get-Date).AddYears(-1))
                {
                    $rptObject.DeletedGreator12Months++
                    $rptObject.DeletedGreator12MonthsSize += $Item.Size
                    $Notadded = $false
                }
            }
            if($dateVal -gt (Get-Date).AddDays(-7))
            {
                $rptObject.AgeLessThan7days++
                $rptObject.AgeLessThan7daysSize += $Item.Size
                $Notadded = $false
            }
            if($dateVal -le (Get-Date).AddDays(-7) -band $dateVal -gt (Get-Date).AddDays(-30))
            {
                $rptObject.Age7To30Days++
                $rptObject.Age7To30DaysSize += $Item.Size
                $Notadded = $false
            }
            if($dateVal -le (Get-Date).AddDays(-30) -band $dateVal -gt (Get-Date).AddMonths(-6))
            {
                $rptObject.Age1to6Months++
                $rptObject.Age1to6MonthsSize += $Item.Size
                $Notadded = $false
            }
            if($dateVal -le (Get-Date).AddMonths(-6) -band $dateVal -gt (Get-Date).AddMonths(-12))
            {
                $rptObject.Age6To12Months++
                $rptObject.Age6To12MonthsSize += $Item.Size
                $Notadded = $false
            }
            if($dateVal -le (Get-Date).AddYears(-1))
            {
                $rptObject.AgeGreator12Months++
                $rptObject.AgeGreator12MonthsSize += $Item.Size
                $Notadded = $false
            }
        }
        $ivItemView.Offset += $fiItems.Items.Count
    }while($fiItems.MoreAvailable -eq $true)
    # if($rptObject.AgeLessThan7daysSize -gt 0){
    #     $rptObject.AgeLessThan7daysSize = [Math]::Round($rptObject.AgeLessThan7daysSize/ 1mb,2)
    # }
    # if($rptObject.Age7To30DaysSize -gt 0){
    #     $rptObject.Age7To30DaysSize = [Math]::Round($rptObject.Age7To30DaysSize/ 1mb,2)
    # }
    # if($rptObject.Age1to6MonthsSize -gt 0){
    #     $rptObject.Age1to6MonthsSize = [Math]::Round($rptObject.Age1to6MonthsSize/ 1mb,2)
    # }
    # if($rptObject.Age6To12MonthsSize -gt 0){
    #     $rptObject.Age6To12MonthsSize = [Math]::Round($rptObject.Age6To12MonthsSize/ 1mb,2)
    # }
    # if($rptObject.AgeGreator12MonthsSize -gt 0){
    #     $rptObject.AgeGreator12MonthsSize = [Math]::Round($rptObject.AgeGreator12MonthsSize/ 1mb,2)
    # }
    # if($rptObject.DeletedLessThan7daysSize -gt 0){
    #     $rptObject.DeletedLessThan7daysSize = [Math]::Round($rptObject.DeletedLessThan7daysSize/ 1mb,2)
    # }
    # if($rptObject.Deleted7To30DaysSize -gt 0){
    #     $rptObject.Deleted7To30DaysSize = [Math]::Round($rptObject.Deleted7To30DaysSize/ 1mb,2)
    # }
    # if($rptObject.Deleted1to6MonthsSize -gt 0){
    #     $rptObject.Deleted1to6MonthsSize = [Math]::Round($rptObject.Deleted1to6MonthsSize/ 1mb,2)
    # }
    # if($rptObject.Deleted6To12MonthsSize -gt 0){
    #     $rptObject.Deleted6To12MonthsSize = [Math]::Round($rptObject.Deleted6To12MonthsSize/ 1mb,2)
    # }
    # if($rptObject.DeletedGreator12MonthsSize -gt 0){
    #     $rptObject.DeletedGreator12MonthsSize = [Math]::Round($rptObject.DeletedGreator12MonthsSize/ 1mb,2)
    # }
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

    $EwsManagedApiFail = "EWS Managed API 2.2 is required, specifically product/file version 15.00.0913.015.`r`n" +
    "Download: https://www.microsoft.com/en-us/download/details.aspx?id=42951"

    if (-not (Test-Path -Path $EwsManagedApiDllPath)) { throw $EwsManagedApiFail }
    else {
        $EwsManagedApiDll = Get-ChildItem -Path $EwsManagedApiDllPath -ErrorAction Stop
    }

    if ($EwsManagedApiDll.VersionInfo.FileVersion -ne '15.00.0913.015') {

        throw $EwsManagedApiFail
    }
    Import-Module $EwsManagedApiDll -ErrorAction Stop

    writeLog @writeLogParams 'Successfully verified version and imported EWS Managed API 2.2 DLL (with Import-Module).'
    #endregion Initialization

    #region Mailbox Loop
    $MainProgressParams = @{

        Id       = 0
        Activity = "$($PSCmdlet.MyInvocation.MyCommand.Name) (Primary Mailboxes) - Start time: $($dtNow)"
    }

    $MailboxCounter = 0

    foreach ($Mailbox in $Mailboxes.SmtpAddress) {

        $MailboxCounter++

        $MainProgressParams['Status'] = "Mailbox $($MailboxCounter) of $($Mailboxes.Count)"
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

            # Attempt to find SmtpAddress connecting/authenticating user:
            if ($PSCmdlet.ParameterSetName -like '*NoImpersonation') {

                $Mailbox = Get-ConnectingUserSmtpAddress -ExSvc $ExSvc
            }

            $MainProgressParams['CurrentOperation'] = "Current mailbox: $($Mailbox)"

            Write-Debug "STOP here to develop the script"

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

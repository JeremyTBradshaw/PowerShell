<#
    .Synopsis
    Create 'Large Items (###MB+)' Search Folder in mailboxes (using EWS Managed API 2.2).

    .Description
    Create 'Large Items (###MB+)' Search Folder in mailboxes (using EWS Managed API 2.2) to help users find their
    large items.  This task commonly comes up when planning a migration to Exchange Online.  A good idea is to use the
    sibling script - Get-MailboxLargeItems.ps1 - to determine which mailboxes have large items in them.  Alternatively,
    just provide users with this new search folder and advise them to backup the items it finds so they don't lose them
    during their migration.

    .Notes
    - The account used for -Credential parameter needs to be assigned the ApplicationImpersonation RBAC role.
    - Logging is only done when -MailboxListCSV is used, whereas -PrimarySmtpAddress is meant for processing
    individual mailboxes one at a time (i.e. testing and/or one-off's).
    - The default value for -LargeItemSizeMB is 150, which is the maximum allowed message size when migrating to
    Exchange Online via Hybrid Exchange / Remote Moves (as of November 23 2020).

    .Parameter Credential
    Specifies a PSCredential object for the account to be used for connecting to EWS.

    .Parameter EwsManagedApiDllPath
    Specifies the path the the Exchange.WebServices.dll file.
    Download: https://www.microsoft.com/en-us/download/details.aspx?id=42951
    Defaults to 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'

    .Parameter LargeItemSizeMB
    Sets the size (in MB) of the large items to search for.  Default is 150MB.  If migrating via EWS (i.e. 3rd party
    tools for migration to EXO), this value should be set to 25.

    .Parameter MailboxListCSV
    Specifies the source CSV file containing mailboxes to create the new Search Folder in. There must be a "PrimarySmtpAddress" column header.

    .Parameter PrimarySmtpAddress
    Specifies the PrimarySmtpAddress of the mailbox to create the search folder in.

    .Parameter Archive
    Indicates to create the Search Folder in the archive mailbox (if one exists), rather than the primary mailbox.

    .Parameter EwsUrl
    Specifies the URL for the Exchange Web Services endpoint.

    .Example
    # Create large items search folder, in mailboxes supplied in the CSV file:
    New-LargeItemsSearchFolder.ps1 -EwsUrl https://ex2016.contoso.com/ews/exchange.asmx -Credential <PSCredential> -MailboxListCSV .\LIUsers.csv

    .Example
    # Create large items search folder, in mailboxes supplied in the CSV file (using a large item definition other than 150MB):
    New-LargeItemsSearchFolder.ps1 -EwsUrl https://ex2016.contoso.com/ews/exchange.asmx -Credential <PSCredential> -MailboxListCSV .\LIUsers.csv -LargeItemSizeMB <Value in MB>

    .Example
    # Create large items search folder, both in the primary and archive mailbox, in mailboxes supplied in the CSV file:
    New-LargeItemsSearchFolder.ps1 -EwsUrl https://ex2016.contoso.com/ews/exchange.asmx -Credential <PSCredential> -LargeItemSizeMB <Value in MB> -Archive -MailboxListCSV .\LIUsers.csv

    .Link
    Install the EWS Managed API 2.2:  http://www.microsoft.com/en-us/download/details.aspx?id=42951

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/edit/master/New-LargeItemsSearchFolder.ps1

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/blob/master/Get-MailboxLargeItems.ps1

    .Outputs
    # Sample log file (when using -MailboxListCSV):
    New-LargeItemsSearchFolder_2020-11-25_11-03-41.log:
        [ 2020-11-25 11:03:41 AM ] New-LargeItemsSearchFolder.ps1 - Script begin.
        [ 2020-11-25 11:03:41 AM ] LargeItemsSizeMB set to 150 MB.
        [ 2020-11-25 11:03:41 AM ] Targeting Archive mailboxes (-Archive switch parameter was used).
        [ 2020-11-25 11:03:41 AM ] Successfully imported $MailboxListCSV, with 33 mailboxes to process.
        [ 2020-11-25 11:03:41 AM ] Successfully verified version and imported EWS Managed API 2.2 DLL (with Import-Module).
        [ 2020-11-25 11:03:41 AM ] Mailbox: 1 of 33
        [ 2020-11-25 11:03:41 AM ] Mailbox: Larry.Iceberg@jb365.ca | Connecting to EWS as contoso\ExAdmin123.
        [ 2020-11-25 11:03:42 AM ] Mailbox: Larry.Iceberg@jb365.ca | Created Search Folder 'Large Items (150MB+)'.
        [ 2020-11-25 11:03:42 AM ] Mailbox: 2 of 33
        [ 2020-11-25 11:03:43 AM ] Mailbox: Louis.Isaacson@jb365.ca | Connecting to EWS as contoso\ExAdmin123.
        [ 2020-11-25 11:03:43 AM ] Mailbox: Louis.Isaacson@jb365.ca | Created Search Folder 'Large Items (150MB+)'.
        ...
        ...
        [ 2020-11-25 11:04:47 AM ] New-LargeItemsSearchFolder.ps1 - Script end.
#>
#Requires -Version 5.1

using namespace Microsoft.Exchange.WebServices.Data

[CmdletBinding(
    DefaultParameterSetName = 'MailboxListCSV'
)]
param(
    [Parameter(Mandatory)]
    [System.Management.Automation.PSCredential]$Credential,

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

    [Parameter(
        Mandatory,
        ParameterSetName = 'MailboxListCSV'
    )]
    [ValidateScript(
        {
            if (Test-Path -Path $_) { $true } else {

                throw "Could not find CSV file $($_)."
            }
        }
    )]
    [System.IO.FileInfo]$MailboxListCSV,

    [Parameter(
        Mandatory,
        ParameterSetName = 'PrimarySmtpAddress'
    )]
    [ValidatePattern('(^.*\@.*\..*$)')]
    [string]$PrimarySmtpAddress,

    [switch]$Archive,
	
    [Parameter(Mandatory)]
    [uri]$EwsUrl
)

function writeLog {
    param(
        [Parameter(Mandatory)]
        [string]$LogName,
    
        [Parameter(Mandatory)]
        [string]$Message,
    
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$Folder,
    
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [datetime]$LogDateTime,
        [switch]$DisableLogging
    )

    if (-not $DisableLogging) {

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
        $MessageText | Out-File -FilePath $LogFile -Append
        
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

function New-SearchFolder {
    [CmdletBinding()]
    param(
        [ExchangeService]$ExchangeService,
        [int16]$LargeItemSizeMB,
        [switch]$Archive
    )

    try {
        $SearchFilter = [SearchFilter+IsGreaterThanOrEqualTo]::new([ItemSchema]::Size, ($LargeItemSizeMB * 1MB))
        $SearchFolder = [SearchFolder]::new($ExchangeService)
        $SearchFolder.DisplayName = "Large Items ($($LargeItemSizeMB)MB+)"
        $SearchFolder.SearchParameters.Traversal = 'Deep'
        $SearchFolder.SearchParameters.SearchFilter = $SearchFilter

        if ($Archive) {
    
            $ArchiveSearchFolders = $ExchangeService.FindFolders(

                [WellKnownFolderName]::ArchiveRoot,
                [SearchFilter+IsEqualTo]::new([FolderSchema]::DisplayName, 'Finder'), [FolderView]::new(1)
            )
            if ($ArchiveSearchFolders) {
    
                $SearchFolder.SearchParameters.RootFolderIds.Add([WellKnownFolderName]::ArchiveMsgFolderRoot)
                $SearchFolder.Save($ArchiveSearchFolders.Folders.Id)
            }
            else {
                throw "Finder folder (a.k.a. 'Search Folders') wasn't found in the Archive mailbox."
            }
        }
        else {
            $SearchFolder.SearchParameters.RootFolderIds.Add([WellKnownFolderName]::MsgFolderRoot)
            $SearchFolder.Save([WellKnownFolderName]::SearchFolders)
        }
    }
    catch { throw $_ }
}

function New-EwsBinding {

    param(
        [string]$Mailbox,
        [System.Management.Automation.PSCredential]$Credential,
        [uri]$Url
    )

    # Going with Exchange2010_SP1 because it is the earliest version of the EWS schema that does what we need, per:
    # https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/ews-schema-versions-in-exchange#designing-your-application-with-schema-version-in-mind
    $ExchangeService = [ExchangeService]::new(
        
        [ExchangeVersion]::Exchange2010_SP1
    )
	
    $ExchangeService.Url = $Url.AbsoluteUri
    $ExchangeService.Credentials = [System.Net.NetworkCredential]($Credential)
    $ExchangeService.ImpersonatedUserId = [ImpersonatedUserId]::new(
    
        [ConnectingIdType]::SmtpAddress, $Mailbox
    )
	
    $ExchangeService
}

#region Main Script

try {
    $dtNow = [datetime]::Now

    $writeLogParams = @{

        LogName     = "$($MyInvocation.MyCommand.Name -replace '\.ps1')"
        Folder      = "$($PSScriptRoot)\$($MyInvocation.MyCommand.Name -replace '\.ps1')_Logs"
        LogDateTime = $dtNow
        ErrorAction = 'Stop'
    }

    $Mailboxes = @()

    if ($PSCmdlet.ParameterSetName -eq 'MailboxListCSV') {

        # Check for and if necessary create logs folder:
        if (-not (Test-Path -Path "$($writeLogParams['Folder'])")) {
            
            [void](New-Item -Path "$($writeLogParams['Folder'])" -ItemType Directory -ErrorAction Stop)
        }

        writeLog @writeLogParams -Message 'New-LargeItemsSearchFolder.ps1 - Script begin.'
        writeLog @writeLogParams -Message "LargeItemsSizeMB set to $($LargeItemSizeMB) MB."
        
        if ($PSBoundParameters.ContainsKey('Archive')) {

            writeLog @writeLogParams -Message 'Targeting Archive mailboxes (-Archive switch parameter was used).'
        }
        else {
            writeLog @writeLogParams -Message 'Targeting Primary mailboxes (-Archive switch parameter was not used).'
        }

        $Mailboxes += Import-Csv $MailboxListCSV -ErrorAction Stop

        if ($Mailboxes.Count -ge 1) {

            if (-not ($Mailboxes | Get-Member -Name 'PrimarySmtpAddress')) {

                throw "CSV file '$($MailboxListCSV)' is missing the mandatory PrimarySmtpAddress column header.  Exiting script."
            }

            writeLog @writeLogParams -Message "Successfully imported `$MailboxListCSV, with $($Mailboxes.Count) mailboxes to process."
        }
        else {
            throw "Script-ending failure:  CSV file '$($MailboxListCSV)' had no rows to process."
        }
    }
    else {
        # Disable logging for single PrimarySmtpAddress processing.
        $writeLogParams['DisableLogging'] = $true

        $Mailboxes += [PSCustomObject]@{ PrimarySmtpAddress = $PrimarySmtpAddress }
    }

    $EwsManagedApiDll = Get-ChildItem -Path $EwsManagedApiDllPath -ErrorAction Stop

    if ($EwsManagedApiDll.VersionInfo.FileVersion -ne '15.00.0913.015') {

        $errorMessage = "EWS Managed API 2.2 is required, specifically product/file version 15.00.0913.015.`r`n" +
        "Download: https://www.microsoft.com/en-us/download/details.aspx?id=42951"
        
        throw $errorMessage
    }
    Import-Module $EwsManagedApiDll -ErrorAction Stop

    writeLog @writeLogParams 'Successfully verified version and imported EWS Managed API 2.2 DLL (with Import-Module).'

    $MainProgressParams = @{

        Id               = 0
        Activity         = "New-LargeItemsSearchFolder.ps1 - Start time: $($dtNow)"
    }
    
    $MailboxCounter = 0

    foreach ($Mailbox in $Mailboxes.PrimarySmtpAddress) {

        $MailboxCounter++

        $MainProgressParams['Status'] = "Finding large items ($($LargeItemSizeMB)+ MB) | Mailbox $($MailboxCounter) of $($Mailboxes.Count))"
        $MainProgressParams['CurrentOperation'] = "Current mailbox: $($Mailbox)"
        $MainProgressParams['PercentComplete']  = (($MailboxCounter / $Mailboxes.Count) * 100)

        if ($Archive) {

            $MainProgressParams['CurrentOperation'] = $MainProgressParams['CurrentOperation'] -replace 'mailbox:', 'mailbox (Archive):'
        }
        Write-Progress @MainProgressParams
        
        try {
            writeLog @writeLogParams -Message "Mailbox: $($MailboxCounter) of $($Mailboxes.Count)"
    
            writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Connecting to EWS as $($Credential.UserName)."
            $ExchangeService = New-EwsBinding -Mailbox $Mailbox -Credential $Credential -Url $EwsUrl
            
            [void](New-SearchFolder -ExchangeService $ExchangeService -LargeItemSizeMB $LargeItemSizeMB -Archive:$Archive -ErrorAction Stop)
            writeLog @writeLogParams -Message "Mailbox: $($Mailbox) | Created Search Folder 'Large Items ($($LargeItemSizeMB)MB+)'."
        }
        catch [System.Management.Automation.ActionPreferenceStopException] { <# Suppress warning/error when halting after debugging. #> }
        catch {
            if (
                # Script will continue processing mailboxes (if using -MailboxListCSV) for the following errors:
                ($_.ToString().contains('The SMTP address has no mailbox associated with it.')) -or
                ($PSBoundParameters.ContainsKey('Archive') -and $_.ToString().contains('The specified folder could not be found in the store.')) -or
                ($_.ToString().Contains('"A folder with the specified name already exists.'))
             ) {
                $currentMsg = $null
                $currentMsg = "Mailbox: $($Mailbox) | Non-script-ending error: $($_.Exception.Message)."

                Write-Warning -Message $currentMsg
                writeLog @writeLogParams -Message $currentMsg -ErrorRecord $_
                Write-Debug -Message 'Suspend the script here to investigate the error. Otherwise, unless Halted, the script will continue.'

                continue
            }
            else {
                $currentMsg = $null
                $currentMsg = "Mailbox $($Mailbox) | Script-ending error: $($_.Exception.Message)"

                Write-Warning -Message $currentMsg
                writeLog @writeLogParams -Message $currentMsg -ErrorRecord $_
                throw $_
            }
        }	
    }
}
catch [System.Management.Automation.ActionPreferenceStopException] { break }
catch { throw $_ }

finally { writeLog @writeLogParams -Message 'New-LargeItemsSearchFolder.ps1 - Script end.' }

#endRegion Main Script

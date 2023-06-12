<#
    .Synopsis
    EXO-exclusive alternative to Get-MailboxTrustee.ps1, using the newer V3 cmdlets, which are REST-backed.

    .Parameter MailboxId
    Specifies a unique identifier property for one or more mailboxes to process, accepting pipeline input or manual
    parameter entry.  Example properties to use: Guid, ExternalDirectoryObjectId, PrimarySmtpAddress, ExchangeGuid.

    .Parameter All
    Targets all mailboxes in the connected Exchange Online organization.

    .Parameter IncludePermissionTypes
    If not specified, all permission types are collected - FullAccess, SendAs, Send on Behalf, and common folders:
        - \ (mailbox root (a.k.a. Top of Information Store)), Inbox, Sent Items
        - Calendar, Contacts, Tasks
    Specify one or more options to control which permissions are collected.

    .Example
    .\Get-EXOMailboxTrustee.ps1 -All

    .Example
    Get-Mailbox Conference* | .\Get-EXOMailboxTrustee.ps1

    .Example
    .\Get-EXOMailboxTrustee.ps1 -All -IncludePermissionTypes FullAccess, SendAs

    .Example
    .\Get-EXOMailboxTrustee.ps1 -ExternalDirectoryObjectId 12345678-1234-1234-123456789012 -IncludePermissionTypes CommonFolders.

    .Outputs
    The outputted objects contain five properties for both the mailbox and for the trustee, for guaranteed unique
    identification:
        - DisplayName, PrimarySmtpAddress, RecipientTypeDetails
        - Guid, ExternalDirectoryObjectId
    Additionally, the folder (if applicable) and permission (i.e. AccessRights) are included:
        - Folder, Permission

    For example:
    [PS]\> Get-Mailbox test* | .Get-EXOMailboxTrustee.ps1

    MailboxDisplayName               : Test User 1
    MailboxPrimarySmtpAddress        : TestUser1@jb365ca.onmicrosoft.com
    MailboxRecipientTypeDetails      : UserMailbox
    MailboxGuid                      : d1ce4896-3278-4285-8aa3-a1f5e9098cd3
    MailboxExternalDirectoryObjectId : 66b05549-c470-42e2-af64-bb843cfc8130
    Folder                           : #N/A
    Permission                       : Send-As
    TrusteeDisplayName               : Test User 2
    TrusteePrimarySmtpAddress        : TestUser2@jb365ca.onmicrosoft.com
    TrusteeRecipientTypeDetails      : UserMailbox
    TrusteeGuid                      : 7ee0436c-a982-4447-a58e-e219f10883fd
    TrusteeExternalDirectoryObjectId : 72943358-59b4-45f7-bd3c-4407bcd2363f

    .Notes
     -  As of 2022-12-19, anywhere that Write-Progress is used, it is preceded with $ProgressPreference = 'Continue'.
    This is due to an issue with the ExchangeOnlineManagement module where it sets the global variable to
    'SilentlyContinue', hiding all progress.

     - As of 2023-06-07, FullAccess, Send-As, and Send-on-Behalf permissions are stored on mailbox objects in an
    undocumented way.  FullAccess/Send-As trustees are stored as their UserPrincipalName.  Send-on-Behalf trustees are
    stored as their Name/cn property value, I think**.  The following TechCommunity thread has been opened to discuss:
    https://techcommunity.microsoft.com/t5/exchange/exo-s-quot-user-quot-and-quot-trustee-quot-properties-returned/m-p/3834679#M11590
    Until this is sorted out concretely, I'm using Get-Recipient to filter by UPN or Name, when finding the Trustees.
#>
#Requires -Version 5.1
#Requires -Modules @{ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.0.0'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'}

[CmdletBinding(DefaultParameterSetName = 'CallerSpecified')]
param(
    [Parameter(
        ParameterSetName = 'CallerSpecified',
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        HelpMessage = 'Use a unique identifier of the mailbox (e.g., Guid, ExternalDirectoryObjectId, PrimarySmtpAddress, etc.).'
    )]
    [Alias('ExternalDirectoryObjectId', 'Guid', 'ExchangeGuid')]
    [string[]]$MailboxId,

    [Parameter(ParameterSetName = 'All')]
    [switch]$All,

    [ValidateSet(
        'All',
        'FullAccess', 'SendAs', 'SendOnBehalf',
        'CommonFolders', '\', 'Inbox', 'SentItems', 'Calendar', 'Contacts', 'Tasks'
    )]
    [string[]]$IncludePermissionTypes = 'All'
)

begin {

    $StopWatch1 = [System.Diagnostics.Stopwatch]::StartNew()
    $Progress = @{
        Activity         = "Get-EXOMailboxTrustee.ps1 - Start time: $([datetime]::Now-$StopWatch1.Elapsed)"
        PercentComplete  = -1
        Status           = '...'
        CurrentOperation = "Initializing"
    }
    $ProgressPreference = 'Continue'; Write-Progress @Progress

    $_requiredCommands = @(
        'Get-Mailbox', 'Get-Recipient',
        'Get-MailboxPermission', 'Get-RecipientPermission', 'Get-MailboxFolderPermission'
    )
    if (-not (Get-Command $_requiredCommands -ErrorAction SilentlyContinue)) {

        throw "This script requires an active connection with v3.0.0 or newer of the ExchangeOnlineManagement module (i.e., Connect-ExchangeOnline).  " +
        "Once connected, the following commands are required to be available: $($_requiredCommands -join ', ')"
    }

    if ($PSCmdlet.ParameterSetName -eq 'All') {

        $Progress['CurrentOperation'] = 'Getting all EXO mailboxes'
        $ProgressPreference = 'Continue'; Write-Progress @Progress

        $MailboxId = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)
    }

    if (-not ($PSBoundParameters.ContainsKey('IncludePermissionTypes'))) {

        $IncludePermissionTypes = 'All'
    }

    $TrusteeTracker = @{ }
    $ProcessedCounter = 0
    $StopWatch2 = [System.Diagnostics.Stopwatch]::StartNew()

    $Progress['CurrentOperation'] = 'Starting to process mailboxes'
    $ProgressPreference = 'Continue'; Write-Progress @Progress
}

process {

    $MailboxId | ForEach-Object {

        if ($PSCmdlet.ParameterSetName -eq 'CallerSpecified') {

            $currentMBX = Get-Mailbox -Identity $_.ToString() -ErrorAction Stop
        }
        else { $currentMBX = $_ }

        $ProcessedCounter++
        if ($StopWatch2.Elapsed.Milliseconds -ge 500) {

            if ($PSCmdlet.ParameterSetName -eq 'All') {

                $Progress['CurrentOperation'] = "Mailboxes processed: $ProcessedCounter of $($MailboxId.Count); Time elapsed: $($StopWatch1.Elapsed -replace '\..*')"
                $Progress['PercentComplete'] = ($ProcessedCounter / $MailboxId.Count) * 100
            }
            else {
                $Progress['CurrentOperation'] = "Mailboxes processed: $ProcessedCounter; Time elapsed: $($StopWatch1.Elapsed -replace '\..*')"
                $Progress['PercentComplete'] = -1
            }
            $Progress['Status'] = 'Processing...'
            $ProgressPreference = 'Continue'; Write-Progress @Progress

            $StopWatch2.Reset(); $StopWatch2.Start()
        }

        $currentMBXPermissions = @()
        $currentMBXPermissionLookupFailures = @()

        #region FullAccess

        if ($IncludePermissionTypes -match '(^All$)|(^FullAccess$)') {

            try {
                $currentMBXPermissions += Get-MailboxPermission -Identity $currentMBX.ExternalDirectoryObjectId -ErrorAction Stop |
                Where-Object {
                    $_.IsInherited -ne $true -and
                    $_.Deny -ne $true -and
                    $_.AccessRights -like '*FullAccess*' -and
                    $_.User -ne 'NT AUTHORITY\SELF' -and
                    $_.User -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { '#N/A' } },
                @{Name = 'Permission'; Expression = { 'FullAccess' } },
                User
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = '#N/A'
                    Permission = 'FullAccess'
                }
            }
        }
        #endregion FullAccess

        #region SendAs

        if ($IncludePermissionTypes -match '(^All$)|(^SendAs$)') {

            try {
                $currentMBXPermissions += Get-RecipientPermission -Identity $currentMBX.ExternalDirectoryObjectId -ErrorAction Stop |
                Where-Object {
                    $_.IsInherited -ne $true -and
                    $_.Deny -ne $true -and
                    $_.AccessRights -like '*SendAs*' -and
                    $_.Trustee -ne 'NT AUTHORITY\SELF' -and
                    $_.Trustee -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { '#N/A' } },
                @{Name = 'Permission'; Expression = { 'Send-As' } },
                @{Name = 'User'; Expression = { $_.Trustee } }
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = '#N/A'
                    Permission = 'SendAs'
                }
            }
        }
        #endregion SendAs

        #region SendOnBehalf

        if ($IncludePermissionTypes -match '(^All$)|(^SendOnBehalf$)') {

            if ($currentMBX.GrantSendOnBehalfTo) {

                $currentMBXPermissions += $currentMBX.GrantSendOnBehalfTo |
                Select-Object @{Name = 'Folder'; Expression = { '#N/A' } },
                @{Name = 'Permission'; Expression = { 'Send on Behalf' } },
                @{Name = 'User'; Expression = { $_ } }
            }
        }
        #endregion SendOnBehalf

        #region Calendar

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^Calendar$)') {

            try {
                $currentMBXPermissions += Get-MailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Calendar" -ErrorAction Stop |
                Where-Object {
                    $_.User -notlike 'Default' -and
                    $_.User -notlike 'Anonymous' -and
                    $_.User -notlike 'NT:*' -and
                    $_.User -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { 'Calendar' } },
                @{Name = 'Permission'; Expression = { $_.AccessRights -join ',' } },
                User
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = 'Calendar'
                    Permission = '#N/A'
                }
            }
        }
        #endregion Calendar

        #region Contacts

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^Contacts$)') {

            try {
                $currentMBXPermissions += Get-MailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Contacts" -ErrorAction Stop |
                Where-Object {
                    $_.User -notlike 'Default' -and
                    $_.User -notlike 'Anonymous' -and
                    $_.User -notlike 'NT:*' -and
                    $_.User -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { 'Contacts' } },
                @{Name = 'Permission'; Expression = { $_.AccessRights -join ',' } },
                User
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = 'Contacts'
                    Permission = '#N/A'
                }
            }
        }
        #endregion Contacts

        #region Tasks

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^Tasks$)') {

            try {
                $currentMBXPermissions += Get-MailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Tasks" -ErrorAction Stop |
                Where-Object {
                    $_.User -notlike 'Default' -and
                    $_.User -notlike 'Anonymous' -and
                    $_.User -notlike 'NT:*' -and
                    $_.User -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { 'Tasks' } },
                @{Name = 'Permission'; Expression = { $_.AccessRights -join ',' } },
                User
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = 'Tasks'
                    Permission = '#N/A'
                }
            }
        }
        #endregion Tasks

        #region MailboxRoot

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^\\$)') {

            try {
                $currentMBXPermissions += Get-MailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\" -ErrorAction Stop |
                Where-Object {
                    $_.User -notlike 'Default' -and
                    $_.User -notlike 'Anonymous' -and
                    $_.User -notlike 'NT:*' -and
                    $_.User -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { '\' } },
                @{Name = 'Permission'; Expression = { $_.AccessRights -join ',' } },
                User
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = '\'
                    Permission = '#N/A'
                }
            }
        }
        #endregion MailboxRoot

        #region Inbox

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^Inbox$)') {

            try {
                $currentMBXPermissions += Get-MailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Inbox" -ErrorAction Stop |
                Where-Object {
                    $_.User -notlike 'Default' -and
                    $_.User -notlike 'Anonymous' -and
                    $_.User -notlike 'NT:*' -and
                    $_.User -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { 'Inbox' } },
                @{Name = 'Permission'; Expression = { $_.AccessRights -join ',' } },
                User
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = 'Inbox'
                    Permission = '#N/A'
                }
            }
        }
        #endregion Inbox

        #region SentItems

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^SentItems$)') {

            try {
                $currentMBXPermissions += Get-MailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Sent Items" -ErrorAction Stop |
                Where-Object {
                    $_.User -notlike 'Default' -and
                    $_.User -notlike 'Anonymous' -and
                    $_.User -notlike 'NT:*' -and
                    $_.User -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { 'Sent Items' } },
                @{Name = 'Permission'; Expression = { $_.AccessRights -join ',' } },
                User
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = 'SentItems'
                    Permission = '#N/A'
                }
            }
        }
        #endregion SentItems

        #region Post-processing

        foreach ($cmp in $currentMBXPermissions) {

            Write-Verbose -Message "Looking up trustee: '$($cmp.User)'"
            Write-Debug "Troubleshoot here"
            $Recipient = $null
            if ($TrusteeTracker["$($cmp.User)"]) {

                Write-Verbose -Message "Trustee found in cache, no need to re-lookup."
                $Recipient = $TrusteeTracker["$($cmp.User)"]
            }
            else {
                Write-Verbose -Message "Trustee not found in cache, performing lookup."
                try {
                    $PSmtpOrName = $null

                    if ($cmp.User -like '*@*') {

                        $PSmtpOrName = 'UPN/PSMTP'
                        $Recipient = Get-Recipient -Filter "PrimarySmtpAddress -eq '$($cmp.User)' -or UserPrincipalName -eq '$($cmp.User)'" -ErrorAction Stop
                    }
                    else {

                        $PSmtpOrName = 'Name'
                        $Recipient = Get-Recipient -Identity $($cmp.User) -ErrorAction Stop
                    }
                }
                catch {
                    Write-Warning -Message "Failed on Get-Recipient command."
                    Write-Warning -Message "Detected User ID property = $($PSmtpOrName)"
                    Write-Warning -Message "Mailbox (PSMTP) = '$($currentMBX.PrimarySmtpAddress)'"
                    Write-Warning -Message "Folder = '$($cmp.Folder)'"
                    Write-Warning -Message "Permission = '$($cmp.Permission)'"
                    Write-Warning -Message "`$cmp.User = '$($cmp.User)'"
                }

                if ($Recipient -and (-not($Recipient.Count -gt 1))) {

                    $TrusteeTracker["$($cmp.User)"] = $Recipient |
                    Select-Object -Property DisplayName, PrimarySmtpAddress, RecipientTypeDetails, Guid, ExternalDirectoryObjectId
                }
                else {
                    $TrusteeTracker["$($cmp.User)"] = [PSCustomObject]@{
                        DisplayName               = "Unknown ('User' value = $($cmp.User))"
                        PrimarySmtpAddress        = 'Not found or ambiguous'
                        RecipientTypeDetails      = 'Not found or ambiguous'
                        Guid                      = 'Not found or ambiguous'
                        ExternalDirectoryObjectId = 'Not found or ambiguous'
                    }
                }
            }

            $cmpOutput = [PSCustomObject]@{
                MailboxDisplayName               = $currentMBX.DisplayName
                MailboxPrimarySmtpAddress        = $currentMBX.PrimarySmtpAddress
                MailboxRecipientTypeDetails      = $currentMBX.RecipientTypeDetails
                MailboxGuid                      = $currentMBX.Guid
                MailboxExternalDirectoryObjectId = $currentMBX.ExternalDirectoryObjectId
                Folder                           = $cmp.Folder
                Permission                       = $cmp.Permission
                TrusteeDisplayName               = $TrusteeTracker["$($cmp.User)"].DisplayName
                TrusteePrimarySmtpAddress        = $TrusteeTracker["$($cmp.User)"].PrimarySmtpAddress
                TrusteeRecipientTypeDetails      = $TrusteeTracker["$($cmp.User)"].RecipientTypeDetails
                TrusteeGuid                      = $TrusteeTracker["$($cmp.User)"].Guid
                TrusteeExternalDirectoryObjectId = $TrusteeTracker["$($cmp.User)"].ExternalDirectoryObjectId
            }
            Write-Output -InputObject $cmpOutput
        }

        foreach ($cmplf in $currentMBXPermissionLookupFailures) {

            $cmplfOutput = [PSCustomObject]@{
                MailboxDisplayName               = $currentMBX.DisplayName
                MailboxPrimarySmtpAddress        = $currentMBX.PrimarySmtpAddress
                MailboxRecipientTypeDetails      = $currentMBX.RecipientTypeDetails
                MailboxGuid                      = $currentMBX.Guid
                MailboxExternalDirectoryObjectId = $currentMBX.ExternalDirectoryObjectId
                Folder                           = $cmplf.Folder
                Permission                       = $cmplf.Permission
                TrusteeDisplayName               = 'Unknown (cmdlet failed for unknown reason)'
                TrusteePrimarySmtpAddress        = ''
                TrusteeRecipientTypeDetails      = ''
                TrusteeGuid                      = ''
                TrusteeExternalDirectoryObjectId = ''
            }
            Write-Output -InputObject $cmplfOutput
        }
        #endregion Post-processing
    }
}

End {
    $ProgressPreference = 'Continue'; Write-Progress @Progress -Completed
}

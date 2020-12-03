<#
    .Synopsis
    EXO-exclusive successor to Get-MailboxTrustee.ps1, using only the newer V2 cmdlets.

    .Description
    No more Ivoke-Command, nor using the legacy cmdlets.  Now we're only the REST API cmdlets:
        - Get-EXOMailbox
        - Get-EXORecipient
        - Get-EXOMailboxPermission
        - Get-EXORecipientPermission
        - Get-EXOMailboxFolderPermission
    ...using the Azure AD ObjectId (a.k.a. ExternalDirectoryObjectId in EXO) for the Identity parameter throughout.

    .Parameter All
    Targets all mailboxes in the connected Exchange Online organization.

    .Parameter ExternalDirectoryaObjectId
    Targets only specified mailboxes, accepting pipeline input or manual parameter entry.

    .Parameter IncludePermissionTypes
    If not specified, all permission types are collected - FullAccess, SendAs, Send on Behalf, and common folders:
        - \ (mailbox root (a.k.a. Top of Information Store)), Inbox, Sent Items
        - Calendar, Contacts, Tasks
    Specify one or more options to control which permissions are collected.

    .Example
    .\Get-EXOMailboxTrustee.ps1 -All
    Get-Mailbox Conference* | .\Get-EXOMailboxTrustee.ps1
    .\Get-EXOMailboxTrustee.ps1 -All -IncludePermissionTypes FullAccess, SendAs
    .\Get-EXOMailboxTrustee.ps1 -ExternalDirectoryObjectId 12345678-1234-1234-123456789012 -IncludePermissionTypes CommonFolders

    .Notes
    Currently the -Filter parameter on the V2 cmdlets doesn't support one of PowerShell's quoting rules, so filtering
    for values containing apostrophes is not possible (e.g. "Name -eq 'O''Doyle'" <--:see the doubling up of ').
    The problem with this is that Microsoft return the Name property for users on mailbox folder permission objects.
    So if a user with an apostrophe in their name holds a permission, that user won't be able to be looked up with
    Get-EXORecipient in order to obtain their guaranteed-unique fields like Guid, ExternalDirectoryObjectId.  If/when
    Microsoft resolve this issue, I'll update the script as well.  Until then, I capture these missed users with a note
    in the TrusteeDisplayName output property.

    .Outputs
    The outputted objects contain five properties for both the mailbox and for the trustee:
        - DisplayName, PrimarySmtpAddress, RecipientTypeDetails
        - Guid, ExternalDirectory
    This should allow for plenty slice-and-dice capabilities in Excel, PowerBI, or a database.
    Additionally, the folder (if applicable) and permission (i.e. AccessRights) are included:
        - Folder, Permission

    For example:
    [PS]\> Get-Mailbox test* | .Get-EXOMailboxTrustee.ps1
    
    MailboxDisplayName               : Test User 1
    MailboxPrimarySmtpAddress        : TestUser1@jb365ca.onmicrosoft.com
    MailboxRecipientTypeDetails      : UserMailbox
    MailboxGuid                      : d1ce4896-3278-4285-8aa3-a1f5e9098cd3
    MailboxExternalDirectoryObjectId : 66b05549-c470-42e2-af64-bb843cfc8130
    Folder                           : N/A
    Permission                       : Send-As
    TrusteeDisplayName               : Test User 2
    TrusteePrimarySmtpAddress        : TestUser2@jb365ca.onmicrosoft.com
    TrusteeRecipientTypeDetails      : UserMailbox
    TrusteeGuid                      : 7ee0436c-a982-4447-a58e-e219f10883fd
    TrusteeExternalDirectoryObjectId : 72943358-59b4-45f7-bd3c-4407bcd2363f
#>
#Requires -Version 5.1
#Requires -Modules ExchangeOnlineManagement
[CmdletBinding(DefaultParameterSetName = 'CallerSpecified')]
param(
    [Parameter(
        ParameterSetName = 'CallerSpecified',
        Mandatory,
        ValueFromPipelineByPropertyName,
        HelpMessage = "Use either the ObjectId of the mailbox's associated Azure AD user object, or the ExternalDirectoryObjectId of the mailbox itself (these ID's are linked)."
    )]
    [Alias('ObjectId')]
    [guid[]]$ExternalDirectoryObjectId,

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
        Activity         = "Get-EXOMailboxTrustee - Start time: $([datetime]::Now-$StopWatch1.Elapsed)"
        PercentComplete  = -1
        Status           = '...'
        CurrentOperation = "Initializing"
    }
    Write-Progress @Progress

    if (-not (Get-PSSession | Where-Object { $_.ComputerName -eq 'outlook.office365.com' })) {
        Write-Warning -Message 'Must be connected to Exchange Online PowerShell V2 (i.e. Connect-ExchangeOnline).'
        break
    }
    Write-Verbose -Message "EXO PS session detected.  Proceeding with script."

    if ($PSCmdlet.ParameterSetName -eq 'All') {

        $Progress['CurrentOperation'] = 'Getting all EXO mailboxes'
        Write-Progress @Progress

        $EXOMailboxes = Get-EXOMailbox -Properties DisplayName, RecipientTypeDetails, PrimarySmtpAddress, Guid, ExternalDirectoryObjectId -ResultSize Unlimited
    }
    
    if (-not ($PSBoundParameters.ContainsKey('IncludePermissionTypes'))) {

        $IncludePermissionTypes = 'All'
    }

    $TrusteeTracker = @{ }
}

process {

    if ($PSCmdlet.ParameterSetName -eq 'CallerSpecified') {

        $Progress['CurrentOperation'] = 'Getting specified EXO mailboxes'
        Write-Progress @Progress

        $EXOMailboxes = foreach ($edoid in $ExternalDirectoryObjectId) {

            Get-EXOMailbox -Identity $edoid -Properties DisplayName, RecipientTypeDetails, PrimarySmtpAddress, Guid, ExternalDirectoryObjectId, GrantSendOnBehalfTo
        }
    }

    $Progress['CurrentOperation'] = 'Starting to process mailboxes'
    Write-Progress @Progress

    $ProcessedCounter = 0
    $StopWatch2 = [System.Diagnostics.Stopwatch]::StartNew()

    $EXOMailboxes | ForEach-Object {

        $ProcessedCounter++
        if ($StopWatch2.Elapsed.Milliseconds -ge 500) {

            $Progress['CurrentOperation'] = "Mailbox $ProcessedCounter of $($EXOMailboxes.Count); Time elapsed: $($StopWatch1.Elapsed -replace '\..*')"
            $Progress['PercentComplete'] = (($ProcessedCounter / $EXOMailboxes.Count) * 100)
            $Progress['Status'] = 'Processing...'
            Write-Progress @Progress

            $StopWatch2.Reset(); $StopWatch2.Start()
        }

        $currentMBX = $_
        $currentMBXPermissions = @()
        $currentMBXPermissionLookupFailures = @()

        #region FullAccess

        if ($IncludePermissionTypes -match '(^All$)|(^FullAccess$)') {

            try {
                $currentMBXPermissions += Get-EXOMailboxPermission -Identity $currentMBX.ExternalDirectoryObjectId -ErrorAction Stop |
                Where-Object {
                    $_.IsInherited -ne $true -and
                    $_.Deny -ne $true -and
                    $_.AccessRights -like '*FullAccess*' -and
                    $_.User -ne 'NT AUTHORITY\SELF' -and
                    $_.User -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { 'N/A' } },
                @{Name = 'Permission'; Expression = { 'FullAccess' } },
                User
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = 'N/A'
                    Permission = 'FullAccess'
                }
            }
        }
        #endregion FullAccess

        #region SendAs

        if ($IncludePermissionTypes -match '(^All$)|(^SendAs$)') {

            try {
                $currentMBXPermissions += Get-EXORecipientPermission -Identity $currentMBX.ExternalDirectoryObjectId -ErrorAction Stop |
                Where-Object {
                    $_.IsInherited -ne $true -and
                    $_.Deny -ne $true -and
                    $_.AccessRights -like '*SendAs*' -and
                    $_.Trustee -ne 'NT AUTHORITY\SELF' -and
                    $_.Trustee -notlike '*S-1-5*'
                } |
                Select-Object -Property @{Name = 'Folder'; Expression = { 'N/A' } },
                @{Name = 'Permission'; Expression = { 'Send-As' } },
                @{Name = 'User'; Expression = { $_.Trustee } }
            }
            catch {
                $currentMBXPermissionLookupFailures += [PSCustomObject]@{
                    Folder     = 'N/A'
                    Permission = 'SendAs'
                }
            }
        }
        #endregion SendAs

        #region SendOnBehalf

        if ($IncludePermissionTypes -match '(^All$)|(^SendOnBehalf$)') {

            if ($currentMBX.GrantSendOnBehalfTo) {

                foreach ($gsobt in $currentMBX.GrantSendOnBehalfTo) {

                    $currentMBXPermissions += $currentMBX.GrantSendOnBehalfTo |
                    Select-Object @{Name = 'Folder'; Expression = { 'N/A' } },
                    @{Name = 'Permission'; Expression = { 'Send on Behalf' } },
                    @{Name = 'User'; Expression = { $_ } }
                }
            }
        }
        #endregion SendOnBehalf

        #region Calendar

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^Calendar$)') {

            try {
                $currentMBXPermissions += Get-EXOMailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Calendar" -ErrorAction Stop |
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
                    Permission = 'N/A'
                }
            }
        }
        #endregion Calendar    

        #region Contacts

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^Contacts$)') {

            try {
                $currentMBXPermissions += Get-EXOMailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Contacts" -ErrorAction Stop |
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
                    Permission = 'N/A'
                }
            }
        }
        #endregion Contacts

        #region Tasks

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^Tasks$)') {

            try {
                $currentMBXPermissions += Get-EXOMailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Tasks" -ErrorAction Stop |
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
                    Permission = 'N/A'
                }
            }
        }
        #endregion Tasks

        #region MailboxRoot

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^\\$)') {

            try {
                $currentMBXPermissions += Get-EXOMailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\" -ErrorAction Stop |
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
                    Permission = 'N/A'
                }
            }
        }
        #endregion MailboxRoot

        #region Inbox

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^Inbox$)') {

            try {
                $currentMBXPermissions += Get-EXOMailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Inbox" -ErrorAction Stop |
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
                    Permission = 'N/A'
                }
            }
        }    
        #endregion Inbox

        #region SentItems

        if ($IncludePermissionTypes -match '(^All$)|(^CommonFolders$)|(^SentItems$)') {

            try {
                $currentMBXPermissions += Get-EXOMailboxFolderPermission -Identity "$($currentMBX.ExternalDirectoryObjectId):\Sent Items" -ErrorAction Stop |
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
                    Permission = 'N/A'
                }
            }
        }
        #endregion SentItems

        #region Post-processing

        foreach ($cmp in $currentMBXPermissions) {

            Write-Verbose -Message "Looking up trustee: '$($cmp.User)'"

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

                        $PSmtpOrName = 'UPN/PSmtp'
                        $Recipient = Get-EXORecipient -Identity $cmp.User -Properties DisplayName, PrimarySmtpAddress, RecipientTypeDetails, Guid, ExternalDirectoryObjectId -ErrorAction Stop
                    }
                    elseif ($cmp.User -match "'") {
                        # This elseif{} block is temporary, see the comments in the followin else{} block (and the .Notes help section).
                        # Do nothing, this user will be skipped in order to save on the error that it would occur due to the apostrophe in their name.
                    }
                    else {

                        $PSmtpOrName = 'Name'
                        # The following filter will not actually work in the new EXO V2 Get-EXORecipient/Get-EXOMailbox cmdlets.
                        # Mailbox folder permission objects store the user using the Name property.
                        # And so, you can't successfully filter for users with apostrophes in their Name anymore.
                        # I'm leaving it here because I've reported this and have my fingers crossed it's going to work again in the future.
                        # $Recipient = Get-EXORecipient -Filter "Name -eq '$($cmp.User -replace ""'"",""''"")'" -Properties DisplayName, PrimarySmtpAddress, RecipientTypeDetails, Guid, ExternalDirectoryObjectId -ErrorAction Stop
                        $Recipient = Get-EXORecipient -Filter "Name -eq '$($cmp.User)'" -Properties DisplayName, PrimarySmtpAddress, RecipientTypeDetails, Guid, ExternalDirectoryObjectId -ErrorAction Stop
                    }
                }
                catch {
                    Write-Warning -Message "Failed on Get-EXORecipient command."
                    Write-Warning -Message "Detected User ID property = $($PSmtpOrName)"
                    Write-Warning -Message "Mailbox (PSmtp) = '$($currentMBX.PrimarySmtpAddress)'"
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
    Write-Progress @Progress -Completed
}

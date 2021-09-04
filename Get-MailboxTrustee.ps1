<#
    .Synopsis
    Get all mailbox permissions:

    - FullAccess
    - Send-As
    - Send on Behalf

    ...and common mailbox folders' permissions:

    - \
    - Inbox
    - Calendar
    - Contacts
    - Tasks
    - Sent Items

    - Excludes mailbox permissions that are inherited, or for filtered trustee accounts.

    - Excludes mailbox folder permissions 'None' and 'AvailabilityOnly, as well as permissions for 'Default',
    'Anonymous', or 'Unknown' (deleted/disbaled) users.

    .Description
    This script has been tested against Exchange 2010, Exchange 2016.  It is intended to be used with regular Windows
    PowerShell.

    ***It is currently not intended for use with the Exchange Management Shell (yet/until this note is removed).  This
    is due to how Select-Object behaves with both -Property and -ExpandProperty used together.  Logic is going to be
    added to deal with this on the fly, at which time the Exch. Mgmt. Shell will work too.

    .Parameter Identity [string]
    Accepts pipeline input directly or by property name (caution not to pipe directly from Exchange cmdlets to avoid
    concurrent/busy pipeline errors (piping single objects is OK).  Instead, first store multiple objects in a
    variable, then pipe the variable (see examples).

    Accepted properties from the pipeline are (because other properties have proven to be failure-prone):

    - SamAccountName
    - Alias
    - PrimarySmtpAddress
    - Guid

    .Parameter FilterTrustees [string[]]
    Trustees to exclude permissions for. Review the $FilteredTrustees definition
    in the begin block to see which accounts are automatically filtered out.  Use
    the -FilterTrustees parameter to exclude additional trustees.


    .Parameter BypassPermissionTypes [string[]]
    Only the following values are accepted.  Multiple values can be comma-
    separated.

    - FullAccess
    - SendAs
    - SendOnBehalf
    - Folders

    .Outputs
    System.Management.Automation.PSCustomObject

    Default properties that are output:
    - DisplayName (mailbox')
    - PrimarySmtpAddress (mailbox')
    - RecipientTypeDetails (mailbox')
    - Guid (mailbox')
    - PermissionType (Mailbox, [Folder Name])
    - AccessRights (FullAccess, Send-As, Send on Behalf, [semicolon-delimited folder permissions' AccessRights])
    - TrusteeGuid (trustee's Guid)
    - TrusteeType (trustee's RecipientTypeDetails)
    - TrusteeDisplayName (trustee's DisplayName)
    - TrusteePSmtp (trustee's PrimarySmtpAddress)

    .Example
    # Via the pipeline (single object):

    Get-Mailbox User1@jb365.ca |
    .\Get-MailboxTrustee.ps1 -FilterTrustees "adds\besadmin", "*jsmith*" -ExpandTrusteeGroups -BypassPermissionTypes:Folders

    .Example
    # Via the pipeline (multiple objects):

    $Mailboxes = Get-Mailbox -ResultSize:Unlimited
    $Mailboxes | .\Get-MailboxTrustee.ps1

    .Example
    # Via direct call:

    .\Get-MailboxTrustee.ps1 User1@jb365.ca -BypassPermissionTypes:SendAs,SendOnBehalf

    .\Get-MailboxTrustee.ps1 [Enter]

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrustee.ps1

    .Link
    # New-MailboxTrusteeReverseLookup.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/New-MailboxTrusteeReverseLookup.ps1

    # Get-MailboxTrusteeWebSQLEdition.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrusteeWebSQLEdition.ps1

    # Get-MailboxTrusteeWeb.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrusteeWeb.ps1

    # Optimize-MailboxTrusteeWebInput.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Optimize-MailboxTrusteeWebInput.ps1
#>

#Requires -Version 3

[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
    [Alias('Alias', 'Guid', 'PrimarySmtpAddress', 'SamAccountName')]
    [string]$Identity,

    [string[]]$FilterTrustees,
    [switch]$ExpandTrusteeGroups,

    [ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf', 'Folders')]
    [string[]]$BypassPermissionTypes
)

begin {
    $StartTime = Get-Date
    $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity 'Get-MailboxTrustee.ps1' -Status 'Initializing...' -PercentComplete -1 -CurrentOperation "Command: ""$($PSCmdlet.MyInvocation.MyCommand)"""
    Write-Verbose -Message "Script Get-MailboxTrustee.ps1 begin ($($StartTime.DateTime)).`nCommand: ""$($PSCmdlet.MyInvocation.MyCommand)"""

    Write-Verbose -Message "Determining the connected Exchange environment."
    $ExPSSession = @()
    $ExPSSession += Get-PSSession |
    Where-Object {
        $_.ConfigurationName -eq 'Microsoft.Exchange' -and
        $_.State -eq 'Opened'
    }

    if ($ExPSSession.Count -eq 1) {

        # Check if we're in Exchange Online or On-Premises.
        switch ($ExPSSession.ComputerName) {

            outlook.office365.com { throw "This script is not intended for use with Exchange Online.  Use Get-EXOMailboxTrustee.ps1 for that." }

            default {
                # Set scope to entire forest (important for multi-domain forests).
                Set-ADServerSettings -ViewEntireForest:$true

                # Determine if the connected Exchange server's version is 2010 or newer.
                # The reason for this is that PowerShell 2.0's (i.e. Exchange 2010's)
                # Select-Object cannot provide both -Property and -ExpandedProperty
                # in one go.  So for 2010 sessions, we'll need to make two passes for
                # commands Get-MailboxPermission, Get-ADPermission,
                # Get-MailboxFolderPermission.  <--: UPDATE - testing with 2016 mgmt console has shown the problem is potentially with PS remoting, not Exchange 2010.  So this logic will be replaced with more dynamic logic in the script.
                $ExOnPSrv = Get-ExchangeServer -Identity "$($ExPSSession.ComputerName)"

                switch ($ExOnPSrv.AdminDisplayVersion) {

                    { $_ -match 'Version 14' } { $LegacyExchange = $true }
                    { $_ -match 'Version 15' } { $LegacyExchange = $false }
                    default {
                        throw "Unable to determine connect Exchange On-Premises server version."
                    }
                }
            }
        }
    }
    else {
        Write-Warning -Message "Requires a single** active (State: Opened) remote session to an Exchange (on-premises) server."
        break
    }

    # Define common parameter values for splatting with Invoke-Command throughout the script.
    $icCommon = @{
        Session          = $ExPSSession[0]
        HideComputerName = $true
        ErrorAction      = 'Continue'
    }

    # Define list of trustees to exclude.
    $FilteredTrustees = @(
        "*S-1-*",
        "BUILTIN\*"
        "*\Administrator",
        "*\Discovery Management",
        "*\Organization Management",
        "*\Domain Admins",
        "*\Enterprise Admins",
        "*\Exchange Services",
        "*\Exchange Trusted Subsystem",
        "*\Exchange Servers",
        "*\Exchange View-Only Administrators",
        "*\Exchange Admins",
        "*\Managed Availability Servers",
        "*\Public Folder Administrators",
        "*\Exchange Domain Servers",
        "*\Exchange Organization Administrators",
        "NT AUTHORITY\*",
        "*\JitUsers",
        "*\BESAdmin"
    )

    # Append additional trustees from -FilterTrustees parameter, then convert the array to a string pattern for use with -match throughout the script.
    $FilteredTrustees += $FilterTrustees
    $FilteredTrustees = foreach ($ft in $FilteredTrustees) { [regex]::Escape($ft) }
    $FilteredTrustees = '^(' + ($FilteredTrustees -join '|') + ')$'
    $FilteredTrustees = $FilteredTrustees -replace '\\\*', '.*'

    # To avoid searching for the same trustee more than once, track all trustees and index them by their SID's for fast lookup performance:
    $trusteeTracker = @{}

    function getTrusteeObject {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$trusteeSid,

            [Parameter(Mandatory = $true)]
            [psobject]$mailboxObject,

            [Parameter(Mandatory = $true)]
            [string]$accessRights,

            [switch]$folder,
            [string]$folderName,
            [switch]$expandGroups
        )

        Write-Verbose "[function: getTrusteeObject][Mailbox: $($mailboxObject.PrimarySmtpAddress)][Trustee: $($trusteeSid)][Folder: $($folder)][AccessRights: $($accessRights -join ',')]."

        # Initialize output object and falsify $trusteeGroupExpansionComplete variable.
        $trusteeObject = [PSCustomObject]@{}
        $trusteeGroupExpansionComplete = $false

        # Update the object with properties from our mailbox.
        $mailboxObject |
        Get-Member -MemberType Properties |
        ForEach-Object {

            $trusteeObject |
            Add-Member -NotePropertyName $_.Name -NotePropertyValue $mailboxObject.$($_.Name)
        }

        # Define which trustee properties to get.
        $trusteeProps = @(
            'RecipientTypeDetails',
            'DisplayName',
            'Name', #<--: For non-mail-enabled groups with no DisplayName set, we'll substitute in Name.
            'WindowsEmailAddress',
            'Guid'
        )

        # Next, we'll add the PermissionType and AccessRights properties to our output object.
        switch ($folder) {

            # If folder switch was used (i.e. $true), we're working with mailbox folder permissions.
            $true {
                $trusteeObject |
                Add-Member -NotePropertyName 'PermissionType' -NotePropertyValue $folderName -PassThru |
                Add-Member -NotePropertyName 'AccessRights' -NotePropertyValue $accessRights
            }

            # If folder switch was not used (i.e. $false), we're working with mailbox-level permissions.
            $false {
                switch ($accessRights) {

                    FullAccess {
                        $trusteeObject |
                        Add-Member -NotePropertyName 'PermissionType' -NotePropertyValue 'Mailbox' -PassThru |
                        Add-Member -NotePropertyName 'AccessRights' -NotePropertyValue 'FullAccess'
                    }

                    SendAs {
                        $trusteeObject |
                        Add-Member -NotePropertyName 'PermissionType' -NotePropertyValue 'Mailbox' -PassThru |
                        Add-Member -NotePropertyName 'AccessRights' -NotePropertyValue 'Send-As'
                    }

                    SendOnBehalf {
                        $trusteeObject |
                        Add-Member -NotePropertyName 'PermissionType' -NotePropertyValue 'Mailbox' -PassThru |
                        Add-Member -NotePropertyName 'AccessRights' -NotePropertyValue 'Send on Behalf'
                    }
                }
            } # end $folder -eq $false
        } # end switch ($folder)

        # Check if we've already seen this trustee and if so reuse the info:
        if (-not $trusteeTracker[$trusteeSid]) {

            # Start with a fresh $null $foundTrustee.
            $foundTrustee = $null

            # Determine if we've received a SID or Guid for $trusteeSid
            # Send on Behalf section sends a Guid because Exchange 2010's -ExandedProperty GrantSendOnBehalfTo doesn't include the SID.
            # Meanwhile Get-MailboxPermission and Get-ADPermission supply us with the SID but not the Guid.
            if ($trusteeSid -like 'S-1-5*') {

                $filter = "Sid -eq '$($trusteeSid)' -or SidHistory -eq '$($trusteeSid)'"
            }
            else { $filter = "Guid -eq '$($trusteeSid)" }

            $foundTrustee = Invoke-Command @icCommon -ScriptBlock {

                Get-User -Filter {$using:filter} -ErrorAction:SilentlyContinue |
                Select-Object $Using:trusteeProps
            }

            # If no trustee (user) was found, search groups.
            if ($null -eq $foundTrustee) {

                # Determine if group members should be resolved.
                switch ($expandGroups) {

                    # We are resolving group members (i.e. expanding groups).
                    $true {
                        Write-Verbose -Message "[function: getTrusteeObject][Mailbox: $($mailboxObject.PrimarySmtpAddress)] Searching for and expanding group $($trusteeSid)."

                        $foundTrusteeGroup = $null
                        $foundTrustees = @()

                        $foundTrusteeGroup = Invoke-Command @icCommon -ScriptBlock {

                            Get-Group -Filter {$using:filter} -ErrorAction:SilentlyContinue |
                            Select-Object $Using:trusteeProps
                        }

                        # Then get its members.
                        $foundTrustees += Invoke-Command @icCommon -ScriptBlock {

                            # -ReadFromDomainController allows us to get members from non-universal groups in other domains than the current user's.
                            Get-Group -Filter {$using:filter} -ReadFromDomainController -ErrorAction:SilentlyContinue |
                            Select-Object -ExpandProperty Members
                        }

                        # Send the group members back into getTrusteeObject for processing.
                        $foundTrustees |
                        Select-Object -Index (0..999) |
                        ForEach-Object {

                            # We define this here so we don't lose it with $_ when we enter the $folder switch.  $_.SecurityIdentifierString was another option, but isn't always available, while ObjectGuid seems to be available in all Exchange (2010 +).
                            $trusteeGroupMemberSid = $_.ObjectGuid.Guid

                            switch ($folder) {
                                $true {
                                    getTrusteeObject -trusteeSid $trusteeGroupMemberSid -mailboxObject $mailboxObject -accessRights $accessRights -folder -folderName $folderName -expandGroups:$true
                                }
                                $false {
                                    getTrusteeObject -trusteeSid $trusteeGroupMemberSid -mailboxObject $mailboxObject -accessRights $accessRights -expandGroups:$true
                                }
                            }
                        }

                        Write-Verbose -Message "[function: getTrusteeObject][Mailbox: $($mailboxObject.PrimarySmtpAddress)] Expanded group $($trusteeSid)."

                        if ($null -ne $foundTrusteeGroup) { $trusteeGroupExpansionComplete = $true }
                    } # end $expandGroups -eq $true

                    # We aren't expanding groups, so just search and return trustee group (if found).
                    $false {
                        $foundTrustee = Invoke-Command @icCommon -ScriptBlock {

                            Get-Group -Filter {$using:filter} -ErrorAction:SilentlyContinue |
                            Select-Object $Using:trusteeProps
                        }
                    } # end $expandGroups -eq $false
                } # end switch ($expandGroups)
            } # end if ($null -eq $foundTrustee) {}

            # If trustee was not found (or returned multiple matches), make note of this in the TrusteeGuid property.
            if (($null -eq $foundTrustee) -or ($foundTrustee.Count -gt 1)) {

                # But first, if it was a group that was expanded, report so in the TrusteeType property, and return the trustee group object.
                if ($trusteeGroupExpansionComplete -eq $true) {

                    $trusteeObject |
                    Add-Member -NotePropertyName 'TrusteeGuid' -NotePropertyValue $foundTrusteeGroup.Guid -PassThru |
                    Add-Member -NotePropertyName 'TrusteeType' -NotePropertyValue "EXPANDED:$($foundTrusteeGroup.RecipientTypeDetails.Value)" -PassThru |
                    Add-Member -NotePropertyName 'TrusteePSmtp' -NotePropertyValue $foundTrusteeGroup.WindowsEmailAddress

                    if ([string]::IsNullOrEmpty($foundTrusteeGroup.DisplayName)) {

                        $trusteeObject | Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue $foundTrusteeGroup.Name
                    }
                    else { $trusteeObject | Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue $foundTrusteeGroup.DisplayName }

                }
                else {
                    $trusteeObject |
                    Add-Member -NotePropertyName 'TrusteeGuid' -NotePropertyValue "Not found or ambiguous ($($trusteeSid))" -PassThru |
                    Add-Member -NotePropertyName 'TrusteeType' -NotePropertyValue '' -PassThru |
                    Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue '' -PassThru |
                    Add-Member -NotePropertyName 'TrusteePSmtp' -NotePropertyValue ''
                }
            }
            # Otherwise add the successfully found trustee's pertinent properties.
            else {
                $trusteeObject |
                Add-Member -NotePropertyName 'TrusteeGuid' -NotePropertyValue $foundTrustee.Guid -PassThru |
                Add-Member -NotePropertyName 'TrusteeType' -NotePropertyValue $foundTrustee.RecipientTypeDetails -PassThru |
                Add-Member -NotePropertyName 'TrusteePSmtp' -NotePropertyValue $foundTrustee.WindowsEmailAddress

                if ([string]::IsNullOrEmpty($foundTrustee.DisplayName)) {

                    $trusteeObject | Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue $foundTrustee.Name
                }
                else { $trusteeObject | Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue $foundTrustee.DisplayName }
            }

            # Save user trustees (skip group trustees) into $trusteeTracker:
            if (-not ($trusteeObject -match '(Group)')) {

                $trusteeTracker[$trusteeSid] = @{

                    TrusteeDisplayName = $trusteeObject.TrusteeDisplayName
                    TrusteePSmtp       = $trusteeObject.TrusteePSmtp
                    TrusteeType        = $trusteeObject.TrusteeType
                    TrusteeGuid        = $trusteeObject.Guid
                }
            }

        } # end if (-not $trusteeTracker[$trusteeSid]) {}
        else {
            # This trustee was previously seen, reusing the info (for performance-savings):
            $trusteeObject |
                Add-Member -NotePropertyName 'TrusteeGuid' -NotePropertyValue $trusteeTracker[$trusteeSid].TrusteeGuid -PassThru |
                Add-Member -NotePropertyName 'TrusteeType' -NotePropertyValue $trusteeTracker[$trusteeSid].TrusteeType -PassThru |
                Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue $trusteeTracker[$trusteeSid].TrusteeDisplayName -PassThru |
                Add-Member -NotePropertyName 'TrusteePSmtp' -NotePropertyValue $trusteeTracker[$trusteeSid].TrusteePSmtp
        }

        # Finally, return the finished product.
        $trusteeObject |
        Select-Object -Property DisplayName, PrimarySmtpAddress, RecipientTypeDetails, Guid,
        TrusteeDisplayName, TrusteePSmtp, TrusteeType, TrusteeGuid

    } # end function getTrusteeObject

    # Fire up the engines.  Brap brap...
    $MailboxProcessedCounter = 0
    $MainProgress = @{
        Activity         = "Get-MailboxTrustee.ps1 - Start time: $($StartTime.DateTime)"
        Status           = "Mailboxes processed: $($MailboxProcessedCounter) ; Time elapsed $($StopWatch.Elapsed -replace '\..*')"
        Id               = 0
        ParentId         = -1
        PercentComplete = -1
    }
    Write-Progress @MainProgress

    $faProgressProps = @{
        Activity         = 'FullAccess'
        Id               = 1
        ParentId         = 0
        PercentComplete = -1
    }
    Write-Progress @faProgressProps -Status 'Ready'

    $saProgressProps = @{
        Activity         = 'Send-As'
        Id               = 2
        ParentId         = 0
        PercentComplete = -1
    }
    Write-Progress @saProgressProps -Status 'Ready'

    $sobProgressProps = @{
        Activity         = 'Send on Behalf'
        Id               = 3
        ParentId         = 0
        PercentComplete = -1
    }
    Write-Progress @sobProgressProps -Status 'Ready'

    $fpProgressProps = @{
        Activity         = 'Folder Permissions'
        Id               = 4
        ParentId         = 0
        PercentComplete = -1
    }
    Write-Progress @fpProgressProps -Status 'Ready'
}

process {
    # Placing all of the process block inside this single try block.  The purpose is to kill the script when something breaking occurs.  Use Verbose/Debug for troubleshooting.
    try {
        #-----------------------------------#
        #----- Initial Mailbox Lookoup -----#
        #-----------------------------------#

        Write-Verbose -Message "Getting mailbox with identity '$($Identity)'."

        $Mailbox = $null
        $Mailbox = Invoke-Command -Session $ExPSSession[0] -HideComputerName -ErrorAction SilentlyContinue -ScriptBlock {

            Get-Mailbox -Identity $Using:Identity -WarningAction:SilentlyContinue -ErrorAction:Stop |
            Select-Object -Property DisplayName, PrimarySmtpAddress, RecipientTypeDetails, Guid, GrantSendOnBehalfTo
        }
        if ($null -eq $Mailbox) {

            Write-Warning -Message "Failed to find a mailbox (via Get-Mailbox) for identity '$($Identity)'."
            return
        }

        # Store the mailbox' PrimarySmtpAddress for use with Write-Progress|Verbose|Debug.
        # Store the mailbox' Guid for use with -Identiy parameters.
        [string]$mPSmtp = $Mailbox.PrimarySmtpAddress
        [string]$mGuid = $Mailbox.Guid

        Write-Verbose -Message "[Mailbox: $($mPSmtp)] Mailbox lookup complete."

        $MainProgress['Status'] = "Mailboxes processed: $($MailboxProcessedCounter) ; Time elapsed $($StopWatch.Elapsed -replace '\..*')"
        Write-Progress @MainProgress -CurrentOperation "Current mailbox: $($Mailbox.DisplayName) ($($mPSmtp))"
        $MailboxProcessedCounter++

        #-----------------------------------#
        #----------- FullAccess ------------#
        #-----------------------------------#

        if ($BypassPermissionTypes -notcontains 'FullAccess') {

            Write-Progress @faProgressProps -Status 'Getting FullAccess permissions with Get-MailboxPermission'
            Write-Verbose -Message "[Mailbox: $($mPSmtp)] $($faProgressProps.Status)."

            switch ($LegacyExchange) {
                $true {
                    $MailboxPermissions = @()
                    $MailboxPermissions += Invoke-Command @icCommon -ScriptBlock {

                        Get-MailboxPermission -Identity $Using:mGuid -ErrorAction:SilentlyContinue |
                        Select-Object User, AccessRights, IsInherited, Deny
                    }

                    # Apply filters.
                    $FullAccess = @()
                    $FullAccess += $MailboxPermissions |
                    Where-Object {
                        -not
                        ($_.User -match $FilteredTrustees) -and
                        ($_.AccessRights -like '*FullAccess*') -and
                        ($_.IsInherited -eq $false) -and
                        ($_.Deny -ne $true) # <--: Comes back empty instead of $false when remoting.
                    }

                    $faIndex = @()
                    $FullAccess |
                    ForEach-Object {
                        $faIndex += $MailboxPermissions.IndexOf($_)
                    }

                    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Expanding 'User' property for all FullAccess permissions."

                    $faTrustees = @()
                    $faTrustees += Invoke-Command @icCommon -ScriptBlock {

                        Get-MailboxPermission -Identity $Using:mGuid -ErrorAction:SilentlyContinue |
                        Select-Object -Index $Using:faIndex |
                        Select-Object -ExpandProperty User |
                        Select-Object -Property SecurityIdentifier
                    }

                    if ($FullAccess.Count -ge 1) {
                        Write-Verbose -Message "[Mailbox: $($mPSmtp)] Processing FullAccess trustees."
                    }

                    Write-Progress @faProgressProps -Status 'Resolving trustees with Get-User/Get-Group'
                    $FullAccess |
                    Select-Object -Index (0..999) |
                    ForEach-Object {

                        $getTrusteeObjectProps = @{

                            trusteeSid    = $faTrustees[$($faIndexCounter)].SecurityIdentifier.Value
                            mailboxObject = $Mailbox
                            accessRights  = 'FullAccess'
                            expandGroup   = $ExpandTrusteeGroups
                        }
                        getTrusteeObject @getTrusteeObjectProps
                    }
                } # end $LegacyExchange -eq $true

                $false {
                    $MailboxPermissions = @()
                    $MailboxPermissions += Invoke-Command @icCommon -ScriptBlock {

                        Get-MailboxPermission -Identity $Using:mGuid -ErrorAction:SilentlyContinue |
                        Select-Object AccessRights, IsInherited, Deny -ExpandProperty User
                    }

                    # Apply filters.
                    $FullAccess = @()
                    $FullAccess += $MailboxPermissions |
                    Where-Object {
                        -not
                        ($_.RawIdentity -match $FilteredTrustees) -and
                        ($_.AccessRights -like '*FullAccess*') -and
                        ($_.IsInherited -eq $false) -and
                        ($_.Deny -ne $true) # <--: Comes back empty instead of $false when remoting.
                    }

                    if ($FullAccess.Count -ge 1) {
                        Write-Verbose -Message "[Mailbox: $($mPSmtp)] Processing FullAccess trustees."
                    }

                    Write-Progress @faProgressProps -Status 'Resolving trustees with Get-User/Get-Group'
                    $FullAccess |
                    Select-Object -Index (0..999) |
                    ForEach-Object {

                        $getTrusteeObjectProps = @{

                            trusteeSid    = $_.SecurityIdentifier
                            mailboxObject = $Mailbox
                            accessRights  = 'FullAccess'
                            expandGroup   = $ExpandTrusteeGroups
                        }
                        getTrusteeObject @getTrusteeObjectProps
                    }
                } # end $LegacyExchange -eq $false
            } # end switch ($LegacyExchange)

            Write-Progress @faProgressProps -Status 'Ready'
            Write-Verbose -Message "[Mailbox: $($mPSmtp)] FullAccess discovery complete."

        } # end if ($BypassPermissionTypes -notcontains 'FullAccess')

        #-----------------------------------#
        #------------- Send-As -------------#
        #-----------------------------------#

        if ($BypassPermissionTypes -notcontains 'SendAs') {

            Write-Progress @saProgressProps -Status 'Getting Send-As permissions with Get-ADPermission'
            Write-Verbose -Message "[Mailbox: $($mPSmtp)] $($saProgressProps.Status)."

            switch ($LegacyExchange) {
                $true {
                    $ADPermissions = @()
                    $ADPermissions += Invoke-Command @icCommon -ScriptBlock {

                        Get-ADPermission -Identity $Using:mGuid -ErrorAction:SilentlyContinue |
                        Select-Object User, ExtendedRights, IsInherited, Deny
                    }

                    # Apply filters.
                    $SendAs = @()
                    $SendAs += $ADPermissions |
                    Where-Object {
                        -not
                        ($_.User -match $FilteredTrustees) -and
                        ($_.ExtendedRights -like '*Send-As*') -and
                        ($_.IsInherited -eq $false) -and
                        ($_.Deny -ne $true) # <--: Comes back empty instead of $false when remoting.
                    }

                    $saIndex = @()
                    $SendAs |
                    ForEach-Object {
                        $saIndex += $ADPermissions.IndexOf($_)
                    }

                    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Expanding 'User' property for all SendAs permissions."

                    $saTrustees = @()
                    $saTrustees += Invoke-Command @icCommon -ScriptBlock {

                        Get-ADPermission -Identity $Using:mGuid -ErrorAction:SilentlyContinue |
                        Select-Object -Index $Using:saIndex |
                        Select-Object -ExpandProperty User |
                        Select-Object -Property SecurityIdentifier
                    }

                    if ($SendAs.Count -ge 1) {
                        Write-Verbose -Message "[Mailbox: $($mPSmtp)] Processing Send-As trustees."
                    }

                    Write-Progress @saProgressProps -Status 'Resolving trustees with Get-User/Get-Group'
                    $SendAs |
                    Select-Object -Index (0..999) |
                    ForEach-Object {

                        $getTrusteeObjectProps = @{

                            trusteeSid    = $saTrustees[$($saIndexCounter)].SecurityIdentifier.Value
                            mailboxObject = $Mailbox
                            accessRights  = 'SendAs'
                            expandGroup   = $ExpandTrusteeGroups
                        }
                        getTrusteeObject @getTrusteeObjectProps
                    }
                } # end switch ($LegacyExchange) { $true {*} }

                $false {
                    $ADPermissions = Invoke-Command @icCommon -ScriptBlock {

                        Get-ADPermission -Identity $Using:mGuid -ErrorAction:SilentlyContinue |
                        Select-Object ExtendedRights, IsInherited, Deny -ExpandProperty User
                    }

                    # Apply filters.
                    $SendAs = @()
                    $SendAs += $ADPermissions |
                    Where-Object {
                        -not
                        ($_.RawIdentity -match $FilteredTrustees) -and
                        ($_.ExtendedRights -like '*Send-As*') -and
                        ($_.IsInherited -eq $false) -and
                        ($_.Deny -ne $true) # <--: Comes back empty instead of $false when remoting.
                    }

                    if ($SendAs.Count -ge 1) {
                        Write-Verbose -Message "[Mailbox: $($mPSmtp)] Processing SendAs trustees."
                    }

                    Write-Progress @saProgressProps -Status 'Resolving trustees with Get-User/Get-Group'
                    $SendAs |
                    Select-Object -Index (0..999) |
                    ForEach-Object {

                        $getTrusteeObjectProps = @{

                            trusteeSid    = $_.SecurityIdentifier
                            mailboxObject = $Mailbox
                            accessRights  = 'SendAs'
                            expandGroup   = $ExpandTrusteeGroups
                        }
                        getTrusteeObject @getTrusteeObjectProps
                    }
                }  # end switch ($LegacyExchange) { $false {*} }
            } # end switch ($LegacyExchange) {*}

            Write-Progress @saProgressProps -Status 'Ready'
            Write-Verbose -Message "[Mailbox: $($mPSmtp)] Send-As discovery complete."

        } # end if ($BypassPermissionTypes -notcontains 'SendAs') {}

        #-----------------------------------#
        #---------- Send on Behalf ---------#
        #-----------------------------------#

        if ($BypassPermissionTypes -notcontains 'SendOnBehalf') {

            Write-Progress @sobProgressProps -Status 'Checking for Send on Behalf trustees in GrantSendOnBehalfTo property'
            Write-Verbose -Message "[Mailbox: $($mPSmtp)] $($sobProgressProps.Status)."

            if ($Mailbox.GrantSendOnBehalfTo.Count -ge 1) {

                Write-Verbose -Message "[Mailbox: $($mPSmtp)] Processing Send on Behalf trustees."

                $sobTrustees = Invoke-Command @icCommon -ScriptBlock {

                    Get-Mailbox -Identity "$($Using:mGuid)" -WarningAction:SilentlyContinue -ErrorAction:SilentlyContinue |
                    Select-Object -ExpandProperty GrantSendOnBehalfTo
                }

                Write-Progress @sobProgressProps -Status 'Resolving trustees with Get-User/Get-Group'
                $sobTrustees |
                Select-Object -Index (0..999) |
                ForEach-Object {

                    if (-not ([string]::IsNullOrEmpty($_.ObjectGuid.Guid))) {

                        getTrusteeObject -trusteeSid $_.ObjectGuid.Guid -mailboxObject $Mailbox -accessRights 'SendOnBehalf' -expandGroups $ExpandTrusteeGroups
                    }
                } # sobTrustees | ForEach-Object {}
            } # end if ($Mailbox.GrantSendOnBehalfTo.Count -ge 1) {}

            Write-Progress @sobProgressProps -Status 'Ready'
            Write-Verbose -Message "[Mailbox: $($mPSmtp)] Send of Behalf discovery complete."

        } # end if ($BypassPermissionTypes -notcontains 'SendOnBehalf') {}

        #-----------------------------------#
        #-------- Folder Permissions--------#
        #-----------------------------------#

        if ($BypassPermissionTypes -notcontains 'Folders') {

            # List mailbox folders as an array and send them down the pipeline.
            $Folders = @(
                'Mailbox root',
                'Inbox',
                'Calendar',
                'Contacts',
                'Sent Items',
                'Tasks'
            )

            $Folders | ForEach-Object {

                # Store the current folder name for use throughout the pipeline.
                [string]$Folder = "$($_)"

                Write-Progress @fpProgressProps -Status "Resolving $($Folder) permissions with Get-MailboxFolderPermission"
                Write-Verbose -Message "[Mailbox: $($mPSmtp)][Folder: $($Folder)] $($fpProgressProps.Status)."

                switch ($LegacyExchange) {
                    $true {
                        $FolderPermissions = @()
                        $FolderPermissions += Invoke-Command -ScriptBlock {

                            Get-MailboxFolderPermission -Identity "$($args[0]):\$($args[1])" -ErrorAction:SilentlyContinue |
                            Select-Object User, AccessRights

                        } @icCommon -ArgumentList $mGuid, "$($Folder -replace 'Mailbox root','')"

                        # Apply filters.
                        $PertinentFolderPermissions = @()
                        $PertinentFolderPermissions += $FolderPermissions |
                        Where-Object {
                            -not
                            ($_.User -like '*Default*') -and -not
                            ($_.User -like '*Anonymous*') -and -not
                            ($_.User -like '*S-1-5-*') -and -not
                            ($_.AccessRights -like '*None*') -and -not
                            ($_.AccessRights -like '*AvailabilityOnly*')
                        }

                        $pfpIndex = @()
                        $PertinentFolderPermissions |
                        ForEach-Object {
                            $pfpIndex += $FolderPermissions.IndexOf($_)
                        }

                        Write-Verbose -Message "[Mailbox: $($mPSmtp)][Folder: $($Folder)] Expanding 'User' property for all folder permissions."

                        $fpTrustees = @()
                        $fpTrustees += Invoke-Command -ScriptBlock {

                            Get-MailboxFolderPermission -Identity "$($args[0]):\$($args[1])" -ErrorAction:SilentlyContinue |
                            Select-Object -Index $Using:pfpindex |
                            Select-Object -ExpandProperty User |
                            Select-Object -Property ADRecipient

                        } @icCommon -ArgumentList $mGuid, "$($Folder -replace 'Mailbox root','')", $pfpIndex


                        if ($PertinentFolderPermissions.Count -ge 1) {
                            Write-Verbose -Message "[Mailbox: $($mPSmtp)][Folder: $($Folder)] Processing folder permission trustees."
                        }

                        Write-Progress @fpProgressProps -Status "Resolving $($Folder) trustees with Get-User/Get-Group"
                        $PertinentFolderPermissions |
                        Select-Object -Index (0..999) |
                        ForEach-Object {

                            if (-not ([string]::IsNullOrEmpty("$($fpTrustees[$($pfpIndexCounter)].ADRecipient.Sid)"))) {

                                $getTrusteeObjectProps = @{

                                    trusteeSid    = $fpTrustees[$($pfpIndexCounter)].ADRecipient.Sid
                                    mailboxObject = $Mailbox
                                    accessRights  = $_.AccessRights -join ';'
                                    folder        = $true
                                    folderName    = $Folder
                                    expandGroup   = $ExpandTrusteeGroups
                                }
                                getTrusteeObject @getTrusteeObjectProps
                            }
                        } # end $PertinentFolderPermissions | ForEach-Object {}
                    } # end $LegacyExchange -eq $true

                    $false {
                        $FolderPermissions = @()
                        $FolderPermissions += Invoke-Command -ScriptBlock {

                            Get-MailboxFolderPermission -Identity "$($args[0]):\$($args[1])" -ErrorAction:SilentlyContinue |
                            Select-Object User, AccessRights -ExpandProperty User

                        } @icCommon -ArgumentList $mGuid, "$($Folder -replace 'Mailbox root','')"

                        # Apply filters.
                        $PertinentFolderPermissions = @()
                        $PertinentFolderPermissions += $FolderPermissions |
                        Where-Object {
                            -not
                            ($_.User -like '*Default*') -and -not
                            ($_.User -like '*Anonymous*') -and -not
                            ($_.User -like '*S-1-5-*') -and -not
                            ($_.AccessRights -like '*None*') -and -not
                            ($_.AccessRights -like '*AvailabilityOnly*')
                        }

                        if ($PertinentFolderPermissions.Count -ge 1) {
                            Write-Verbose -Message "[Mailbox: $($mPSmtp)][Folder: $($Folder)] Processing folder permission trustees."
                        }

                        Write-Progress @fpProgressProps -Status "Resolving $($Folder) trustees with Get-User/Get-Group"
                        $PertinentFolderPermissions |
                        Select-Object -Index (0..999) |
                        ForEach-Object {

                            if (-not ([string]::IsNullOrEmpty("$($_.ADRecipient.Sid)"))) {

                                $getTrusteeObjectProps = @{

                                    trusteeSid    = $_.ADRecipient.Sid
                                    mailboxObject = $Mailbox
                                    accessRights  = $_.AccessRights -join ';'
                                    folder        = $true
                                    folderName    = $Folder
                                    expandGroup   = $ExpandTrusteeGroups
                                }
                                getTrusteeObject @getTrusteeObjectProps
                            }
                        } # end $PertientFolderPermissions | ForEach-Object {}
                    } # end $LegacyExchange -eq $false
                } # end switch ($LegacyExchange)
            } # end $Folders | ForEach-Object {}

            Write-Progress @fpProgressProps -Status 'Ready'
            Write-Verbose -Message "[Mailbox: $($mPSmtp)] Folder permissions discovery complete."

        } # end if ($BypassPermissionTypes -notcontains 'Folders')

        #-----------------------------------#
        #--------- Process Wrap-Up ---------#
        #-----------------------------------#

    } # Close try block.

    # Session problems go here (and terminate the script):
    catch [System.Management.Automation.CommandNotFoundException],
    [System.Management.Automation.Remoting.PSRemotingTransportException] {

        'Session problems (most likely) have caused the script to fail.  ' +
        "Mailboxes processed/total (if available): $($MyInvocation.PipelinePosition) / $($MyInvocation.PipelineLength)" | Write-Warning
        throw
    }

    # Other problems go here, and do not terminate the script, just this mailbox:
    catch {
        'A problem has caused the script to fail (error to follow):' | Write-Warning
        Write-Error $_
        "Mailboxes processed: $($MailboxProcessedCounter)" | Write-Warning
        "Date/time: $(Get-Date -Format G)`n" +
        'Moving onto the next mailbox, if any remain.' | Write-Warning
    }
} # end process

End {
    Write-Progress @fpProgressProps -Completed
    Write-Progress @sobProgressProps -Completed
    Write-Progress @saProgressProps -Completed
    Write-Progress @faProgressProps -Completed
    Write-Progress @MainProgress -Completed
    Write-Verbose -Message "Script Get-MailboxTrustee.ps1 end."
}

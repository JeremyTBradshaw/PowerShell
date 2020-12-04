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

    - Excludes mailbox permissions that are inherited, or for filtered trustee
    accounts.

    - Excludes mailbox folder permissions 'None' and 'AvailabilityOnly, as well
    as permissions for 'Default', 'Anonymous', or 'Unknown' (deleted/disbaled)
    users.


    .Description

    This script has been tested against Exchange 2010, Exchange 2016, and
    Exchange Online.  It is intended to be used with regular Windows PowerShell.

    ***It is currently not intended for use with the Exchange Management Shell
    (yet/until this note is removed).  This is due to how Select-Object behaves
    with both -Property and -ExpandProperty used together.  Logic is going to be
    added to deal with this on the fly, at which time the Exch. Mgmt. Shell will
    work too.


    .Parameter Identity [string]

    Accepts pipeline input directly or by property name (caution not to pipe
    directly from Exchange cmdlets to avoid concurrent/busy pipeline errors
    (piping single objects is OK).  Instead, first store multiple objects in a
    variable, then pipe the variable (see examples).

    Accepted properties from the pipeline are (because other properties have
    proven to be failure-prone):

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


    .Parameter MinimizeOutput [switch]

    When used, only the essential properties will be returned for both the
    mailbox and the trustee. Implies $ExpandTrusteeGroups:$true.


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

    Properties that are output when -MinimizeOutput switch is used:
    - Guid (mailbox' Guid)
    - PermissionType (same as above)
    - AccessRights (same as above)
    - TrusteeGuid (trustee's Guid)


    .Example

    # Via the pipeline (single object):

    Get-Mailbox User1@jb365.ca |
    .\Get-MailboxTrustee.ps1 -FilterTrustees "adds\besadmin", "*jsmith*" -ExpandTrusteeGroups -BypassPermissionTypes:Folders


    .Example

    # Via the pipeline (multiple objects):

    $Mailboxes = Get-Mailbox -ResultSize:Unlimited
    $Mailboxes | .\Get-MailboxTrustee.ps1 -MinimizeOutput


    .Example

    # Via direct call:

    .\Get-MailboxTrustee.ps1 User1@jb365.ca -BypassPermissionTypes:SendAs,SendOnBehalf

    .\Get-MailboxTrustee.ps1 [Enter]


    .Link

    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrustee.ps1
    # ^ Get-MailboxTrustee.ps1


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
    [Alias(
        'Alias',
        'Guid',
        'PrimarySmtpAddress',
        'SamAccountName')]
    [string]$Identity,

    [string[]]$FilterTrustees,

    # Making -MinimizeOutput imply -ExpandTrusteeGroups.
    [switch]$ExpandTrusteeGroups = $MinimizeOutput,

    [ValidateSet(
        'FullAccess',
        'SendAs',
        'SendOnBehalf',
        'Folders')]
    [string[]]$BypassPermissionTypes,

    [switch]$MinimizeOutput

)

begin {

    Write-Debug -Message "begin {}"
    $StartTime = Get-Date
    Write-Progress -Activity 'Get-MailboxTrustee.ps1' -Status 'Initializing...' -SecondsRemaining -1 -CurrentOperation "Command: ""$($PSCmdlet.MyInvocation.MyCommand)"""
    Write-Verbose -Message "Script Get-MailboxTrustee.ps1 begin ($($StartTime.DateTime)).`nCommand: ""$($PSCmdlet.MyInvocation.MyCommand)"""

    # Attempt to get a count of objects in the pipeline.
    $PipelineObjectCount = $null

    # Suppressing this one block's errors.
    try {
        $MyInvocationLineSplit = $PSCmdlet.MyInvocation.Line -split '\|'
        $MyInvocationLineSplit = $MyInvocationLineSplit -replace 'cls', '' -replace 'Clear-Host', '' -replace '.*=', ''

        $InvocationNameMatch = $MyInvocationLineSplit |
        Where-Object { $_ -like "*$($PSCmdlet.MyInvocation.InvocationName)*" }

        $InvocationNameIndex = $MyInvocationLineSplit.IndexOf($InvocationNameMatch)
        $MeasureableCommand = ($MyInvocationLineSplit | Select-Object -Index (0..($InvocationNameIndex - 1))) -join '|'
        $PipelineObjectCount = (Invoke-Expression -Command $MeasureableCommand | Measure-Object).Count
    }
    catch { <#Suppressed#> }

    if ($PipelineObjectCount -is [int] -and $PipelineObjectCount -gt 0) {

        $MailboxCounter = 0
        $MailboxPercentCompletePossible = $true
    }

    Write-Verbose -Message "Determining the connected Exchange environment."

    $ExPSSession = @()
    $ExPSSession += Get-PSSession |
    Where-Object {
        $_.ConfigurationName -eq 'Microsoft.Exchange' -and
        $_.State -eq 'Opened'
    }

    if ($ExPSSession.Count -eq 1) {
        $Exchange = $null

        # Check if we're in Exchange Online or On-Premises.
        switch ($ExPSSession.ComputerName) {

            outlook.office365.com {
                $Exchange = 'Exchange Online'
                $LegacyExchange = $false
            }

            default {
                $Exchange = 'Exchange On-Premises'

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
        Write-Verbose -Message "Connected environment is $($Exchange)."
    }

    else {
        Write-Warning -Message "Requires a single** active (State: Opened) remote session to an Exchange server or EXO."
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


    function getTrustee {
        <#
        .SYNOPSIS
        Helper function to get pertinent details for mailbox and folder permission
        trustees (i.e. users/groups).
    #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$trusteeSid,

            [Parameter(Mandatory = $true)]
            [psobject]$mailboxObject,

            [Parameter(Mandatory = $true)]
            [string]$accessRights,

            [switch]$folder,
            [string]$folderName, # <--: May eventually be converted to a dynamic parameter.
            [switch]$expandGroups
        )

        Write-Verbose "[function: getTrustee][Mailbox: $($mailboxObject.PrimarySmtpAddress)][Trustee: $($trusteeSid)][Folder: $($folder)][AccessRights: $($accessRights -join ',')]."

        # Initialize output object and falsify $trusteeGroupExpansionComplete variable.
        $trusteeReturned = [PSCustomObject]@{placeholder = $null }
        $trusteeGroupExpansionComplete = $false

        # Update the object with properties from our mailbox.  While we're in the $MinimizeOutput switch, let's define which trustee properties to get.
        switch ($MinimizeOutput) {
            $true {
                $trusteeReturned |
                Add-Member -NotePropertyName 'Guid' -NotePropertyValue $mailboxObject.Guid

                $trusteeProps = @('Guid')
            }
            $false {
                $mailboxObject |
                Get-Member -MemberType Properties |
                ForEach-Object {

                    $trusteeReturned |
                    Add-Member -NotePropertyName $_.Name -NotePropertyValue $mailboxObject.$($_.Name)
                }

                $trusteeProps = @(
                    'RecipientTypeDetails',
                    'DisplayName',
                    'WindowsEmailAddress',
                    'Guid'
                )
            }
        }

        # Next, we'll add the PermissionType and AccessRights properties to our output object.
        switch ($folder) {

            # If folder switch was used (i.e. $true), we're working with mailbox folder permissions.
            $true {
                $trusteeReturned |
                Add-Member -NotePropertyName 'PermissionType' -NotePropertyValue $folderName -PassThru |
                Add-Member -NotePropertyName 'AccessRights' -NotePropertyValue $accessRights
            }

            # If folder switch was not used (i.e. $false), we're working with mailbox-level permissions.
            $false {
                switch ($accessRights) {

                    FullAccess {
                        $trusteeReturned |
                        Add-Member -NotePropertyName 'PermissionType' -NotePropertyValue 'Mailbox' -PassThru |
                        Add-Member -NotePropertyName 'AccessRights' -NotePropertyValue 'FullAccess'
                    }

                    SendAs {
                        $trusteeReturned |
                        Add-Member -NotePropertyName 'PermissionType' -NotePropertyValue 'Mailbox' -PassThru |
                        Add-Member -NotePropertyName 'AccessRights' -NotePropertyValue 'Send-As'
                    }

                    SendOnBehalf {
                        $trusteeReturned |
                        Add-Member -NotePropertyName 'PermissionType' -NotePropertyValue 'Mailbox' -PassThru |
                        Add-Member -NotePropertyName 'AccessRights' -NotePropertyValue 'Send on Behalf'
                    }
                }
            } # end $folder -eq $false
        } # end switch ($folder)

        # Start with a fresh $null $foundTrustee.
        $foundTrustee = $null

        # Determine if we've received a SID or Guid for $trusteeSid
        # Send on Behalf section sends a Guid because Exchange 2010's -ExandedProperty GrantSendOnBehalfTo doesn't include the SID.
        # Meanwhile Get-MailboxPermission and Get-ADPermission supply us with the SID but not the Guid.
        if ($trusteeSid -like 'S-1-5*') {
            $getTrusteeUseSid = $true
        }
        else {
            $getTrusteeUseSid = $false
        }

        switch ($getTrusteeUseSid) {

            # $trusteeSid came in as a SID.
            $true {
                $foundTrustee = Invoke-Command @icCommon -ScriptBlock {

                    Get-User -Filter "Sid -eq '$($Using:trusteeSid)' -or SidHistory -eq '$($Using:trusteeSid)'" -ErrorAction:SilentlyContinue |
                    Select-Object $Using:trusteeProps
                }
            }

            # $trusteeSid came in as a Guid (i.e. from Send on Behalf section).
            $false {
                $foundTrustee = Invoke-Command @icCommon -ScriptBlock {

                    Get-User -Filter "Guid -eq '$($Using:trusteeSid)'" -ErrorAction:SilentlyContinue |
                    Select-Object $Using:trusteeProps
                }
            }
        } # end switch ($getTrusteeUseSid) (#1)

        # If no trustee (user) was found, search groups.
        if ($null -eq $foundTrustee) {

            # Determine if group members should be resolved.
            switch ($expandGroups) {

                # We are resolving group members (i.e. expanding groups).
                $true {

                    Write-Verbose -Message "[function: getTrustee][Mailbox: $($mailboxObject.PrimarySmtpAddress)] Searching for and expanding group $($trusteeSid)."

                    $foundTrusteeGroup = $null
                    $foundTrustees = @()

                    switch ($getTrusteeUseSid) {
                        # ugh, this guy again..

                        $true {

                            # Get the current group's details.
                            $foundTrusteeGroup = Invoke-Command @icCommon -ScriptBlock {

                                Get-Group -Filter "Sid -eq '$($Using:trusteeSid)' -or SidHistory -eq '$($Using:trusteeSid)'" -ErrorAction:SilentlyContinue |
                                Select-Object $Using:trusteeProps
                            }

                            # Then get its members.
                            if ($Exchange -eq 'Exchange On-Premises') {

                                $foundTrustees += Invoke-Command @icCommon -ScriptBlock {

                                    # -ReadFromDomainController allows us to get members from non-universal groups in other domains than the current user's.
                                    Get-Group -Filter "Sid -eq '$($Using:trusteeSid)' -or SidHistory -eq '$($Using:trusteeSid)'" -ReadFromDomainController -ErrorAction:SilentlyContinue |
                                    Select-Object -ExpandProperty Members
                                }
                            }
                            else {
                                $foundTrustees += Invoke-Command @icCommon -ScriptBlock {

                                    Get-Group -Filter "Sid -eq '$($Using:trusteeSid)' -or SidHistory -eq '$($Using:trusteeSid)'" -ErrorAction:SilentlyContinue |
                                    Select-Object -ExpandProperty Members
                                }
                            }
                        }

                        $false {

                            Write-Debug "getTrusteeUseSid equals false"
                        
                            # Same thing, get the current group's details.
                            $foundTrusteeGroup = Invoke-Command @icCommon -ScriptBlock {



                                Get-Group -Filter "Guid -eq '$($Using:trusteeSid)'" -ErrorAction:SilentlyContinue |
                                Select-Object $Using:trusteeProps
                            }

                            # Then get its members.
                            if ($Exchange -eq 'Exchange On-Premises') {

                                $foundTrustees += Invoke-Command @icCommon -ScriptBlock {

                                    # -ReadFromDomainController allows us to get members from non-universal groups in other domains than the current user's.
                                    Get-Group -Filter "Guid -eq '$($Using:trusteeSid)'" -ReadFromDomainController -ErrorAction:SilentlyContinue |
                                    Select-Object -ExpandProperty Members
                                }
                            }
                            else {
                                Get-Group -Filter "Guid -eq '$Using:trusteeSid'" -ErrorAction:SilentlyContinue |
                                Select-Object -ExpandProperty Members
                            }
                        }
                    } # end switch ($getTrusteeUseSid) (#2)

                    # Send the group members back into getTrustee for processing.
                    $foundTrustees |
                    Select-Object -Index (0..999) |
                    ForEach-Object {

                        # We define this here so we don't lose it with $_ when we enter the $folder switch.  $_.SecurityIdentifierString was another option, but isn't always available, while ObjectGuid seems to be available in all Exchange (2010 +).
                        $trusteeGroupMemberSid = $_.ObjectGuid.Guid

                        switch ($folder) {
                            $true {
                                getTrustee -trusteeSid $trusteeGroupMemberSid -mailboxObject $mailboxObject -accessRights $accessRights -folder -folderName $folderName -expandGroups:$true
                            }
                            $false {
                                getTrustee -trusteeSid $trusteeGroupMemberSid -mailboxObject $mailboxObject -accessRights $accessRights -expandGroups:$true
                            }
                        }
                    }

                    Write-Verbose -Message "[function: getTrustee][Mailbox: $($mailboxObject.PrimarySmtpAddress)] Expanded group $($trusteeSid)."

                    if ($null -ne $foundTrusteeGroup) { $trusteeGroupExpansionComplete = $true }
                } # end $expandGroups -eq $true

                # We aren't expanding groups, so just search and return trustee group (if found).
                $false {

                    switch ($getTrusteeUseSid) {
                        # one last time...

                        $true {
                            $foundTrustee = Invoke-Command @icCommon -ScriptBlock {

                                Get-Group -Filter "Sid -eq '$($Using:trusteeSid)' -or SidHistory -eq '$($Using:trusteeSid)'" -ErrorAction:SilentlyContinue |
                                Select-Object $Using:trusteeProps
                            }
                        }

                        $false {
                            $foundTrustee = Invoke-Command @icCommon -ScriptBlock {

                                Get-Group -Filter "Guid -eq '$($Using:trusteeSid)'" -ErrorAction:SilentlyContinue |
                                Select-Object $Using:trusteeProps
                            }
                        }
                    } # end switch ($getTrusteeUseSid) (#3)
                } # end $expandGroups -eq $false
            } # end switch ($expandGroups)
        } # end if ($null -eq $foundTrustee) {}

        # If trustee was not found (or returned multiple matches), make note of this in the TrusteeGuid property (because TrusteeGuid is output even with -MinimizeOutput)
        if (($null -eq $foundTrustee) -or ($foundTrustee.Count -gt 1)) {

            # But first, if it was a group that was expanded, report so in the TrusteeType property, and return the trustee group object.
            if ($trusteeGroupExpansionComplete -eq $true) {

                $trusteeReturned |
                Add-Member -NotePropertyName 'TrusteeGuid' -NotePropertyValue $foundTrusteeGroup.Guid

                if (-not ($MinimizeOutput)) {

                    $trusteeReturned |
                    Add-Member -NotePropertyName 'TrusteeType' -NotePropertyValue "EXPANDED:$($foundTrusteeGroup.RecipientTypeDetails.Value)" -PassThru |
                    Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue $foundTrusteeGroup.DisplayName -PassThru |
                    Add-Member -NotePropertyName 'TrusteePSmtp' -NotePropertyValue $foundTrusteeGroup.WindowsEmailAddress
                }
            }

            else {
                $trusteeReturned |
                Add-Member -NotePropertyName 'TrusteeGuid' -NotePropertyValue "Not found or ambiguous ($($trusteeSid))"

                if (-not ($MinimizeOutput)) {

                    $trusteeReturned |
                    Add-Member -NotePropertyName 'TrusteeType' -NotePropertyValue '' -PassThru |
                    Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue '' -PassThru |
                    Add-Member -NotePropertyName 'TrusteePSmtp' -NotePropertyValue ''
                }
            }
        }

        # Otherwise add the successfully found trustee's pertinent properties.
        else {
            $trusteeReturned |
            Add-Member -NotePropertyName 'TrusteeGuid' -NotePropertyValue $foundTrustee.Guid

            if (-not ($MinimizeOutput)) {

                $trusteeReturned |
                Add-Member -NotePropertyName 'TrusteeType' -NotePropertyValue $foundTrustee.RecipientTypeDetails -PassThru |
                Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue $foundTrustee.DisplayName -PassThru |
                Add-Member -NotePropertyName 'TrusteePSmtp' -NotePropertyValue $foundTrustee.WindowsEmailAddress
            }
        }

        # Finally, return the finished product.
        Write-Output $trusteeReturned |
        Select-Object -Property * -ExcludeProperty GrantSendOnBehalfTo, placeholder, PS*ComputerName, RunspaceId

    } # end function getTrustee


    # Fire up the engines.  Brap brap...
    $MainProgress = @{
        Activity         = "Get-MailboxTrustee.ps1 - Start time: $($StartTime.DateTime)"
        Status           = "Working in $($Exchange) ($($ExPSSession.ComputerName)) [ExpandTrusteeGroups:`$$($ExpandTrusteeGroups)]"
        Id               = 0
        ParentId         = -1
        SecondsRemaining = -1
    }
    Write-Progress @MainProgress
    Start-Sleep -Milliseconds 500

    $faProgressProps = @{
        Activity         = 'FullAccess'
        Id               = 1
        ParentId         = 0
        SecondsRemaining = -1
    }
    Write-Progress @faProgressProps -Status 'Ready'
    Start-Sleep -Milliseconds 500

    $saProgressProps = @{
        Activity         = 'Send-As'
        Id               = 2
        ParentId         = 0
        SecondsRemaining = -1
    }
    Write-Progress @saProgressProps -Status 'Ready'
    Start-Sleep -Milliseconds 500

    $sobProgressProps = @{
        Activity         = 'Send on Behalf'
        Id               = 3
        ParentId         = 0
        SecondsRemaining = -1
    }
    Write-Progress @sobProgressProps -Status 'Ready'
    Start-Sleep -Milliseconds 500

    $fpProgressProps = @{
        Activity         = 'Folder Permissions'
        Id               = 4
        ParentId         = 0
        SecondsRemaining = -1
    }
    Write-Progress @fpProgressProps -Status 'Ready'
    Start-Sleep -Milliseconds 500

}

process {

    # Placing all of the process block inside this single try block.  The purpose is to kill the script when something breaking occurs.  Use Verbose/Debug for troubleshooting.
    try {


        #-----------------------------------#
        #----- Initial Mailbox Lookoup -----#
        #-----------------------------------#

        Write-Verbose -Message "Getting mailbox with identity '$($Identity)'."

        $Mailbox = $null
        $Mailbox = Invoke-Command @icCommon -ScriptBlock {

            Get-Mailbox -Identity $Using:Identity -ErrorAction:Stop |
            Select-Object -Property DisplayName,
            PrimarySmtpAddress,
            RecipientTypeDetails,
            Guid,
            GrantSendOnBehalfTo
        }

        # Store the mailbox' PrimarySmtpAddress for use with Write-Progress|Verbose|Debug.
        # Store the mailbox' Guid for use with -Identiy parameters.
        [string]$mPSmtp = $Mailbox.PrimarySmtpAddress
        [string]$mGuid = $Mailbox.Guid

        Write-Verbose -Message "[Mailbox: $($mPSmtp)] Mailbox lookup complete."

        if ($MailboxPercentCompletePossible) {

            $MailboxCounter++
            Write-Progress @MainProgress -CurrentOperation "Mailbox #$MailboxCounter of $PipelineObjectCount`: $($Mailbox.DisplayName) ($($mPSmtp))" -PercentComplete (($MailboxCounter / $PipelineObjectCount) * 100)
        }
        else {
            Write-Progress @MainProgress -CurrentOperation "Current mailbox: $($Mailbox.DisplayName) ($($mPSmtp))"
        }


        #-----------------------------------#
        #----------- FullAccess ------------#
        #-----------------------------------#

        if ($BypassPermissionTypes -notcontains 'FullAccess') {

            $faProgressCounter = 0
            Write-Progress @faProgressProps -Status 'Getting FullAccess permissions with Get-MailboxPermission'
            Write-Verbose -Message "[Mailbox: $($mPSmtp)] Getting FullAccess permissions with Get-MailboxPermission."

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

                    # Initialize counter for FullAccess permissions to process.
                    $faIndexCounter = 0

                    $FullAccess |
                    Select-Object -Index (0..999) |
                    ForEach-Object {

                        $faProgressCounter++
                        Write-Progress @faProgressProps -PercentComplete (($faProgressCounter / $FullAccess.Count) * 100) -CurrentOperation "Trustee: $($_.User)" -Status 'Getting trustees with Get-User/Get-Group'

                        $getTrusteeProps = @{
                            trusteeSid    = $faTrustees[$($faIndexCounter)].SecurityIdentifier.Value
                            mailboxObject = $Mailbox
                            accessRights  = 'FullAccess'
                            expandGroup   = $ExpandTrusteeGroups
                        }
                        getTrustee @getTrusteeProps

                        $faIndexCounter++
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

                    $FullAccess |
                    Select-Object -Index (0..999) |
                    ForEach-Object {

                        $faProgressCounter++
                        Write-Progress @faProgressProps -PercentComplete (($faProgressCounter / $FullAccess.Count) * 100) -CurrentOperation "Trustee: $($_.RawIdentity)" -Status 'Getting trustees with Get-User/Get-Group'

                        $getTrusteeProps = @{
                            trusteeSid    = $_.SecurityIdentifier
                            mailboxObject = $Mailbox
                            accessRights  = 'FullAccess'
                            expandGroup   = $ExpandTrusteeGroups
                        }
                        getTrustee @getTrusteeProps

                        $faIndexCounter++
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

            switch ($Exchange) {

                "Exchange Online" {

                    $saProgressCounter = 0
                    Write-Progress @saProgressProps -Status 'Getting Send-As permissions with Get-RecipientPermission'
                    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Getting Send-As permissions with Get-RecipientPermission."

                    $RecipientPermissions = @()
                    $RecipientPermissions += Invoke-Command @icCommon -ScriptBlock {

                        Get-RecipientPermission -Identity $Using:mGuid -ErrorAction:SilentlyContinue |
                        Select-Object Trustee, AccessRights, IsInherited, Deny
                    }

                    # Apply filters.
                    $SendAs = @()
                    $SendAs += $RecipientPermissions |
                    Where-Object {
                        -not
                        ($_.Trustee -like 'S-1-5*') -and -not
                        ($_.Trustee -like 'NT AUTHORITY\*') -and
                        ($_.AccessRights -like '*SendAs*') -and
                        ($_.IsInherited -eq $false) -and
                        ($_.Deny -ne $true) # <--: Comes back empty instead of $false when remoting.
                    }

                    if ($SendAs.Count -ge 1) {
                        Write-Verbose -Message "[Mailbox: $($mPSmtp)] Processing Send-As trustees."
                    }

                    $SendAs |
                    Select-Object -Index (0..999) |
                    ForEach-Object {

                        $saProgressCounter++
                        Write-Progress @saProgressProps -PercentComplete (($saProgressCounter / $SendAs.Count) * 100) -CurrentOperation "Trustee: $($_.Trustee)" -Status 'Getting trustees with Get-User/Get-Group'

                        # We will output our EXO Send-As trustee here since we can't get a Sid or a Guid from Get-RecipientPermission's Trustee property, hence getTrustee won't work.
                        $saTrusteeId = $null
                        $saTrusteeId = $_.Trustee

                        $saFoundTrustee = @()
                        $saFoundTrustee += Invoke-Command @icCommon -ScriptBlock {

                            Get-User -Identity $Using:saTrusteeId -ErrorAction:SilentlyContinue
                        }

                        if ($null -eq $saFoundTrustee) {

                            $saFoundTrustee += Invoke-Command @icCommon -ScriptBlock {

                                Get-Group -Identity "$($Using:saTrusteeId)" -ErrorAction:SilentlyContinue
                            }
                        }

                        # Initialize output object.
                        $saTrusteeReturned = $null
                        $saTrusteeReturned = [PSCustomObject]@{placeholder = $null }

                        # Load the current mailbox' properties into our custom object.
                        switch ($MinimizeOutput) {

                            $true {
                                $saTrusteeReturned |
                                Add-Member -NotePropertyName 'Guid' -NotePropertyValue $Mailbox.Guid
                            }
                            $false {
                                $Mailbox |
                                Get-Member -MemberType Properties |
                                ForEach-Object {

                                    $saTrusteeReturned |
                                    Add-Member -NotePropertyName $_.Name -NotePropertyValue $Mailbox.$($_.Name)
                                }
                            }
                        }

                        # Add our PermissionType and AccessRights properties.
                        $saTrusteeReturned |
                        Add-Member -NotePropertyName 'PermissionType' -NotePropertyValue 'Mailbox' -PassThru |
                        Add-Member -NotePropertyName 'AccessRights' -NotePropertyValue 'Send-As'


                        # If trustee was not found (or returned multiple matches), make note of this in the TrusteeGuid property (since it is output even when -MinimizeOutput is used).
                        if ($safoundTrustee.Count -ne 1) {

                            $saTrusteeReturned |
                            Add-Member -NotePropertyName 'TrusteeGuid' -NotePropertyValue "Not found or ambiguous ($($saTrusteeId))"

                            if (-not ($MinimizeOutput)) {

                                $saTrusteeReturned |
                                Add-Member -NotePropertyName 'TrusteeType' -NotePropertyValue '' -PassThru |
                                Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue '' -PassThru |
                                Add-Member -NotePropertyName 'TrusteePSmtp' -NotePropertyValue ''
                            }

                            Write-Output $saTrusteeReturned |
                            Select-Object -Property * -ExcludeProperty GrantSendOnBehalfTo, placeholder, PS*ComputerName, RunspaceId
                        }

                        # If trustee was found, is a group, and we're expanding groups... send its SID up to getTrustee for processing.
                        elseif (($saFoundTrustee.RecipientTypeDetails -like '*group*') -and ($ExpandTrusteeGroups -eq $true)) {

                            $getTrusteeProps = @{
                                trusteeSid    = $saFoundTrustee.Sid
                                mailboxObject = $Mailbox
                                accessRights  = 'SendAs'
                                expandGroup   = $true
                            }
                            getTrustee @getTrusteeProps
                        }

                        # If our logic is correct, we've got a fully loaded trustee object to return.
                        else {
                            $saTrusteeReturned |
                            Add-Member -NotePropertyName 'TrusteeGuid' -NotePropertyValue $saFoundTrustee.Guid

                            if (-not ($MinimizeOutput)) {

                                $saTrusteeReturned |
                                Add-Member -NotePropertyName 'TrusteeType' -NotePropertyValue $saFoundTrustee.RecipientTypeDetails -PassThru |
                                Add-Member -NotePropertyName 'TrusteeDisplayName' -NotePropertyValue $saFoundTrustee.DisplayName -PassThru |
                                Add-Member -NotePropertyName 'TrusteePSmtp' -NotePropertyValue $saFoundTrustee.WindowsEmailAddress
                            }

                            Write-Output $saTrusteeReturned |
                            Select-Object -Property * -ExcludeProperty GrantSendOnBehalfTo, placeholder, PS*ComputerName, RunspaceId
                        }
                        $saIndexCounter++

                    } # end SendAs | ForEach-Object {}
                } # end $Exchange -eq 'Exchange Online'

                "Exchange On-Premises" {

                    $saProgressCounter = 0
                    Write-Progress @saProgressProps -Status 'Getting Send-As permissions with Get-ADPermission'
                    Write-Verbose -Message "[Mailbox: $($mPSmtp)] Getting Send-As permissions with Get-ADPermission."

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

                            # Initialize counter for FullAccess permissions to process.
                            $saIndexCounter = 0

                            $SendAs |
                            Select-Object -Index (0..999) |
                            ForEach-Object {

                                $saProgressCounter++
                                Write-Progress @saProgressProps -PercentComplete (($saProgressCounter / $SendAs.Count) * 100) -CurrentOperation "Trustee: $($_.User)" -Status 'Getting trustees with Get-User/Get-Group'

                                $getTrusteeProps = @{
                                    trusteeSid    = $saTrustees[$($saIndexCounter)].SecurityIdentifier.Value
                                    mailboxObject = $Mailbox
                                    accessRights  = 'SendAs'
                                    expandGroup   = $ExpandTrusteeGroups
                                }
                                getTrustee @getTrusteeProps

                                $saIndexCounter++
                            }
                        } # end $Legacy -eq $true

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

                            $SendAs |
                            Select-Object -Index (0..999) |
                            ForEach-Object {

                                $saProgressCounter++
                                Write-Progress @saProgressProps -PercentComplete (($saProgressCounter / $SendAs.Count) * 100) -CurrentOperation "Trustee: $($_.RawIdentity)" -Status 'Getting trustees with Get-User/Get-Group'

                                $getTrusteeProps = @{
                                    trusteeSid    = $_.SecurityIdentifier
                                    mailboxObject = $Mailbox
                                    accessRights  = 'SendAs'
                                    expandGroup   = $ExpandTrusteeGroups
                                }
                                getTrustee @getTrusteeProps

                                $saIndexCounter++
                            }
                        } # end $LegacyExchange -eq $false
                    } # end switch ($LegacyExchange)
                } # end $Exchange -eq 'Exchange On-Premises'
            } # end switch ($Exchange)

            Write-Progress @saProgressProps -Status 'Ready'
            Write-Verbose -Message "[Mailbox: $($mPSmtp)] Send-As discovery complete."

        } # end if ($BypassPermissionTypes -notcontains 'SendAs') {}


        #-----------------------------------#
        #---------- Send on Behalf ---------#
        #-----------------------------------#

        if ($BypassPermissionTypes -notcontains 'SendOnBehalf') {

            $sobProgressCounter = 0
            Write-Progress @sobProgressProps -Status 'Checking for Send on Behalf trustees in GrantSendOnBehalfTo property'

            Write-Verbose -Message "[Mailbox: $($mPSmtp)] Checking for Send on Behalf trustees in GrantSendOnBehalfTo property."

            if ($Mailbox.GrantSendOnBehalfTo.Count -ge 1) {

                Write-Verbose -Message "[Mailbox: $($mPSmtp)] Processing Send on Behalf trustees."

                $sobTrustees = Invoke-Command @icCommon -ScriptBlock {

                    Get-Mailbox -Identity "$($Using:mGuid)" -ErrorAction:SilentlyContinue |
                    Select-Object -ExpandProperty GrantSendOnBehalfTo
                }

                $sobTrustees |
                Select-Object -Index (0..999) |
                ForEach-Object {

                    $sobProgressCounter++
                    Write-Progress @sobProgressProps -PercentComplete (($sobProgressCounter / $Mailbox.GrantSendOnBehalfTo.Count) * 100) -CurrentOperation "Trustee: $($_)" -Status  'Getting trustees with Get-User/Get-Group'

                    if (-not ([string]::IsNullOrEmpty($_.ObjectGuid.Guid))) {

                        getTrustee -trusteeSid $_.ObjectGuid.Guid -mailboxObject $Mailbox -accessRights 'SendOnBehalf' -expandGroups $ExpandTrusteeGroups
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

            $Folders |
            ForEach-Object {

                # Store the current folder name for use throughout the pipeline.
                [string]$Folder = "$($_)"

                $fpProgressCounter = 0

                Write-Progress @fpProgressProps -Status "Getting $($Folder) permissions with Get-MailboxFolderPermission"
                Write-Verbose -Message "[Mailbox: $($mPSmtp)][Folder: $($Folder)] Getting $($Folder) permissions with Get-MailboxFolderPermission."

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

                        # Initialize counter for pertinent folder permissions to process.
                        $pfpIndexCounter = 0

                        $PertinentFolderPermissions |
                        Select-Object -Index (0..999) |
                        ForEach-Object {

                            $fpProgressCounter++
                            Write-Progress @fpProgressProps -PercentComplete (($fpProgressCounter / $PertinentFolderPermissions.Count) * 100) -CurrentOperation "Trustee: $($_.User)" -Status "Getting $($Folder) trustees with Get-User/Get-Group"

                            if (-not ([string]::IsNullOrEmpty("$($fpTrustees[$($pfpIndexCounter)].ADRecipient.Sid)"))) {

                                $getTrusteeProps = @{
                                    trusteeSid    = $fpTrustees[$($pfpIndexCounter)].ADRecipient.Sid
                                    mailboxObject = $Mailbox
                                    accessRights  = $_.AccessRights -join ';'
                                    folder        = $true
                                    folderName    = $Folder
                                    expandGroup   = $ExpandTrusteeGroups
                                }

                                getTrustee @getTrusteeProps
                            }
                            $pfpIndexCounter++

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

                        $PertinentFolderPermissions |
                        Select-Object -Index (0..999) |
                        ForEach-Object {

                            $fpProgressCounter++
                            Write-Progress @fpProgressProps -PercentComplete (($fpProgressCounter / $PertinentFolderPermissions.Count) * 100) -CurrentOperation "Trustee: $($_.User)" -Status "Getting $($Folder) trustees with Get-User/Get-Group"

                            if (-not ([string]::IsNullOrEmpty("$($_.ADRecipient.Sid)"))) {

                                $getTrusteeProps = @{
                                    trusteeSid    = $_.ADRecipient.Sid
                                    mailboxObject = $Mailbox
                                    accessRights  = $_.AccessRights -join ';'
                                    folder        = $true
                                    folderName    = $Folder
                                    expandGroup   = $ExpandTrusteeGroups
                                }

                                getTrustee @getTrusteeProps
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

        # Close try block.
    }

    # Session problems go here:
    catch [System.Management.Automation.CommandNotFoundException],
    [System.Management.Automation.Remoting.PSRemotingTransportException] {

        Write-Warning -Message 'Session problems (most likely) have caused the script to fail.'
        Write-Warning -Message 'Mailboxes processed/total (if available): $MailboxCounter / $PipelineObjectCount'
        Write-Warning -Message 'Consider using a helper script to test and recreate the session every X number of processed mailboxes.'
        Write-Warning -Message 'Also try using the -Verbose or -Debug switches.'
        Write-Error $Error[0]
        Write-Warning -Message "Date/time: $(Get-Date -Format G)"
        Write-Warning -Message 'Sleeping for 60 seconds before moving onto the next mailbox, if any remain.'
        Start-Sleep -Seconds 60
    }

    # Other problems go here:
    catch {

        Write-Warning -Message 'A problem has caused the script to fail.'
        Write-Warning -Message 'Mailboxes processed/total (if available): $MailboxCounter / $PipelineObjectCount'
        Write-Warning -Message 'Try using -Verbose or -Debug switches.'
        Write-Error $Error[0]
        Write-Warning -Message "Date/time: $(Get-Date -Format G)"
        Write-Warning -Message 'Sleeping for 60 seconds before moving onto the next mailbox, if any remain.'
        Start-Sleep -Seconds 60
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

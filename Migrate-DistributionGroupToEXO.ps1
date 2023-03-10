<#
    .Synopsis
    Migrate (non-security, simple) distribution groups from Exchange on-premises to Exchange Online.

    .Description
    This script offers several modes mapping to the steps involved in migrating a distribution group from on-premises
    to Exchange Online.
      - Security-enabled groups are not supported, only true Distribution groups are.
      - Groups which are nested as a member in one or more other groups are not supported.
      - Group which have one or more nested groups as members are not supported.

    .Parameter BackupFromOnPremises
    Switch parameter to select the migration step of backing up the group from on-premises.  This should be used first
    to get a backup of the group's important properties and members list.

    .Parameter GlobalCatalogDomainControllerFQDN
    Specify the forest, domain, or domain controller FQDN which will respond to ActiveDirectory PowerShell module
    commands on port 3268.  Required in backup mode.

    .Parameter Identity
    Specifies the group to be backed up to XML and the supplied value is used for commands: Get-DistributionGroup,
    Get-DistributionGroupMember.  Required in backup mode.

    .Parameter FallbackManagedByPSMTP
    Specifies the PrimarySmtpAddress of a valid recipient to set as the group owner (ManagedBy), in the event none of
    the group owners in the backup can be found in EXO.  If not specified and no owners could be found, the account
    used to run the script will automatically be set as the owner.

    .Parameter DistributionGroupBackupFolderPath
    Specifies the output folder to which a backup XML file will be saved as: "DGBackup_<group's PrimarySmtpAddress>_<date and time>.xml"
    Required in backup mode.

    .Parameter DistributionGroupBackupFilePath
    Specifies the file path of a previously backed-up group's XML file.  Required in all modes except backup mode.

    .Parameter BACKOUTRecreateInOnPremises
    Switch parameter to select the migration step of backing out and recreating the group on-premises.  This mode of
    the script is not yet implemented, but the backup XML file created in the backup mode contains the necessary info
    to be able to manually recreate the group as it was.

    .Parameter RecreateInEXO
    Switch parameter to select the migration step of creating the group in EXO.

    .Parameter PlaceholderOnly
    Switch parameter intended for use with the RecreateInEXO mode.  Indicates to create the group in EXO but to forego
    the step of adding the email addresses, and prepends the suffix "zzzTmpDLMigration_" to the Name, DisplayName,
    Alias, and PrimarySmtpAddress properties.

    .Parameter UpdateEXOPlaceholder
    Switch parameter to select the migration step of finalizing the previously-created placeholder group in EXO, by
    removing the "zzzTmpDLMigration_" suffix and applying the original email addresses (including X500's (including 
    X500:<legacyExchangeDN>)).
#>
#Requires -Version 4.0
[CmdletBinding(
    SupportsShouldProcess = $true,
    ConfirmImpact = 'High'
)]
param (
    [Parameter(ParameterSetName = 'BackupFromOnPremises')]
    [switch]$BackupFromOnPremises,

    [Parameter(ParameterSetName = 'BACKOUTRecreateInOnPremises')]
    [switch]$BACKOUTRecreateInOnPremises,

    [Parameter(ParameterSetName = 'RecreateInEXO')]
    [switch]$RecreateInEXO,

    [Parameter(ParameterSetName = 'UpdatePlaceholderInEXO')]
    [switch]$UpdateEXOPlaceholder,

    [Parameter(Mandatory, ParameterSetName = 'BackupFromOnPremises')]
    [Parameter(Mandatory, ParameterSetName = 'BACKOUTRecreateInOnPremises')]
    [string]$GlobalCatalogDomainControllerFQDN,

    [Parameter(Mandatory, ParameterSetName = 'BackupFromOnPremises')]
    [string]$Identity,

    [Parameter(ParameterSetName = 'BackupFromOnPremises')]
    [ValidateScript(
        {
            if (Test-Path -Path $_ -PathType Container) { $true }
            else { throw "Failed to validate folder path: $($_)" }
        }
    )]
    [System.IO.FileInfo]$DistributionGroupBackupFolderPath = $ENV:HOMEPATH,
    
    [Parameter(ParameterSetName = 'BackupFromOnPremises')]
    [string]$FallbackManagedByPSMTP,

    [Parameter(Mandatory, ParameterSetName = 'BACKOUTRecreateInOnPremises')]
    [Parameter(Mandatory, ParameterSetName = 'RecreateInEXO')]
    [Parameter(Mandatory, ParameterSetName = 'UpdatePlaceholderInEXO')]
    [ValidateScript(
        {
            if (Test-Path -Path $_ -PathType Leaf) { $true }
            else { throw "Failed to validate file path: $($_)" }
        }
    )]
    [System.IO.FileInfo]$DistributionGroupBackupFilePath,

    [Parameter(ParameterSetName = 'RecreateInEXO')]
    [switch]$PlaceholderOnly
)

#======#-----------#
#region# Functions #
#======#-----------#

$Script:_rcptTracker = @{}
function getRecipients ([Object[]]$List) {

    foreach ($_rcpt in $List) {
        # Keep track of found recipients to avoid re-running Get-Recipient unnecessarily:
        if (-not ($Script:_rcptTracker[$_rcpt])) {
            $_currentRcpt = $null
            $_currentRcpt = Get-Recipient -Identity $_rcpt -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PrimarySmtpAddress
            if ($null -ne $_currentRcpt) {
                $Script:_rcptTracker[$_rcpt] = $_currentRcpt
                $_currentRcpt
            }
        }
        else { $Script:_rcptTracker[$_rcpt] }
    }
}

#=========#-----------#
#endregion# Functions #
#=========#-----------#



#======#----------------#
#region# Initialization #
#======#----------------#

# Verify required commands are available:
$_requiredCmdlets = if ($BackupFromOnPremises) {
    @(
        'Get-ADGroup', 'Get-ADGroupMember',
        'Set-ADServerSettings', 'Get-DistributionGroup', 'Get-DistributionGroupMember', 'Get-ADPermission',
        'Export-Clixml'
    )
}
elseif ($BACKOUTRecreateInOnPremises) {
    @(
        'Import-Clixml', 'Set-ADServerSettings',
        'Get-DistributionGroup', 'Get-DistributionGroupMember', 'New-DistributionGroup', 'Set-DistributionGroup', #'Add-DistributionGroupMember'
        'Add-ADPermission'
    )
}
elseif ($RecreateInEXO) {
    @(
        'Import-Clixml', 'Get-ConnectionInformation',
        'Get-DistributionGroup', 'Get-DistributionGroupMember', 'New-DistributionGroup', 'Set-DistributionGroup', #'Add-DistributionGroupMember'
        'Add-RecipientPermission'
    )
}
elseif ($UpdateEXOPlaceholder) {
    @(
        'Import-Clixml', 'Get-ConnectionInformation', 'Get-DistributionGroup', 'Set-DistributionGroup'
    )
}

$_missingCmdlets = @()

foreach ($_cmdlet in $_requiredCmdlets) {

    if (-not (Get-Command $_cmdlet -ErrorAction SilentlyContinue)) { $_missingCmdlets += $_cmdlet }
}
if ($_missingCmdlets.Count -ge 1) {

    throw "Missing cmdlets: $($_missingCmdlets -join ', ').  Required cmdlets for this migration step: $($_requiredCmdlets -join ', ')."
}

# Verify (if applicable) file being imported:
if ($BACKOUTRecreateInOnPremises -or $RecreateInEXO -or $UpdateEXOPlaceholder) {

    $DistributionGroup = Import-Clixml $DistributionGroupBackupFilePath -ErrorAction Stop
    $_requiredProperties = @(
        # Informational Properties:
        'Name', 'DisplayName', 'PrimarySmtpAddress', 'EmailAddresses', 'Alias', 'SamAccountName', 'OrganizationalUnit',
        'CustomAttribute1', 'CustomAttribute2', 'CustomAttribute3', 'CustomAttribute4', 'CustomAttribute5',
        'CustomAttribute6', 'CustomAttribute7', 'CustomAttribute8', 'CustomAttribute9', 'CustomAttribute10',
        'CustomAttribute11', 'CustomAttribute12', 'CustomAttribute13', 'CustomAttribute14', 'CustomAttribute15',
        'Guid', 'LegacyExchangeDN',

        # Settings:
        'BypassNestedModerationEnabled',
        'HiddenFromAddressListsEnabled',
        'MailTip', 'MemberDepartRestriction', 'MemberJoinRestriction', 'ModerationEnabled',        
        'ReportToManagerEnabled', 'ReportToOriginatorEnabled', 'RequireSenderAuthenticationEnabled',
        'SendModerationNotifications', 'SendOofMessageToOriginatorEnabled',

        # Recipient Lists:
        'AcceptMessagesOnlyFromSendersOrMembers',
        'BypassModerationFromSendersOrMembers',
        'GrantSendOnBehalfTo',
        'ManagedBy', 'Members', 'ModeratedBy',
        'RejectMessagesFromSendersOrMembers',
        'SendAs'
    )
    $_includedProperties = $DistributionGroup | Get-Member -MemberType NoteProperty
    $_missingProperties = @()
    foreach ($_property in $_requiredProperties) {
        if ($_includedProperties.Name -notcontains $_property) { $_missingProperties += $_property }
    }
    if ($_missingProperties.Count -ge 1) {
        throw "Backup file is missing properties: $($_missingProperties -join ', ').  Required properties for migration step -RecreateInEXO: $($_requiredProperties -join ', ')."
    }
}

# Verify (if applicable) EXO module/connection:
if ($RecreateInEXO -or $UpdateEXOPlaceholder) {

    # Verify EXO module is loaded and current enough (v3.0.0 minimum):
    $EXOModule = Get-Module ExchangeOnlineManagement
    if (-not $EXOModule) {
        throw 'This script requires that the ExchangeOnlineManagement module (v3.0.0 or newer) be installed, imported, and connected (Connect-ExchangeOnline).'
    }
    elseif ([int]$EXOModule.Version.Major -lt 3) {
        throw 'This script requires v3.0.0 or newer for the ExchangeOnlineManagement module.'
    }

    if (-not (Get-ConnectionInformation)) {
        throw 'This script requires Connect-ExchangeOnline to have been completed in advance.'    
    }
}

#=========#----------------#
#endregion# Initialization #
#=========#----------------#



#======#-------------------------#
#region# Backup from On-Premises #
#======#-------------------------#
if ($BackupFromOnPremises) {
    try {
        # Find and inspect the group for suitability:
        Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
        $DistributionGroup = Get-DistributionGroup -Identity $Identity -ErrorAction Stop
        $DistributionGroupMembers = Get-DistributionGroupMember -Identity $Identity -ErrorAction Stop -ResultSize Unlimited
        $ADGroup = Get-ADGroup -Identity $DistributionGroup.Guid.Guid -Properties Member, MemberOf -Server "$($GlobalCatalogDomainControllerFQDN):3268" -ErrorAction Stop
        $SendAs = Get-ADPermission -Identity $DistributionGroup.Guid.Guid -ErrorAction Stop | Where-Object { ($_.IsInherited -ne $true) -and ($_.ExtendedRights -like '*Send-As*') }

        if ($ADGroup.GroupCategory -eq 'Security') {
            throw "This script currently intends to only process non-security enabled distribution groups.  Group with identity '$($Identity)' was found to be security-enabled, and could be used for access control to resources."
        }
        elseif ($ADGroup.MemberOf.Count -ge 1) {
            throw "This script currently intends to only process groups that aren't nested as a member in another group.  Group with identity '$($Identity)' was found to be a member of one or more other groups."
        }
        elseif ($DistributionGroupMembers | Where-Object { $_.RecipientType -like '*Group*' }) {
            throw "This script currently intends to only process groups that don't have other groups nested as members.  Group with identity '$($Identity)' was found to have one or more members that is a group."
        }

        # Backup pertinent information:
    
        # Single-value properties:
        $DGBackup = @{

            Name                               = $DistributionGroup.Name
            DisplayName                        = $DistributionGroup.DisplayName
            PrimarySmtpAddress                 = $DistributionGroup.PrimarySmtpAddress
            Alias                              = $DistributionGroup.Alias
            MailTip                            = $DistributionGroup.MailTip
            SamAccountName                     = $DistributionGroup.SamAccountName
            ModerationEnabled                  = $DistributionGroup.ModerationEnabled
            MemberJoinRestriction              = $DistributionGroup.MemberJoinRestriction.ToString()
            MemberDepartRestriction            = $DistributionGroup.MemberDepartRestriction.ToString()
            RequireSenderAuthenticationEnabled = $DistributionGroup.RequireSenderAuthenticationEnabled
            HiddenFromAddressListsEnabled      = $DistributionGroup.HiddenFromAddressListsEnabled
            BypassNestedModerationEnabled      = $DistributionGroup.BypassNestedModerationEnabled
            ReportToManagerEnabled             = $DistributionGroup.ReportToManagerEnabled
            ReportToOriginatorEnabled          = $DistributionGroup.ReportToOriginatorEnabled
            SendOofMessageToOriginatorEnabled  = $DistributionGroup.SendOofMessageToOriginatorEnabled
            SendModerationNotifications        = $DistributionGroup.SendModerationNotifications.ToString()
            OrganizationalUnit                 = $DistributionGroup.OrganizationalUnit
            Guid                               = $DistributionGroup.Guid
            LegacyExchangeDN                   = $DistributionGroup.LegacyExchangeDN
        }
        # CustomAttributes:
        foreach ($_ca in 1..15) {
            $_thisCA = "CustomAttribute$($_ca)"
            $DGBackup[$_thisCA] = if ($DistributionGroup.$_thisCA) { $DistributionGroup.$_thisCA } else { $null }
        }

        # Relevant email addresses:
        $DGBackup['EmailAddresses'] = [string[]]@($DistributionGroup.EmailAddresses |
            where-Object { $_ -match '(smtp)|(x500)' }) + "X500:$($DistributionGroup.LegacyExchangeDN)"

        # Group Owners:
        $DGBackup['ManagedBy'] = if ($DistributionGroup.ManagedBy) { [string[]]@(getRecipients -List $DistributionGroup.ManagedBy) }
        elseif ($FallbackManagedByPSMTP) { $FallbackManagedByPSMTP }
        else { $null }

        # Sender/Delivery Restrictions and Moderators:
        foreach ($_property in 'AcceptMessagesOnlyFromSendersOrMembers',
            'RejectMessagesFromSendersOrMembers',
            'GrantSendOnBehalfTo',
            'BypassModerationFromSendersOrMembers',
            'ModeratedBy') {

            $DGBackup["$($_property)"] = if ($DistributionGroup.$_property) { [string[]]@(getRecipients -List $DistributionGroup.$_property) }
            else { $null }
        }

        # Send-As:
        $DGBackup['SendAs'] = if ($SendAs) { [string[]]@(getRecipients -List $SendAs.User) }
        else { $null }

        # Members:
        $DGBackup['Members'] = if ($DistributionGroupMembers) { $DistributionGroupMembers | ForEach-Object { $_.PrimarySmtpAddress.ToString() } }
        else { $null ; Write-Warning -Message "No recipient-object members (i.e., mail-enabled) found in group '$($Identity)'." }

        # Export/backup:
        $_now = [datetime]::Now; [PSCustomObject]$DGBackup |
        Export-Clixml "$($DistributionGroupBackupFolderPath)\DGBackup_$($DistributionGroup.PrimarySmtpAddress)_$($_now.ToString('yyyy-MM-dd_HH-mm-ss_z')).xml" -Depth 10 -ErrorAction Stop
    }
    catch { throw }

    if (-not $WhatIfPreference.IsPresent) {

        # Advise that group is ready to be un-synced:
        "Group '$($Identity)' can now be un-synced from Azure AD.  This can be done however desired (e.g., " +
        "set adminDescription to 'Group_DoNotSync', move to an un-synced OU, Disable-DistributionGroup, Remove-DistributionGroup, etc.).  " +
        "This must be done before re-creating the group in Exchange Online.  Alternatively, using the -RecreateInEXO -PlaceHolderOnly switches, " +
        "the group can be created as a placeholder in EXO before un-syncing the group." | Write-Host -ForegroundColor Green
    }
}
#=========#-------------------------#
#endregion# Backup from On-Premises #
#=========#-------------------------#



#======#---------------------------------#
#region# BACKOUT Recreate in On-Premises #
#======#---------------------------------#
if ($BACKOUTRecreateInOnPremises) { <# Coming soon... someday :) #> }
#=========#---------------------------------#
#endregion# BACKOUT Recreate in On-Premises #
#=========#---------------------------------#



#======#-----------------#
#region# Recreate in EXO #
#======#-----------------#
if ($RecreateInEXO) {
    try {
        # Including as many properties as allowed with New-DistributionGroup:
        $_newDGParams = @{
            
            Name                               = $DistributionGroup.DisplayName
            DisplayName                        = $DistributionGroup.DisplayName
            PrimarySmtpAddress                 = $DistributionGroup.PrimarySmtpAddress
            Alias                              = $DistributionGroup.Alias
            ModerationEnabled                  = $DistributionGroup.ModerationEnabled
            MemberJoinRestriction              = $DistributionGroup.MemberJoinRestriction
            MemberDepartRestriction            = $DistributionGroup.MemberDepartRestriction
            RequireSenderAuthenticationEnabled = $DistributionGroup.RequireSenderAuthenticationEnabled
            BypassNestedModerationEnabled      = $DistributionGroup.BypassNestedModerationEnabled
            SendModerationNotifications        = $DistributionGroup.SendModerationNotifications
        }

        # Determine available recipient objects for applicable properties:
        $_availableRecipients = @{}
        $_missingRecipients = @{}
        foreach ($_rcptListProperty in 'ManagedBy', 'ModeratedBy',
            'AcceptMessagesOnlyFromSendersOrMembers', 'RejectMessagesFromSendersOrMembers', 'BypassModerationFromSendersOrMembers',
            'GrantSendOnBehalfTo', 'SendAs',
            'Members') {

            $_availableRecipients[$_rcptListProperty] = if ($DistributionGroup.$_rcptListProperty) {
                [string[]]@(getRecipients -List $DistributionGroup.$_rcptListProperty)
            } else { $null }

            $_missingRecipients[$_rcptListProperty] = if ($_availableRecipients[$_rcptListProperty]) {
                $DistributionGroup.$($_rcptListProperty) | Where-Object { $_availableRecipients[$_rcptListProperty] -notcontains $_ }
            } else { $null }
        }
        
        # If any recipients were not found, inquire here if we should proceed:
        if ($_missingRecipients.Values.GetEnumerator().Count -gt 0) {
            
            Write-Warning "One or more backed-up recipients were not found in EXO.  Review the following missing recipients carefully."
            [PSCustomObject]$_missingRecipients

            if ($PSCmdlet.ShouldProcess(
                    "Continuing with group creation in EXO, knowing some recipients could not be found.",
                    "Are you sure you want to continue, knowing some recipients could not be found?",
                    "Group: $($DistributionGroup.PrimarySmtpAddress)"
                )) { <#  #> }
            elseif (-not $WhatIfPreference.IsPresent) {
                Write-Warning -Message 'Script abandoned due to missing recipients in EXO.'
                break
            }
        }
                
        # Add available Members:
        if ($_availableRecipients['Members']) { $_newDGParams['Members'] = $_availableRecipients['Members'] }
        
        # Apply remaining properties via Set-DistributionGroup:
        $_setDGParams = @{

            Identity                          = $_newDGParams['PrimarySmtpAddress']
            BypassSecurityGroupManagerCheck   = $true
            HiddenFromAddressListsEnabled     = $DistributionGroup.HiddenFromAddressListsEnabled
            ReportToManagerEnabled            = $DistributionGroup.ReportToManagerEnabled
            ReportToOriginatorEnabled         = $DistributionGroup.ReportToOriginatorEnabled
            SendOofMessageToOriginatorEnabled = $DistributionGroup.SendOofMessageToOriginatorEnabled
        }
        foreach ($_ca in 1..15) {
            $_thisCA = "CustomAttribute$($_ca)"
            if ($DistributionGroup.$_thisCA) { $_setDGParams[$_thisCA] = $DistributionGroup.$_thisCA }
        }
        foreach ($_rcptListProperty in 'ManagedBy', 'ModeratedBy',
            'AcceptMessagesOnlyFromSendersOrMembers', 'RejectMessagesFromSendersOrMembers', 'GrantSendOnBehalfTo', 'BypassModerationFromSendersOrMembers') {

            if ($_availableRecipients[$_rcptListProperty]) { $_setDGParams[$_rcptListProperty] = $_availableRecipients[$_rcptListProperty] }
        }
        if ($DistributionGroup.MailTip) { $_setDGParams['MailTip'] = $DistributionGroup.MailTip }

        # Special steps for placeholder-only mode:
        if ($PlaceholderOnly) {
            foreach ($_property in 'Name', 'DisplayName', 'PrimarySmtpAddress', 'Alias') {
                
                $_newDGParams[$_property] = "zzzTmpDLMigration_$($_newDGParams[$_property])"
            }
            $_setDGParams['Identity'] = "zzzTmpDLMigration_$($_setDGParams['Identity'])"
        }
        else { $_setDGParams['EmailAddresses'] = $DistributionGroup.EmailAddresses }

        # Confirm, then create the group:
        if ($PSCmdlet.ShouldProcess(
                "Creating a new group in EXO: ""$($_newDGParams['DisplayName'])"" <$($_newDGParams['PrimarySmtpAddress'])>",
                "Are you sure you want to create the new group ""$($_newDGParams['DisplayName'])"" <$($_newDGParams['PrimarySmtpAddress'])>?",
                "Group: $($DistributionGroup.PrimarySmtpAddress)"
            )) {
            New-DistributionGroup @_newDGParams -ErrorAction Stop
            Start-Sleep -Seconds 3
            Set-DistributionGroup @_setDGParams -ErrorAction Stop
            
            # Add Send-As:
            if ($_availableRecipients['SendAs']) {
                foreach ($_rcpt in $_availableRecipients['SendAs']) {
                    Add-RecipientPermission -Identity $_setDGParams['Identity'] -AccessRights SendAs -Trustee $_rcpt -ErrorAction Stop
                }
            }
        }
    }
    catch { throw }
}
#=========#-----------------#
#endregion# Recreate in EXO #
#=========#-----------------#



#======#------------------------#
#region# Update EXO Placeholder #
#======#------------------------#
if ($UpdateEXOPlaceholder) {
    try {
        $_setDGParams = @{
            Identity           = "zzzTmpDLMigration_$($DistributionGroup.PrimarySmtpAddress)"
            Name               = $DistributionGroup.DisplayName
            DisplayName        = $DistributionGroup.DisplayName
            PrimarySmtpAddress = $DistributionGroup.PrimarySmtpAddress
            Alias              = $DistributionGroup.Alias
        }

        # Confirm, then update the group:
        if ($PSCmdlet.ShouldProcess(
                "Updating placeholder group in EXO: ""$($_setDGParams['DisplayName'])"" <$($_setDGParams['PrimarySmtpAddress'])>",
                "Are you sure you want to update the group ""$($_setDGParams['DisplayName'])"" <$($_setDGParams['PrimarySmtpAddress'])>?",
                "Group: $($DistributionGroup.PrimarySmtpAddress)"
            )) {
            Set-DistributionGroup @_setDGParams -ErrorAction Stop
            Start-Sleep -Seconds 3
            Set-DistributionGroup $DistributionGroup.PrimarySmtpAddress -EmailAddresses $DistributionGroup.EmailAddresses -ErrorAction Stop
        }
    }
    catch { throw }
}
#=========#------------------------#
#endregion# Update EXO Placeholder #
#=========#------------------------#

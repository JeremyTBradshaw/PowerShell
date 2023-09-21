<#
    .SYNOPSIS
    Setup RBAC for Apps in Exchange Online.

    .DESCRIPTION
    Have an Azure AD Enterprise App (SPN) and a scoping group (in EXO) ready to go and supply the necessary info to
    this script to configure the EXO parts:
        - New-ServicePrincipal
        - New-DistributionGroup (targeted by the new management scope)
        - New-ManagementScope
        - New-ManagementRoleAssignment (per each role specified, each limited to the new management scope)

    .PARAMETER AADSPNDisplayName
    Specifies the DisplayName of the Enterprise App/service principal in Azure AD.

    .PARAMETER AADSPNAppId
    Specifies the ApplicationId (a.k.a. ClientId) of the Enterprise App/service principal in Azure AD.

    .PARAMETER AADSPNObjectId
    Specifies the ObjectId of the Enterprise App/service principal in Azure AD.

    .PARAMETER Roles
    Specifies one or more roles names to assign the new app.  Role names are documented here:
    https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac#supported-application-roles
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)][string]$AADSPNDisplayName,
    [Parameter(Mandatory)][object]$AADSPNAppId,
    [Parameter(Mandatory)][object]$AADSPNObjectId,
    [Parameter(Mandatory)]
    [ValidateSet(
        'Application Mail.Read', 'Application Mail.ReadBasic', 'Application Mail.ReadWrite', 'Application Mail.Send',
        'Application MailboxSettings.Read', 'Application MailboxSettings.ReadWrite',
        'Application Calendars.Read', 'Application Calendars.ReadWrite',
        'Application Contacts.Read', 'Application Contacts.ReadWrite',
        'Application Mail Full Access', 'Application Exchange Full Access',
        'Application EWS.AccessAsApp'
    )]
    [string[]]$Roles
)
begin {
    try {
        $_requiredCommands = @('New-ServicePrincipal', 'New-DistributionGroup', 'Set-DistributionGroup', 'New-ManagementScope', 'New-ManagementRoleAssignment')
        if (-not (Get-Command $_requiredCommands -ea:si)) {
            throw "Required command missing.  Required commands: $($_requiredCommands -join ', ')"
        }
        $progress = @{ Activity = $PSCmdlet.MyInvocation.MyCommand.Name }
        <# Role Assignment objects' names, among other things maybe, are limited to 64 characters.
        The longest permission name (MailboxSettings.ReadWrite) is 25 characters.
        "RBAC-for-Apps <app-display-name> - " takes up 17 characters.
        64 - 25 - 14 = 22 characters remaining.  We'll use the first 25 characters of the AAD SPN DisplayName:#>
        $commonDisplayName = $AADSPNDisplayName.SubString(0, ([math]::Min(22, $AADSPNDisplayName.Length)))

        Write-Progress @progress -PercentComplete 20
        New-ServicePrincipal -DisplayName $commonDisplayName -AppId $AADSPNAppId -ObjectId $AADSPNObjectId -ea:Stop | Out-Null

        Write-Progress @progress -PercentComplete 40
        $EXOSecDistGroup = New-DistributionGroup -Type Security -Name "RBAC-for-Apps Management Scope - $($commonDisplayName)" -ea:Stop
        Start-Sleep -Seconds 2
        $EXOSecDistGroup | Set-DistributionGroup -HiddenFromAddressListsEnabled:$true -ea:Stop

        Write-Progress @progress -PercentComplete 60
        New-ManagementScope -Name "RBAC-for-Apps $($commonDisplayName)" -RecipientRestrictionFilter "memberOfGroup -eq '$($EXOSecDistGroup.DistinguishedName)'" -ea:Stop | Out-Null

        $ht_Roles = @{
            # https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac#supported-application-roles #<-: 2023-09-18's info:
            # Name = Permissions List, Description
            'Application Mail.Read'                 = 'Mail.Read' #	Allows the app to read email in all mailboxes without a signed -in user.
            'Application Mail.ReadBasic'            = 'Mail.ReadBasic' #	Allows the app to read email except the body, previewBody, attachments, and any extended properties in all mailboxes without a signed -in user
            'Application Mail.ReadWrite'            = 'Mail.ReadWrite' # Allows the app to create, read, update, and delete email in all mailboxes without a signed -in user. Doesn't include permission to send mail.
            'Application Mail.Send'                 = 'Mail.Send' # Allows the app to send mail as any user without a signed-in user.
            'Application MailboxSettings.Read'      = 'MailboxSettings.Read' # Allows the app to read user's mailbox settings in all mailboxes without a signed -in user.
            'Application MailboxSettings.ReadWrite' = 'MailboxSettings.ReadWrite' # Allows the app to create, read, update, and delete user's mailbox settings in all mailboxes without a signed-in user.
            'Application Calendars.Read'            = 'Calendars.Read' # Allows the app to read events of all calendars without a signed-in user.
            'Application Calendars.ReadWrite'       = 'Calendars.ReadWrite' # Allows the app to create, read, update, and delete events of all calendars without a signed-in user.
            'Application Contacts.Read'             = 'Contacts.Read' # Allows the app to read all contacts in all mailboxes without a signed-in user.
            'Application Contacts.ReadWrite'        = 'Contacts.ReadWrite' # Allows the app to create, read, update, and delete all contacts in all mailboxes without a signed-in user.
            'Application Mail Full Access'          = 'Mail Full Access' # Mail.ReadWrite, Mail.Send # Allows the app to create, read, update, and delete email in all mailboxes and send mail as any user without a signed-in user.
            'Application Exchange Full Access'      = 'Exchange Full Access' # Mail.ReadWrite, Mail.Send, MailboxSettings.ReadWrite, Calendars.ReadWrite, Contacts.ReadWrite # Without a signed-in user: Allows the app to create, read, update, and delete email in all mailboxes and send mail as any user. Allows the app to create, read, update, and delete user's mailbox settings in all mailboxes. Allows the app to create, read, update, and delete events of all calendars. Allows the app to create, read, update, and delete all contacts in all mailboxes.
            'Application EWS.AccessAsApp'           = 'EWS.AccessAsApp' # Allows the app to use Exchange Web Services with full access to all mailboxes.
        }

        Write-Progress @progress -PercentComplete 80
        foreach ($role in $Roles) {

            $newMgmtRoleAssignmentParams = @{
                Name                = "RBAC-for-Apps $($commonDisplayName) - $($Script:ht_Roles[$role])"
                Role                = $role
                App                 = $commonDisplayName
                CustomResourceScope = "RBAC-for-Apps $($commonDisplayName)"
            }
            Write-Progress @progress -PercentComplete 90
            New-ManagementRoleAssignment @newMgmtRoleAssignmentParams -ea:Stop | Out-Null
            Write-Progress @progress -PercentComplete 100
        }

        "RBAC-for-Apps configuration is ready to go.  Add mailboxes to the group " +
        """RBAC-for-Apps Management Scope - $($commonDisplayName)"" for them to be accessible to the ""$($AADSPNDisplayName)"" app." |
        Write-Host -ForegroundColor Green
    }
    catch { throw }
}
process {}
end { Write-Progress @progress -Completed }

<#
    .Synopsis
    Create Exchange RBAC Roles and Assignments (direct to user) for the Intune Exchange Connector service account.

    .Parameter IntuneExchangeConnectorUserId
    Specify the Intune Exchange Connector service/user account.  Use any property that would be accepted by Get-User
    for the -Identity parameter.

    .Description
    Creates 4 new Management Roles, and 4 new direct-to-user Management Role Assignments, each named the same:

    - 'Intune Exchange Connector - Mail Recipients',
    - 'Intune Exchange Connector - Organization Client Access',
    - 'Intune Exchange Connector - Recipient Policies',
    - 'Intune Exchange Connector - View-Only Configuration'

    .Example
    .\New-IntuneExchangeConnectorManagementRolesAndDirectAssignments.ps1 -IntuneExchangeConnectorUserId svcIntuneExchConn -Verbose

    .Notes
    To remove the assignments and roles created by this script, run:
    
        [PS] \> Get-ManagementRoleAssignment "Intune Exchange Connector - *"

    Verify only the ones created by this script exist (there would be 4 max), then:

        [PS] \> Get-ManagementRoleAssignment "Intune Exchange Connector - *" | Remove-ManagementRoleAssignment -Confirm:$false

    Repeat these steps with Get-ManagementRole / Remove-ManagementRole.
#>
#Requires -Version 4
[CmdletBinding()]
param (
    [Parameter(
        Mandatory,
        HelpMessage = "Use any property that would be accepted by Get-User's -Identity parameter."
    )]
    [string]$IntuneExchangeConnectorUserId
)

begin {
    if (-not (Get-Command Get-ExchangeServer)) {

        Write-Warning -Message 'Connect to Exchange (or use Exchange Management Shell) before running this script.'
        break
    }

    if (-not (Get-User $IntuneExchangeConnectorUserId -ErrorAction SilentlyContinue)) {

        Write-Warning -Message "Couldn't find user account $($IntuneExchangeConnectorUserId).  Make sure it exists before running this script."
        break
    }
}

process {

    $Script:Continue = $true

    $Script:RequiredCmdlets = @(
        # Required Cmdlets per https://docs.microsoft.com/en-us/mem/intune/protect/exchange-connector-install#exchange-cmdlet-requirements
        'Get-ActiveSyncOrganizationSettings',
        'Set-ActiveSyncOrganizationSettings',
        'Get-CasMailbox',
        'Set-CasMailbox',
        'Get-ActiveSyncMailboxPolicy',
        'Set-ActiveSyncMailboxPolicy',
        'New-ActiveSyncMailboxPolicy',
        'Remove-ActiveSyncMailboxPolicy',
        'Get-ActiveSyncDeviceAccessRule',
        'Set-ActiveSyncDeviceAccessRule',
        'New-ActiveSyncDeviceAccessRule',
        'Remove-ActiveSyncDeviceAccessRule',
        'Get-ActiveSyncDeviceStatistics',
        'Get-ActiveSyncDevice',
        'Get-ExchangeServer',
        'Get-ActiveSyncDeviceClass',
        'Get-Recipient',
        'Clear-ActiveSyncDevice',
        'Remove-ActiveSyncDevice',
        'Set-ADServerSettings',
        'Get-Command'
    )
    $Script:AlreadyEnabledCmdlets

    $ParentRoles = @(
        # Minimum parent roles containing the required Cmdlets:
        'Mail Recipients',
        'Organization Client Access',
        'Recipient Policies',
        'View-Only Configuration'
    )

    Write-Verbose -Message 'Checking for already-existing management roles/assignments that would have been created by this script.'
    foreach ($pr in $ParentRoles) {

        $NameForRoleComponents = "Intune Exchange Connector - $($pr)"
        if (Get-ManagementRoleAssignment $NameForRoleComponents -ErrorAction SilentlyContinue) {

            Write-Warning -Message "Management Role Assignment '$($NameForRoleComponents)' already exists.  Script won't proceed with role/assignment creations."
            $Script:Continue = $false
        }
        if (Get-ManagementRole $NameForRoleComponents -ErrorAction SilentlyContinue) {

            Write-Warning -Message "Management Role '$($NameForRoleComponents)' already exists.  Script won't proceed with role/assignment creations."
            $Script:Continue = $false
        }
    }

    if ($Script:Continue) {

        foreach ($pr in $ParentRoles) {

            $NameForRoleComponents = "Intune Exchange Connector - $($pr)"

            $prEnabledCmdlets = Get-ManagementRoleEntry "$($pr)\*"
            $prEnabledCmdlets = $prEnabledCmdlets |
            Where-Object { ($Script:RequiredCmdlets -contains $_.Name) -and ($Script:AlreadyEnabledCmdlets -notcontains $_.Name) }

            $Script:AlreadyEnabledCmdlets += $prEnabledCmdlets.Name
    
            try {
                Write-Verbose -Message "Creating new management role '$($NameForRoleComponents)' with the following enabled Cmdlets: `n`t$($prEnabledCmdlets.Name -join ""`n`t"" )."
                New-ManagementRole -Name $NameForRoleComponents -Parent "$($pr)" -EnabledCmdlets @($prEnabledCmdlets.Name) -ErrorAction Stop | Out-Null

                Write-Verbose -Message "Creating new management role assignment '$($NameForRoleComponents)' named exactly the same as the new role itself, and assigned to user: '$($IntuneExchangeConnectorUserId)'."
                New-ManagementRoleAssignment -Name $NameForRoleComponents -Role $NameForRoleComponents -User $IntuneExchangeConnectorUserId -ErrorAction Stop | Out-Null
            }
            catch {
                Write-Warning -Message "Script failure.  Error exception:`n$($_)"
                break
            }
        }
    }
}

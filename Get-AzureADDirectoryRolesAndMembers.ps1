<#
    .Synopsis
    Get all Azure AD roles and their members.  Really just combining Get-AzureADDirectoryRole and
    Get-AzureADDirectoryRoleMember
#>
#Requires -Module AzureAD
#Requires -Version 5.1
[CmdletBinding()]
param()

try {
    Get-AzureADTenantDetail -ErrorAction Stop | Out-Null
}
catch {
    Write-Warning -Message 'Connect-AzureAD before running this script.'
    break
}

$AzureADDirectoryRoles = Get-AzureADDirectoryRole
foreach ($role in $AzureADDirectoryRoles) {

    $Members = Get-AzureADDirectoryRoleMember -ObjectId $Role.ObjectId

    foreach ($member in $Members) {

        [PSCustomObject]@{

            RoleObjectId = $role.ObjectId
            RoleDisplayName = $role.DisplayName -replace 'Company Administrator', 'Global Administrator'
            MemberObjectId = $member.ObjectId
            MemberDisplayName = $member.DisplayName
            MemberUserPrincipalName = $member.UserPrincipalName
        }
    }
}

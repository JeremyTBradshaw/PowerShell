<#
    .Synopsis
    Get all direct and nested distribution group members.

    .Notes
    As of March 10, 2022, the ExchangeOnlineManagement module version 2.0.6-preview3 or higher is needed because the
    cmdlets involved (Get-DistributionGroup, Get-DistributionGroupMember) are REST-backed in these versions, meaning
    the script will have a chance at surviving in large orgs with many distribution groups/members.

    If users are a member via multiple avenues (i.e., direct member and/or membership in one or more nested groups)
    they will show up in the output as many times as necessary.  For this reason it is recommended to use a pivot table
    to summarize the results.  The pivot table rows should be in order of: top level group > member > directGroupId.

    .Parameter Identity
    Supply one or more distribution groups to this parameter.  Follow the instructions from Microsoft for the -Identity
    parameter in the Get-DistributionGroupMember documentation.

    .Example
    $DGs = Get-DistributionGroup -ResultSize Unlimited; $DGMembers = .\Get-EXODistributionGroupMemberRecursive -Identity $DGs.PrimarySmtpAddress

    .Example
    .\Get-EXODistributionGroupMemberRecursive -Identity AllStaff@contoso.com | Export-Csv $Home\AllStaff@Contoso.com_Members.csv -NTI
#>
#Requires -Modules @{ ModuleName = 'ExchangeOnlineManagement'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'; RequiredVersion = '2.0.6'}
[CmdletBinding()]
param (
    [Parameter(Mandatory, HelpMessage = "Follow Microsoft's documentation for the Identity parameter on the Get-DistributionGroupMember cmdlet.")]
    [string[]]$Identity
)

if (-not (Get-Command Get-DistributionGroupMember)) {

    throw "Please connect to Exchange Online before running this script (i.e., Connect-ExchangeOnline)."
}

function getMembers ([string]$groupId, [switch]$TopLevelGroup) {
    try {
        if ($Script:Stopwatch.ElapsedMilliseconds -gt 200) {

            Write-Progress @Script:Progress -CurrentOperation "Current group: $($groupId)" -PercentComplete (($pCounter / $Identity.Count) * 100)
            $Script:Stopwatch.Restart()
        }

        $members = @(Get-DistributionGroupMember -Identity "$($groupId)" -ResultSize Unlimited -ErrorAction SilentlyContinue)
        foreach ($member in $members) {

            $thisMember = [PSCustomObject]@{

                MemberDisplayName = $member.DisplayName
                MemberPSmtp       = $member.PrimarySmtpAddress
                MemberType        = $member.RecipientTypeDetails
                MembershipType    = if ($TopLevelGroup) { 'Direct' } else { 'Nested' }
                TopLevelGroupId   = $Script:TopLevelGroupId
                ParentGroupId     = $groupId
            }

            if ($member.RecipientTypeDetails -like 'Mail*Group') {

                if ($Script:Groups[$member.PrimarySmtpAddress]) { $thisMember.MembershipType = 'Redundantly Nested' }
                else {
                    $Script:Groups.Add($member.PrimarySmtpAddress,$true)
                    getMembers -groupId $member.PrimarySmtpAddress
                }
            }

            $thisMember
        }
    }
    catch {
        Write-Warning -Message "Failed in getMembers function for group ID '$($groupID)'."
        throw
    }
}

try {
    $Script:Progress = @{ Activity = $PSCmdlet.MyInvocation.MyCommand.Name }
    $Script:Stopwatch = [System.Diagnostics.Stopwatch]::new()
    $Script:Stopwatch.Start()

    $pCounter = 0
    foreach ($id in $Identity) {

        $Script:Progress['Status'] = "Getting group members recursively for $($id)"
        $Script:pCounter++
        $Script:TopLevelGroupId = $id
        $Script:Groups = @{}

        getMembers -groupId $id -TopLevelGroup
    }
}
catch { throw }

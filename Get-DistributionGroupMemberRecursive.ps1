<#
    .Synopsis
    Find distribution group members recursively, when there are nested member-groups.

    .Parameter StartingGroupPSMTP
    Specifies the PrimarySmtpAddress (ideally) for the top-level group to find members of.

    .Parameter LevelsDeepToGo
    Specifies how many levels of nesting to recurse.  Default is 10.  In case of an infinite nesting, it will be best
    to specify a number low enough to avoid wasting time/resources going in circles.

    .Parameter StartingLevelOverride
    In case this script is being run to resume an earlier effort, and the output will be combined with the output from
    the previous runs of the script, this will ensure the levels are consistent.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$StartingGroupPSMTP,

    [int]$LevelsDeepToGo = 10,

    # Optional parameter in case we want to resume an earlier operation:
    [int]$StartingLevelOverride
)

# Exit script if we do not have the necessary commands available:
if (-not (Get-Command Get-Recipient, Get-DistributionGroup, Get-DistributionGroupMember, Get-DynamicDistributionGroup -ErrorAction SilentlyContinue)) {

    throw "This script requires an Exchange/EXO PowerShell session and the commands Get-Recipient, Get-DistributionGroup, Get-DistributionGroupMember, and Get-DynamicDistributionGroup."
}



#======#----------#
#region Functions #
#======#----------#

function getGroup ($groupId) {
    try {
        $group = Get-Recipient -Identity "$($groupId)" -ErrorAction SilentlyContinue

        # The following non-standard section avoids Get-Recipient failing on recipient type "ExchangeSecurityGroup" which will cause the command to fail.
        if ($group) {
            if ($group.RecipientTypeDetails -notlike '*Group*') {

                throw "$($groupPSMTP) is not a static nor dynamic group.  Its RecipientTypeDetails value is: $($group.RecipientTypeDetails)"
            }
    
            $group
        }
    }
    catch {
        Write-Debug "getGroup: Found a bad recipient"
        Write-Warning "Failed on getGroup.  groupId: $($groupId)"
        throw
    }
}

function getGroupMember ($group, $Level) {
    try {
        if ($group.RecipientTypeDetails -notlike '*Dynamic*') {
    
            Get-DistributionGroupMember -Identity "$($group.Guid.ToString())" -ResultSize Unlimited |
            Select-Object @{Name='ParentGroup';Expression={if ($Level -eq 1) {'#N/A'} else { $group.PrimarySmtpAddress }}},
            @{Name='Level';Expression={$Level}},
            RecipientTypeDetails, PrimarySmtpAddress, DisplayName, Guid
        }
        else {
            $dynamicGroup = Get-DynamicDistributionGroup -Identity "$($group.Guid.ToString())" -ErrorAction Stop
            Get-Recipient -RecipientPreviewFilter $dynamicGroup.RecipientFilter -OrganizationalUnit $dynamicGroup.RecipientContainer -ResultSize Unlimited |
            Select-Object @{Name='ParentGroup';Expression={if ($Level -eq 1) {'#N/A'} else { $group.PrimarySmtpAddress }}},
            @{Name='Level';Expression={$Level}},
            RecipientTypeDetails, PrimarySmtpAddress, DisplayName, Guid
        }
    }
    catch {
        Write-Debug "getGroupMember: Found a bad group"
        Write-Warning "Failed on getGroupMember.  Group GUID: $($group.Guid), Group PSMTP: $($group.PrimarySmtpAddress), Group DisplayName: $($group.DisplayName), Level: $($Level)"
        throw
    }
}

#=========#----------#
#endregion Functions #
#=========#----------#



#======#------------#
#region Main Script #
#======#------------#

try {
    $ProgressSplat = @{

        Activity = 'Getting distribution group members'
        PercentComplete = -1
    }

    $StartingGroup = getGroup -groupId "$($StartingGroupPSMTP)"

    $Level = if ($PSBoundParameters.ContainsKey('StartingLevelOverride')) { $StartingLevelOverride } else { 1 }

    Write-Progress @ProgressSplat -Status "Finding level $($Level) members"
    
    $Members = @(getGroupMember -group $StartingGroup -Level 1)
    
    do {
        $ParentLevel = $Level
        $Level++
        foreach ($g in ($Members | Where-Object {$_.Level -eq $ParentLevel -and $_.RecipientTypeDetails -like '*Group*' -and (-not ($_.RecipientTypeDetails -eq 'ExchangeSecurityGroup'))})) {
    
            Write-Progress @ProgressSplat -Status "Finding members of level $($ParentLevel) member groups"
    
            $Members += getGroupMember -group (getGroup -groupId $g.Guid.ToString()) -Level $Level
        }
    }
    until ($Level -eq $LevelsDeepToGo)
    
    # Output all members:
    $Members
}
catch { throw }

#=========#------------#
#endregion Main Script #
#=========#------------#

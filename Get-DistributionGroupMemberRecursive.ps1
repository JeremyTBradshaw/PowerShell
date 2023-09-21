<#
    .Synopsis
    Find distribution group members recursively, when there are nested member-groups.

    .Parameter Identity
    Specifies the identity for the top-level group to find members of.  This is a direct passthrough to Get-Recipient's
    -Identity parameter.

    .Parameter LevelsDeepToGo
    Specifies how many levels of nesting to recurse.  Default is 10.  In case of an infinite nesting, it will be best
    to specify a number low enough to avoid wasting time/resources going in circles.

    .Parameter StartingLevelOverride
    In case this script is being run to resume an earlier effort, and the output will be combined with the output from
    the previous runs of the script, this will ensure the levels are consistent.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline)]
    [string[]]$Identity,

    [int]$LevelsDeepToGo = 10,

    # Optional parameter in case we want to resume an earlier operation:
    [int]$StartingLevelOverride
)

begin {
    # Exit script if we do not have the necessary commands available:
    if (-not (Get-Command Get-Recipient, Get-DistributionGroup, Get-DistributionGroupMember, Get-DynamicDistributionGroup -ErrorAction SilentlyContinue)) {

        throw "This script requires an Exchange/EXO PowerShell session and the commands Get-Recipient, Get-DistributionGroup, Get-DistributionGroupMember, and Get-DynamicDistributionGroup."
    }



    #======#-----------#
    #region# Functions #
    #======#-----------#

    function getGroup ($groupId) {
        try {
            $group = Get-Recipient -Identity "$($groupId)" -ErrorAction Stop

            if ($group) {
                if ($group.Count -gt 1) {

                    throw "Multiple recipient objects were found for group '$($groupId)'.  Unable to process this group ID."
                }
                elseif ($group.RecipientTypeDetails -notlike 'Mail*Group' -and $group.RecipientTypeDetails -ne 'DynamicDistributionGroup') {

                    throw "$($groupId) is not a typical static nor dynamic distribution group.  Its RecipientTypeDetails value is: $($group.RecipientTypeDetails)"
                }
                else { $group }
            }
        }
        catch {
            Write-Warning "Failed on getGroup.  groupId: $($groupId)"
            throw
        }
    }

    function getGroupMember ($group, $Level) {
        try {
            if ($group.RecipientTypeDetails -ne 'DynamicDistributionGroup') {

                Get-DistributionGroupMember -Identity "$($group.Guid.ToString())" -ResultSize Unlimited -ErrorAction Stop |
                Select-Object @{Name = 'ParentGroup'; Expression = { $group.PrimarySmtpAddress } },
                @{Name = 'Level'; Expression = { $Level } },
                RecipientTypeDetails, PrimarySmtpAddress, DisplayName, Guid
            }
            else {
                $dynamicGroup = Get-DynamicDistributionGroup -Identity "$($group.Guid.ToString())" -ErrorAction Stop
                Get-Recipient -RecipientPreviewFilter $dynamicGroup.RecipientFilter -OrganizationalUnit $dynamicGroup.RecipientContainer -ResultSize Unlimited |
                Select-Object @{Name = 'ParentGroup'; Expression = { $group.PrimarySmtpAddress } },
                @{Name = 'Level'; Expression = { $Level } },
                RecipientTypeDetails, PrimarySmtpAddress, DisplayName, Guid
            }
        }
        catch {
            Write-Warning "Failed on getGroupMember.  Group GUID: $($group.Guid), Group PSMTP: $($group.PrimarySmtpAddress), Group DisplayName: $($group.DisplayName), Level: $($Level)"
            throw
        }
    }

    #=========#-----------#
    #endregion# Functions #
    #=========#-----------#
}
process {

    #======#-------------#
    #region# Main Script #
    #======#-------------#

    try {
        $Progress = @{

            Activity         = 'Getting distribution group members'
            PercentComplete  = -1
            CurrentOperation = "Current (starting) parent group: $($Identity)"
        }

        try { $StartingGroup = getGroup -groupId "$($Identity)" }
        catch { Write-Warning $_; return }

        $Level = if ($PSBoundParameters.ContainsKey('StartingLevelOverride')) { $StartingLevelOverride } else { 1 }

        Write-Progress @Progress -Status "Finding level $($Level) members"

        try { $Members = @(getGroupMember -group $StartingGroup -Level $Level) }
        catch { Write-Warning $_; return }

        do {
            $ParentLevel = $Level
            $Level++
            foreach ($g in ($Members | Where-Object { $_.Level -eq $ParentLevel -and ($_.RecipientTypeDetails -like 'Mail*Group' -or $_.RecipientTypeDetails -eq 'DynamicDistributionGroup') })) {

                Write-Progress @Progress -Status "Finding members of level $($ParentLevel) member groups"

                try { $Members += getGroupMember -group (getGroup -groupId $g.Guid.ToString()) -Level $Level }
                catch { Write-Warning $_; return }
            }
        }
        until ($Level -eq $LevelsDeepToGo)

        # Output all members:
        $Members
    }
    catch { throw }

    #=========#-------------#
    #endregion# Main Script #
    #=========#-------------#
}
end { Write-Progress @Progress Completed }

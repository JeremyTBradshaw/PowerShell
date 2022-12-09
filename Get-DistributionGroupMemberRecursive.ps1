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
if (-not (Get-Command Get-DistributionGroup, Get-DistributionGroupMember -ErrorAction SilentlyContinue)) {

    throw "This script requires an Exchange/EXO PowerShell session and the commands Get-DistributionGroup, Get-DistributionGroupMember."
}

$Level = if ($PSBoundParameters.ContainsKey('StartingLevelOverride')) { $StartingLevelOverride } else { 1 }

$ProgressSplat = @{

    Activity = 'Getting distribution group members'
    PercentComplete = -1
}

Write-Progress @ProgressSplat -Status "Finding level $($Level) members"

$Members = @(Get-DistributionGroupMember $StartingGroupPSMTP -ResultSize Unlimited |
Select-Object @{Name='ParentGroup';Expression={'#N/A'}},
@{Name='Level';Expression={$Level}},
RecipientTypeDetails, PrimarySmtpAddress, DisplayName)

do {
    $ParentLevel = $Level
    $Level++
    foreach ($g in ($Members | Where-Object {$_.Level -eq $ParentLevel -and $_.RecipientTypeDetails -like '*group*'})) {

        Write-Progress @ProgressSplat -Status "Finding members of level $($ParentLevel) member groups"

        $Members += Get-DistributionGroupMember $g.PrimarySmtpAddress.ToString() -ResultSize Unlimited |
        Select-Object @{Name='ParentGroup';Expression={$g.PrimarySmtpAddress}},
        @{Name='Level';Expression={$Level}},
        RecipientTypeDetails, PrimarySmtpAddress, DisplayName
    }
}
until ($Level -eq ($LevelsDeepToGo))

# Output all members:
$Members

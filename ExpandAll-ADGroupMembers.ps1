<#
  .Synopsis
  Get AD group user memberships recursively.  Intended for multi-domain forests where Get-ADGroupMember -Recursive
  won't suffice.

  .Parameter DistinguishedName
  The DistinguishedName of the group whose members should be expanded.  DN is required to identify which domain to
  contact directly for Get-ADGroupMember (without -Recursive).
#>
#Requires -Version 3
#Requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [Parameter(
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName
    )]
    [ValidateScript(
        {
            if ($_ -match '^CN=.*DC=.*') { $true }
            else { throw "'$($_)' doesn't appear to be a valid DistinguishedName." }
        }
    )]
    [string]$DistinguishedName
)

begin {

    function getADGroupMember ([string]$groupDN) {
        try {
            $domain = ($groupDN -replace '^.*?,DC=', '') -replace ',DC=', '.'

            $directMembers = Get-ADGroupMember -Identity $groupDN -Server $domain -ErrorAction:Stop
            foreach ($dm in $directMembers) {

                if ($dm.ObjectClass -eq 'user') {
                    # Add previously unseen users to the main member users collection:
                    if (-not $Script:MemberUsers[$dm.ObjectGuid.Guid]) {

                        $Script:MemberUsers[$dm.ObjectGuid.Guid] = $dm | Select-Object @{Name = 'Group'; Expression = { $groupDN } }, *
                    }
                }
                elseif ($dm.ObjectClass -eq 'group') {
                    # Recursively get previously unseen nested groups' members:
                    if (-not $Script:MemberGroups[$dm.ObjectGuid.Guid]) {

                        getADGroupMember -groupDN $dm.DistinguishedName
                    }
                }
            }
        }
        catch { throw }
    }

    $Progress = @{
        Activity        = "$($PSCmdlet.MyInvocation.MyCommand.Name) - Start time: $([datetime]::Now)"
        PercentComplete = -1
    }
}

process {

    $Script:MemberGroups = @{}
    $Script:MemberUsers = @{}

    try {
        Write-Progress @Progress -Status 'Getting group members recursively...' -CurrentOperation $DistinguishedName

        getADGroupMember -groupDN $DistinguishedName
        $Script:MemberUsers.GetEnumerator() | Select-Object -ExpandProperty Value
    }
    catch {
        Write-Warning "Failed to get group members for group '$($DistinguishedName)'."
        throw
    }

}
end { Write-Progress @Progress -Completed }

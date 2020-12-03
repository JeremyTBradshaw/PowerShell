<#

  .Synopsis
  Get AD group user memberships recursively.  The DistinguishedName of the
  group(s) is required to determine which domain controllers to contact.

  The intended use case for this script is multi-domain forests where
  Get-ADGroupMember won't suffice.

  .Parameter DistinguishedName
  [Required] The DistinguishedName of the group whose members should be
  expanded.

  .Parameter ExpandOutput
  [Optional] Expands output so that each member user is outputted as single
  object with two properties: GroupDN and MemberId.

  When not specified, each group is output as a single object with two
  properties: GroupDN and MemberIds (semicolon-delimited).

  .Parameter OutputMemberGuids
  [Optional] Changes output for members to ObjectGuid instead of
  DistinguishedName.  This can be beneficial for file size purposes.

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
            else { throw 'Invalid DistinguishedName format.' }
        }
    )]
    [Alias('DistinguishedName')]
    [string]$DN,

    [switch]$ExpandOutput,

    [switch]$OutputMemberGuids

)

begin { }

process {

    $Global:MemberGroups = @()
    $Global:MemberUsers = @()
  
    function expandGroup ($groupDN) {

        $domain = $null
        $domain = ($groupDN -replace '^.*?,DC=', '') -replace ',DC=', '.'
    
        try {
            $directMembers = $null
            $directMembers = Get-ADGroupMember -Identity $groupDN -Server $domain -ErrorAction:Stop
        }
        catch {
            Write-Warning -Message "Failure for group $($groupDN).`nCommand: Get-ADGroupMember -Identity '$($groupDN)' -Server $($domain) -ErrorAction:Stop`nError Exception: $($_.Exception)"
        }

        $directMembers |
        ForEach-Object {

            $memberDN   = $_.DistinguishedName
            $memberGuid = $_.ObjectGuid
            $memberType = $_.ObjectClass

            switch ($_.ObjectClass) {
        
                group {
                    switch ($OutputMemberGuids) {
                        $true   {$memberId = $memberGuid}
                        $false  {$memberId = $memberDN}
                    }

                    if ($Global:MemberGroups -notcontains [string]$memberDN) {

                        $member = [PSCustomObject]@{
                            GroupDN     = $DN
                            MemberId    = $memberId
                            MemberType  = $memberType
                        }
                        Write-Output $member

                        expandGroup $memberDN

                        $Global:MemberGroups += [string]$memberDN
                        $Global:MemberGroups += [string]$groupDN
                    }
                }

                user {
                    switch ($OutputMemberGuids) {
                        $true   {$memberId = $memberGuid}
                        $false  {$memberId = $memberDN}
                    }

                    if ($Global:MemberUsers -notcontains [string]$memberId) {

                        $member = [PSCustomObject]@{
                            GroupDN     = $DN
                            MemberId    = $memberId
                            MemberType  = $memberType
                        }
                        Write-Output $member

                        $Global:MemberUsers += [string]$memberId
                    }
                }

            }
        }
    } # end function expandGroup


    $ProgressProps = @{
        Activity         = "ExpandAll-ADGroupMembers.ps1"
        Status           = '...expanding...'
        CurrentOperation = "Get-ADGroupMember -Identity '$($DN)' -Server $($domain)'"
    }
    Write-Progress @ProgressProps

    $Expanded = $null
    $Expanded = expandGroup $DN

    switch ($ExpandOutput) {

        $true {Write-Output -InputObject $Expanded}

        $false {
            if ($null -ne $Expanded) {

                $Collapsed = [PSCustomObject]@{

                    GroupDN   = $Expanded.GroupDN | Select-Object -Unique
                    MemberIds = $Expanded.MemberId -join ';'
                }
                Write-Output $Collapsed
            }
        }

    }
}

end { Write-Progress -Activity "ExpandAll-ADGroupMembers.ps1" -Completed }

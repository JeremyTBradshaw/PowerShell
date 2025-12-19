function getAADGroupBasedAssignedLicenses {

    try {
        $groups = newMGRequest -Request "/groups?`$select=assignedLicenses,displayName,id"
        $memberAssignees = @{}
        foreach ($group in ($groups.value | Where-Object { $_.assignedLicenses })) {

            $grpMembers = newMGRequest -Request "/groups/$($group.id)/transitiveMembers"
            foreach ($gm in $grpMembers.value) {
                foreach ($lic in $group.assignedLicenses.skuId) {

                    if (-not $memberAssignees[$gm.id]) {$memberAssignees[$gm.id] = @{} }
                    if (-not $memberAssignees[$gm.id][$ht_skuId[$lic]]) { $memberAssignees[$gm.id][$ht_skuId[$lic]] = @{} }
                    if (-not $memberAssignees[$gm.id][$ht_skuId[$lic]]['Groups']) { $memberAssignees[$gm.id][$ht_skuId[$lic]]['Groups'] = @() }
                    $memberAssignees[$gm.id][$ht_skuId[$lic]]['Groups'] += $group
                }
            }
        }
        $memberAssignees
    }
    catch { throw }
}

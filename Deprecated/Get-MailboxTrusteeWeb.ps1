<#
    .Synopsis
    Identify webs of mailbox-trustee relationships, in the following manner:

    1.) Reverse Lookup: Find all mailboxes that one or more trustee users have
        access to.

    2.) Forward Lookup: For these same users, find all other users who have access
        to their mailbox.

    3.) Recursively repeat this process against all found mailboxes and trustees
        until the entire web is known (or set a maximum recursion depth).

    .Description
    This is the 2nd of 2 brother scripts to Get-MailboxTrustee.ps1.  Therefore it
    requires that the passed CSV file(s) contain(s) a few headers for matching
    properties output by Get-MailboxTrustee.ps1.  Any other headers (i.e. columns)
    are ignored.

    At this time, the output from the MinimizeOutput mode of Get-MailboxTrustee.ps1
    is not supported.  Instead only the standard mode's output is accepted.

    Mandatory CSV headers:

    - MailboxPSmtp
    - MailboxType
    - PermissionType
    - AccessRights
    - TrusteePSmtp
    - TrusteeType

    .Parameter GetMailboxTrusteeCsvFilePath
    The full or direct path to the CSV file(s) containing objects that have been
    output from Get-MailboxTrustee.ps1.  **Note: Get-MailboxTrustee.ps1's
    MinimizedOutput mode is not support (see description).

    Since it is common to have run Get-MailboxTrustee.ps1 in multiple passes,
    multiple CSV files can be specified here, to spare the need to combine CSV
    files manually.  Example scenarios for multiple CSV files:

     - Get-MailboxTrustee.ps1 was run separately against EXO then Exchange
       on-premises.
     - Get-MailboxTrustee.ps1 was run in multiple jobs or runspaces.

    .Parameter StartingPSmtp
    Specify one or more mailboxes and/or trustees to search recursively as described
    in the synopsis.  Include all SMTP addresses that are intended to be in the same
    web (think 'web' = 'migration batch').

    The reason PrimarySmtpAddress is used for the starting ID property is that this
    script's original intended purpose is to determine mailbox dependencies for
    migration batch planning.  As such, if a user doesn't have a mailbox of their
    own (i.e. no email address), they will not be impacted by being left out of
    this exercise, since they themselves will not need to be migrated.

    .Parameter PermissiveMailboxThreshold
    Some mailboxes have very many trustees.  Examples are popular room and equipment
    mailboxes.  This can plague the process of identifying webs that are truly
    significant.  Use this parameter to specify the maximum # of 1 to 1
    relationships mailboxes can have before they will be excluded from the web.

    - Default: 500
    - Minimum: 1
    - Maximum: Positive [Int32]

    .Parameter PowerTrusteeThreshold
    Some users have access to very many mailboxes.  Examples are administrators and
    Power Users.  Use this parameter to specify the maximum # of 1 to 1
    relationships trustees can have before they will be excluded from the web.

    - Default: 500
    - Minimum: 1
    - Maximum: Positive [Int32]

    .Parameter MaximumDepth
    How many times to recurse when performing the forward and reverse trustee
    lookups.  This can help if the full web of permission relationships is a larger
    group of mailboxes than desired for the task at hand, for example, a migration
    batch.

    - Default: 100
    - Minimum: 1
    - Maximum: Positive [Int32]

    .Parameter IgnoreMailboxPSmtp
    One or more mailboxes to ommit from searches and output.  Expected input is
    PrimarySmtpAddress (of the mailbox (i.e. trust'ING user)), comma-separated or
    an array variable if supplying multiple.

    .Parameter IgnoreTrusteePSmtp
    One or more trustee users to ommit from searches and output. Expected input is
    PrimarySmtpAddress (of the trustee (i.e. trust'ED user), comma-separated or an
    array variable if supplying multiple.

    .Parameter IgnorePermissionType
    One or more of the following permissions types to ommit from searches and
    output:

    - FullAccess
    - SendAs
    - SendOnBehalf
    - AllFolders (excludes all mailbox folder permissions)
    - MailboxRoot
    - Inbox
    - Calendar
    - Contacts
    - Tasks
    - SentItems

    .Example
    .\Get-MailboxTrusteeWeb.ps1 -GetMailboxTrusteeCsvFilePath .\Desktop\All-Mailbox-Trustees-2018-10-17 `
                                    -PermissiveMailboxThreshold 30 `
                                    -PowerTrusteeThreshold 10 `
                                    -IgnoreTrusteePSmtp BESAdmin@blackberry.com `
                                    -IgnoreMailboxPSmtp Executive@contoso.com, ExecAdmin@contoso.com, ExecMeetingRm@contoso.com `
                                    -IgnorePermissionTypes SendOnBehalf, Calendar, Contacts `
                                    -


    .Example
    $AllMailboxes = Get-Mailboxes -ResultSize:Unlimited

    PS C:\> $AllMailboxes | .\Get-MailboxTrustee.ps1 -ExpandTrusteeGroups | Export-Csv .\All-Mailbox-Trustees.csv -NTI

    PS C:\> $Optimzed = .\Optimize-MailboxTrusteeWebInput.ps1 `
                            -GetMailboxTrusteeCsvFilePath .\All-Mailboxes-Trustees.csv `
                            -PermissiveMailboxThreshold 500 `
                            -PowerTrusteeThreshold 500

    PS C:\> $Optimized | Export-Csv .\All-Mailbox-Trustees-Optimized.csv -NTI

    PS C:\> $MigrationBatch1ProposedList = Import-Csv .\MigrationBatch1Users.csv

    PS C:\> $MigBatch1UsersWeb = .\Gete-MailboxTrusteeWebInut.ps1 `
                                    -GetMailboxTrusteeCsvFilePath .\All-Mailbox-Trustees-Optimized.csv `
                                    -StartingPSmtp $MigrationBatch1ProposedList.PrimarySmtpAddress `
                                    -OptimizedInput

    PS C:\> $MigBatch1UsersWeb | Export-Csv .\MigBatchUsers1_MailboxTrusteeWeb.csv -NTI

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrusteeWeb.ps1
    # ^ Get-MailboxTrusteeWeb.ps1

    .Link
    # Get-MailoxTrusteeWebSQLEdition.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrusteeWebSQLEdition.ps1

    # Get-MailboxTrustee.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrustee.ps1

    # Optimize-MailboxTrusteeWebInput.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Optimize-MailboxTrusteeWebInput.ps1

    # New-MailboxTrusteeReverseLookup.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/New-MailboxTrusteeReverseLookup.ps1
#>

#Requires -Version 3
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript( {
        $_ | ForEach-Object {
            if ((Test-Path -Path $_) -eq $false) {throw "Can't find file '$($_)'."}
            if ($_ -notmatch '(\.csv$)') {throw "Only .csv files are accepted."}
            $true
        }
    })]
    [System.IO.FileInfo[]]$GetMailboxTrusteeCsvFilePath,

    [Parameter(Mandatory = $true)]
    [ValidateScript({
        $_ | ForEach-Object {

            if ($_.Length -gt 320) {throw 'Must be 320 characters or less (maximum for SMTP address)'}
            elseif ($_ -notmatch '^.*\@.*\..*$') {throw "'$($_)' is not a valid SMTP address."}
            else {$true}
        }
    })]
    [string[]]$StartingPSmtp,

    [ValidateRange(1,[int32]::MaxValue)]
    [int]$PermissiveMailboxThreshold = 500,

    [ValidateRange(1,[int32]::MaxValue)]
    [int]$PowerTrusteeThreshold = 500,

    [ValidateRange(1,[int32]::MaxValue)]
    [int]$MaximumDepth = 100,

    [ValidateScript({
        $_ | ForEach-Object {

            if ($_.Length -gt 320) {throw 'Must be 320 characters or less (maximum for SMTP address)'}
            elseif ($_ -notmatch '^.*\@.*\..*$') {throw "'$($_)' is not a valid SMTP address."}
            else {$true}
        }
    })]
    [string[]]$IgnoreTrusteePSmtp,

    [ValidateScript({
        $_ | ForEach-Object {

            if ($_.Length -gt 320) {throw 'Must be 320 characters or less (maximum for SMTP address)'}
            elseif ($_ -notmatch '^.*\@.*\..*$') {throw "'$($_)' is not a valid SMTP address."}
            else {$true}
        }
    })]
    [string[]]$IgnoreMailboxPSmtp,

    [ValidateSet(
        'FullAccess', 'SendAs', 'SendOnBehalf',
        'AllFolders', 'MailboxRoot', 'Inbox', 'Calendar', 'Contacts', 'Tasks', 'SentItems'
    )]
    [string[]]$IgnorePermissionType

)

begin {
    $StartTime = Get-Date
    $MainProgress = @{

        Activity            = "Get-MailboxTrusteeWeb.ps1 (Start time: $($StartTime.DateTime))"
        Id                  = 0
        ParentId            = -1
        Status              = 'Initializing'
        PercentComplete     = -1
        SecondsRemaining    = -1
    }

    Write-Progress @MainProgress

    $StandardHeaders = @(

        'MailboxPSmtp',
        'MailboxType',
        'PermissionType',
        'AccessRights',
        'TrusteePSmtp',
        'TrusteeType'
    )

    $CsvCounter = 0

    $MainProgress['Status'] = 'Importing CSV file(s)'

    $GetMailboxTrusteeCsvFilePath |
    ForEach-Object {

        $CsvCounter++

        Write-Progress @MainProgress -CurrentOperation "CSV file $($CsvCounter) of $($GetMailboxTrusteeCsvFilePath.Count): $($_)"

        $CurrentCsv = @()
        $CurrentCsv += Import-Csv -Path $_

        $CurrentCsvHeaders    = $CurrentCsv |
                                Get-Member -MemberType NoteProperty

        $StandardHeaders |
        ForEach-Object {
            if ($CurrentCsvHeaders.Name -notcontains $_) {

                $WarningMessage =   "CSV file '$($GetMailboxTrusteeCsvFilePath)' is missing one or more mandatory headers.`n`n" +
                                    "See help:`n`n`t" +
                                    '.\Get-MailboxTrusteeWeb.ps1 -?'

                Write-Warning -Message $WarningMessage
                break
            }
        }

        $MailboxTrustees = @()
        $MailboxTrustees += $CurrentCsv
    }

    $MainProgress['Status'] = "Processing imported CSV content.  Total CSV data rows: $($MailboxTrustees.Count)."
    Write-Progress @MainProgress -CurrentOperation "Step 1 of 3: Filtering out ignored mailboxes & trustees"

    $IgnoredPermissionTypes = @()
    $IgnorePermissionType |
    ForEach-Object {
        if ($_ -eq 'AllFolders') {

            $IgnoredPermissionTypes += 'Mailbox root', 'Inbox', 'Calendar', 'Contacts', 'Tasks','Sent Items'
        }
        else {
            $IgnoredPermissionTypes += $_ -replace 'SendAs', 'Send-As' -replace 'SendOnBehalf', 'Send on behalf' -replace 'MailboxRoot','Mailbox root'
        }
    }

    $FilteredMailboxTrustees += $MailboxTrustees |
    Where-Object {
        ($_.TrusteeType -notmatch '(Not found)|(Expanded).*') -and
        ([string]::IsNullOrEmpty($_.TrusteePSmtp) -eq $false) -and
        ($_.TrusteePSmtp -match '^.*\@.*\..*$') -and
        ($IgnoreTrusteePSmtp -notcontains $_.TrusteePSmtp) -and
        ($IgnoreMailboxPSmtp -notcontains $_.MailboxPSmtp) -and
        ($IgnoredPermissionTypes -notcontains $_.PermissionType) -and
        ($IgnoredPermissionTypes -notcontains $_.AccessRights)
    }

    $MainProgress['Status'] = "Processing imported CSV content.  Post-filter CSV data rows: $($FilteredMailboxTrustees.Count)."
    Write-Progress @MainProgress -CurrentOperation "Step 2 of 3: Grouping mailbox-trustee records into unique 1 to 1 relationships"

    $GroupedRelationships = @()
    $GroupedRelationships +=    $FilteredMailboxTrustees |
                                Group-Object -Property  MailboxPSmtp,
                                                        TrusteePSmtp

    $MainProgress['Status'] = "Processing imported CSV content.  Unique 1 to 1 relationships: $($GroupedRelationships.Count)"
    Write-Progress @MainProgress -CurrentOperation "Step 3 of 3: Storing unique relationships for processing"

    $UniqueRelationships =  @()
    $UniqueRelationships += $GroupedRelationships |
                            ForEach-Object {
                                $_.Group |
                                Select-Object -Index 0 |
                                Select-Object -Property MailboxPSmtp,
                                                        TrusteePSmtp
                            }


    if ($PSBoundParameters.ContainsKey('IgnoreMailboxPSmtp')) {

        $MainProgress['Status'] = "Filtering out ignored mailboxes & trustees"
        Write-Progress @MainProgress

        $UniqueRelationships =  $UniqueRelationships |
                                Where-Object {
                                    ($IgnoreTrusteePSmtp -notcontains $_.TrusteePSmtp) -and
                                    ($IgnoreMailboxPSmtp -notcontains $_.MailboxPSmtp)
                                }
    }

    if ($PSBoundParameters.ContainsKey('PermissiveMailboxThreshold')) {

        $MainProgress['Status'] = "Applying thresholds"
        Write-Progress @MainProgress -CurrentOperation "Permissive mailbox threshold: $PermissiveMailboxThreshold"
        Start-Sleep -Seconds 10

        $UniqueRelationships =  $UniqueRelationships |
                                Group-Object -Property MailboxPSmtp |
                                Where-Object {$_.Count -le $PermissiveMailboxThreshold} |
                                Select-Object -ExpandProperty Group
    }

    if ($PSBoundParameters.ContainsKey('PowerTrusteeThreshold')) {

        $MainProgress['Status'] = "Applying thresholds"
        Write-Progress @MainProgress -CurrentOperation "Power Trustee threshold: $PowerTrusteeThreshold"
        Start-Sleep -Seconds 10

        $UniqueRelationships =  $UniqueRelationships |
                                Group-Object -Property TrusteePSmtp |
                                Where-Object {$_.Count -le $PowerTrusteeThreshold} |
                                Select-Object -ExpandProperty Group
    }


    function lookup {
    [CmdletBinding()]
    param(
        [string]$Id,
        [string]$SearchProperty
    )
        $UniqueRelationships |
        Where-Object {$_.$($SearchProperty) -eq $Id}
    }

    $Web        = @()
    $MemberId   = 0

}

process {

    $MainCounter = 0
    $MainProgress['Status'] = 'Working'

    $StartingPSmtp |
    ForEach-Object {

        $MemberId++

        $Web += [PSCustomObject]@{

            Id          = $MemberId
            PSmtp       = $_
            SourceId    = 0
            SourcePSmtp = 'None'
            SourceType  = 'Initial web'
            Depth       = 0
        }
        $StartingWeb = $Web
    }

    $DepthTracker = 0

    $StartingWeb |
    ForEach-Object {

        $CurrentPSmtp   = $null
        $CurrentPSmtp   = $_.PSmtp
        $CurrentId      = $null
        $CurrentId      = $_.Id

        $MainCounter++

        $MainProgress['Status'] = 'Starting recursive relationship lookup'

        Write-Progress @MainProgress -CurrentOperation "Mailbox/Trustee $($MainCounter) of $($StartingPSmtp.Count): $($CurrentPSmtp)"

        $Depth1ForwardLookup =  @()
        $Depth1ReverseLookup =  @()

        if ($Web -notcontains $CurrentPSmtp) {

            lookup -Id $CurrentPSmtp -SearchProperty MailboxPSmtp |
            ForEach-Object {
                $Depth1ForwardLookup += [PSCustomObject]@{
                                            MailboxPSmtp  = $_.MailboxPSmtp
                                            TrusteePSmtp        = $_.TrusteePSmtp
                                            SourceId            = $CurrentId
                                            SourceType          = 'Mailbox'
                                        }
            }

            lookup -Id $CurrentPSmtp -SearchProperty TrusteePSmtp |
            ForEach-Object {
                $Depth1ReverseLookup += [PSCustomObject]@{
                                            MailboxPSmtp  = $_.MailboxPSmtp
                                            TrusteePSmtp        = $_.TrusteePSmtp
                                            SourceId            = $CurrentId
                                            SourceType          = 'Trustee'
                                        }
            }

            if ($Depth1ForwardLookupResults) {$Depth1ForwardLookupResults += $Depth1ForwardLookup}
            else {$Depth1ForwardLookupResults = $Depth1ForwardLookup}

            if ($Depth1ReverseLookupResults) {$Depth1ReverseLookupResults += $Depth1ReverseLookup}
            else {$Depth1ReverseLookupResults = $Depth1ReverseLookup}
        }

        $InnerProgress = @{
            Activity    = 'Performing lookups'
            Id          = 1
            ParentId    = 0
        }

        for (
            $i = 1
            $i -le $MaximumDepth
            $i++
        ) {
            $MainProgress['Status'] = "Mailbox/Trustee $($MainCounter) of $($StartingPSmtp.Count): $($CurrentPSmtp) (Current recursion depth = $($i+1); web size = $($Web.Count))"
            Write-Progress @MainProgress

            New-Variable -Name "Depth$($i+1)ForwardLookup" -Value @() -Force
            New-Variable -Name "Depth$($i+1)ReverseLookup" -Value @() -Force

            $FLItemCounter = 0

            $InnerProgress['Status'] = "Processing level $i's forward lookup results (step 1 of 2 (per level))"

            $CurrentLevelFLItems =  $null
            $CurrentLevelFLItems =  Get-Variable -Name "Depth$($i)ForwardLookup" -ValueOnly
            $CurrentLevelFLItems |
            ForEach-Object {

                $MemberId++

                $FLItemCounter++

                if (($Web.PSmtp -notcontains $_.TrusteePSmtp) -and
                    ($Web.SourcePSmtp -notcontains $_.TrusteePSmtp)) {

                    $InnerProgress['PercentComplete'] = (($FLItemCounter/$CurrentLevelFLItems.Count) * 100)
                    Write-Progress @InnerProgress -CurrentOperation "Forward/reverse lookup $($FLItemCounter) of $($CurrentLevelFLItems.Count): $($_.TrusteePSmtp)"

                    $CurrentFLItem = $_

                    lookup -Id $_.TrusteePSmtp -SearchProperty MailboxPSmtp |
                    Where-Object {
                        ($Web.PSmtp -notcontains $_.TrusteePSmtp) -and
                        ($Web.SourcePSmtp -notcontains $_.TrusteePSmtp)
                    }|
                    ForEach-Object {
                        (Get-Variable -Name "Depth$($i+1)ForwardLookup").Value +=

                            [PSCustomObject]@{
                                MailboxPSmtp  = $_.MailboxPSmtp
                                TrusteePSmtp        = $_.TrusteePSmtp
                                SourceId            = $MemberId
                                SourceType          = 'Mailbox'
                            }
                    }

                    lookup -Id $_.TrusteePSmtp -SearchProperty TrusteePSmtp |
                    Where-Object {
                        ($Web.PSmtp -notcontains $_.MailboxPSmtp) -and
                        ($Web.SourcePSmtp -notcontains $_.MailboxPSmtp) -and
                        ($_.MailboxPSmtp -ne $CurrentFLItem.MailboxPSmtp)
                    } |
                    ForEach-Object {

                        (Get-Variable -Name "Depth$($i+1)ReverseLookup").Value +=

                            [PSCustomObject]@{
                                MailboxPSmtp  = $_.MailboxPSmtp
                                TrusteePSmtp        = $_.TrusteePSmtp
                                SourceId            = $MemberId
                                SourceType          = 'Trustee'
                            }
                    }

                    $Web += [PSCustomObject]@{

                        Id          = $MemberId
                        PSmtp       = $_.TrusteePSmtp
                        SourceId    = $_.SourceId
                        SourcePSmtp = $_.MailboxPSmtp
                        SourceType  = $_.SourceType
                        Depth       = $i
                    }
                }
            }

            $RLItemCounter = 0

            $InnerProgress['Status'] = "Processing level $i's reverse lookup results (step 2 of 2 per level)"

            $CurrentLevelRLItems = $null
            $CurrentLevelRLItems = Get-Variable -Name "Depth$($i)ReverseLookup" -ValueOnly
            $CurrentLevelRLItems |
            ForEach-Object {

                $MemberId++

                $RLItemCounter++

                if (($Web.PSmtp -notcontains $_.MailboxPSmtp) -and
                    ($Web.SourcePSmtp -notcontains $_.MailboxPSmtp)) {

                    $InnerProgress['PercentComplete'] = (($RLItemCounter/$CurrentLevelRLItems.Count) * 100)
                    Write-Progress @InnerProgress -CurrentOperation "Forward/reverse lookup $($RLItemCounter) of $($CurrentLevelRLItems.Count): $($_.MailboxPSmtp)"

                    $CurrentRLItem = $_

                    lookup -Id $_.MailboxPSmtp -SearchProperty MailboxPSmtp |
                    Where-Object {
                        ($Web.PSmtp -notcontains $_.TrusteePSmtp) -and
                        ($Web.SourcePSmtp -notcontains $_.TrusteePSmtp) -and
                        ($_.TrusteePSmtp -ne $CurrentRLItem.TrusteePSmtp)
                    } |
                    ForEach-Object {

                        (Get-Variable -Name "Depth$($i+1)ForwardLookup").Value +=

                            [PSCustomObject]@{
                                MailboxPSmtp  = $_.MailboxPSmtp
                                TrusteePSmtp        = $_.TrusteePSmtp
                                SourceId            = $MemberId
                                SourceType          = 'Mailbox'
                            }
                    }

                    lookup -Id $_.MailboxPSmtp -SearchProperty TrusteePSmtp |
                    Where-Object {
                        ($Web.PSmtp -notcontains $_.MailboxPSmtp) -and
                        ($Web.SourcePSmtp -notcontains $_.MailboxPSmtp) -and
                        ($_.TrusteePSmtp -ne $CurrentRLItem.TrusteePSmtp)
                    } |
                    ForEach-Object {

                        (Get-Variable -Name "Depth$($i+1)ReverseLookup").Value +=

                            [PSCustomObject]@{
                                MailboxPSmtp  = $_.MailboxPSmtp
                                TrusteePSmtp        = $_.TrusteePSmtp
                                SourceId            = $MemberId
                                SourceType          = 'Trustee'
                            }
                    }

                    $Web += [PSCustomObject]@{

                        Id          = $MemberId
                        PSmtp       = $_.MailboxPSmtp
                        SourceId    = $_.SourceId
                        SourcePSmtp = $_.TrusteePSmtp
                        SourceType  = $_.SourceType
                        Depth       = $i
                    }
                }
            }

            if (
                ((Get-Variable -Name "Depth$($i+1)ForwardLookup").Value.Count -eq 0) -and
                ((Get-Variable -Name "Depth$($i+1)ReverseLookup").Value.Count -eq 0) ) {

                if ($i -gt $DepthTracker) {$DepthTracker = $i}
                break
            }
            else {
                if (Get-Variable -Name "Depth$($i+1)ForwardLookupResults" -ErrorAction:SilentlyContinue) {

                    (Get-Variable -Name "Depth$($i+1)ForwardLookupResults").Value +=

                        Get-Variable -Name "Depth$($i+1)ForwardLookup" -ValueOnly
                }
                else {
                    New-Variable -Name "Depth$($i+1)ForwardLookupResults" -Value (Get-Variable -Name "Depth$($i+1)ForwardLookup" -ValueOnly) -Force
                }

                if (Get-Variable -Name "Depth$($i+1)ReverseLookupResults" -ErrorAction:SilentlyContinue) {

                    (Get-Variable -Name "Depth$($i+1)ReverseLookupResults").Value +=

                        Get-Variable -Name "Depth$($i+1)ReverseLookup" -ValueOnly
                }
                    else {
                    New-Variable -Name "Depth$($i+1)ReverseLookupResults" -Value (Get-Variable -Name "Depth$($i+1)ReverseLookup" -ValueOnly) -Force
                }
            }
        } # end for ()
    } # end $StartingPSmtp | ForEach-Object {}

} # end process {}

end {

    Write-Progress @InnerProgress -Completed
    Write-Progress @MainProgress -Completed

    $Web

    if([Environment]::GetCommandLineArgs() -notmatch '-noni*') {

        $EndTime = Get-Date

        $ScriptRuntimeDetails = [Ordered]@{

            'Final web size'                = $Web.Count
            'Starting web size'             = $StartingPSmtp.Count
            'Start time'                    = $StartTime.ToLongTimeString()
            'End time'                      = $EndTime.ToLongTimeString()
            'Duration'                      = $EndTime-$StartTime -replace '\..*',''
            'Depth reached'                 = $DepthTracker
            'Maximum depth'                 = $MaximumDepth
            'Permissive mailbox threshold'  = $PermissiveMailboxThreshold
            'Power trustee threshold'       = $PowerTrusteeThreshold
            'Ignored mailboxes'             = $IgnoreMailboxPSmtp.Count
            'Ignored trustees'              = $IgnoreTrusteePSmtp.Count
            'Ignored permission types'      = $IgnorePermissionType -join ','
        }

        if ($PSBoundParameters.ContainsKey('OutVariable')) {

            $ScriptRuntimeDetails.Add('OutVariable', "`$$($PSBoundParameters.OutVariable)")
        }

        if ($PSBoundParameters.ContainsKey('InformationVariable')) {

            $ScriptRuntimeDetails.Add('InformationVariable', "`$$($PSBoundParameters.InformationVariable)")
        }

        $DepthTallies = Get-Variable -Name Depth*LookupResults
        $DepthTallies |
        ForEach-Object {

            $ScriptRuntimeDetails.Add(

                "$(($_.Name -replace '((Depth)|([0-9]+))','$1 ' -replace 'Lookup',' Lookup ').ToLower() -replace 'depth','Depth')",
                $_.Value.Count
            )
        }
        $LongestKey = ($ScriptRuntimeDetails.Keys | Measure-Object -Property Length -Maximum).Maximum
        $AfterActionReport = "`n`tGet-MailboxTrusteeWebSQLEdition.ps1 - After Action Report`n"
        $ScriptRuntimeDetails.Keys |
        ForEach-Object {

            $AfterActionReport +=   "`n`t$($_[0..40] -join '')" +
                                    "." * (($LongestKey+4)-($_[0..40].Length)) +
                                    ": $($ScriptRunTimeDetails.$($_))"
        }

        Write-Information -MessageData $AfterActionReport -InformationAction:Continue
        Write-Information -MessageData "`nCommand:`n`n$($PSCmdlet.MyInvocation.Line)" -InformationAction:Continue
    }

}

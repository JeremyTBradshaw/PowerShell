<#

    .Synopsis

    This script mirrors the portion of Get-MailboxTrusteeWeb.ps1 - the part which
    groups the data into unique 1 to 1 relationships, after which the relationships
    are sent through the recursive lookup process.  Since this part of the script
    can be very time consuming, this script can be used to do the same process, but
    save the optimzed input data to file so that it can be processed later by
    Get-MailboxTrusteeWeb.ps1.

    Having the optimzed input data on file is very time-saving when running
    Get-MailboxTrusteeWeb.ps1 multiple times (e.g. for multiple batches), since it
    removes the need to re-pre-rocess the base data every time.


    .Description

    This is the 2nd of 2 brother scripts to Get-MailboxTrustee.ps1.  Therefore it
    requires that the passed CSV file(s) contain(s) a few headers for matching
    properties output by Get-MailboxTrustee.ps1.  Any other headers (i.e. columns)
    are ignored.

    At this time, the output from the MinimizeOutput mode of Get-MailboxTrustee.ps1
    is not supported.  Instead only the standard mode's output is accepted.

    Mandatory CSV headers:

    - PrimarySmtpAddress
    - RecipientTypeDetails*
    - PermissionType
    - AccessRights
    - TrusteePSmtp
    - TrusteeType*

    * = future implementions will harness these.


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

    $AllMailboxes = Get-Mailboxes -ResultSize:Unlimited

    PS C:\> $AllMailboxes | .\Get-MailboxTrustee.ps1 -ExpandTrusteeGroups | Export-Csv .\All-Mailbox-Trustees.csv -NTI

    PS C:\> $Optimzed = .\Optimize-MailboxTrusteeWebInput.ps1 `
                            -GetMailboxTrusteeCsvFilePath .\All-Mailboxes-Trustees.csv `
                            -PermissiveMailboxThreshold 500 `
                            -PowerTrusteeThreshold 500

    PS C:\> $Optimized | Export-Csv .\All-Mailbox-Trustees-Optimized.csv -NTI

    PS C:\> $MigrationBatch1Users = Import-Csv .\MigrationBatch1Users.csv

    PS C:\> $MigBatch1UsersWeb = .\Optimize-MailboxTrusteeWebInut.ps1 `
                                    -GetMailboxTrusteeCsvFilePath .\All-Mailbox-Trustees-Optimized.csv `
                                    -OptimizedInput

    PS C:\> $MigBatch1UsersWeb | Export-Csv .\MigBatchUsers1_MailboxTrusteeWeb.csv -NTI


    .Link

    https://github.com/JeremyTBradshaw/PowerShell/blob/master/Optimize-MailboxTrusteeWebInput.ps1
    # ^ Optimize-MailboxTrusteeWebInput.ps1


    .Link

    # Get-MailboxTrustee.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/master/Get-MailboxTrustee.ps1

    # Get-MailboxTrusteeWeb.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/master/Get-MailboxTrusteeWeb.ps1

    # Get-MailboxTrusteeWebSQLEdition.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/master/Get-MailboxTrusteeWebSQLEdition.ps1

    # New-MailboxTrusteeReverseLookup.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/master/New-MailboxTrusteeReverseLookup.ps1

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

    [ValidateRange(1,[int32]::MaxValue)]
    [int]$PermissiveMailboxThreshold = 500,

    [ValidateRange(1,[int32]::MaxValue)]
    [int]$PowerTrusteeThreshold = 500,

    [ValidateScript({
        $_ | ForEach-Object {

            if ($_.Length -gt 320) {throw 'Must be 320 characters or less (maximum for SMTP address)'}
            elseif ($_ -notmatch '^.*\@.*\..*$') {throw "'$($_)' is not a valid SMTP address."}
            else {$true}
        }
    })]
    [string[]]$IgnoreMailboxPSmtp,

    [ValidateScript({
        $_ | ForEach-Object {

            if ($_.Length -gt 320) {throw 'Must be 320 characters or less (maximum for SMTP address)'}
            elseif ($_ -notmatch '^.*\@.*\..*$') {throw "'$($_)' is not a valid SMTP address."}
            else {$true}
        }
    })]
    [string[]]$IgnoreTrusteePSmtp,

    [ValidateSet(
        'FullAccess', 'SendAs', 'SendOnBehalf',
        'AllFolders', 'MailboxRoot', 'Inbox', 'Calendar', 'Contacts', 'Tasks', 'SentItems'
    )]
    [string[]]$IgnorePermissionType

)

begin {

    $StartTime = Get-Date

    $MainProgress = @{

        Activity            = "Optimize-MailboxTrusteeWebInput.ps1 (Start time: $($StartTime.DateTime))"
        Id                  = 0
        ParentId            = -1
        Status              = 'Initializing'
        PercentComplete     = -1
        SecondsRemaining    = -1
    }

    Write-Progress @MainProgress

    $StandardHeaders = @(

        'PrimarySmtpAddress',
        'RecipientTypeDetails',
        'PermissionType',
        'AccessRights',
        'TrusteePSmtp',
        'TrusteeType'
    )

    $MailboxTrustees = @()

    $CsvCounter = 0

    $MainProgress['Status'] = 'Importing CSV file(s)'

    $GetMailboxTrusteeCsvFilePath |
    ForEach-Object {

        $CsvCounter++

        Write-Progress @MainProgress -CurrentOperation "CSV file $($CsvCounter) of $($GetMailboxTrusteeCsvFilePath.Count): $($_)"

        $CurrentCsv =   @()
        $CurrentCsv +=  Import-Csv -Path $_

        $CurrentCsvHeaders    = $CurrentCsv |
                                Get-Member -MemberType NoteProperty

        $StandardHeaders |
        ForEach-Object {
            if ($CurrentCsvHeaders.Name -notcontains $_) {

                $WarningMessage =   "CSV file '$($GetMailboxTrusteeCsvFilePath)' is missing one or more mandatory headers.`n`n" +
                                    "See help:`n`n`t" +
                                    '.\Optimize-MailboxTrusteeWebInput.ps1 -?'

                Write-Warning -Message $WarningMessage
                break
            }
        }

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
                                    ($IgnoreTrusteePSmtp -notcontains $_.TrusteePSmtp) -and
                                    ($IgnoreMailboxPSmtp -notcontains $_.PrimarySmtpAddress) -and
                                    ($IgnoredPermissionTypes -notcontains $_.PermissionType) -and
                                    ($IgnoredPermissionTypes -notcontains $_.AccessRights)
                                }

    $MainProgress['Status'] = "Processing imported CSV content.  Post-filter CSV data rows: $($MailboxTrustees.Count)."
    Write-Progress @MainProgress -CurrentOperation "Step 2 of 3: Grouping mailbox-trustee records into unique 1 to 1 relationships"

    $GroupedRelationships = $MailboxTrustees |
                            Group-Object -Property  PrimarySmtpAddress,
                                                    TrusteePSmtp

    $MainProgress['Status'] = "Processing imported CSV content.  Unique 1 to 1 relationships: $($GroupedRelationships.Count)"
    Write-Progress @MainProgress -CurrentOperation "Step 3 of 3: Storing unique relationships for processing"

    $UniqueRelationships =  $GroupedRelationships |
                            ForEach-Object {
                                $_.Group |
                                Select-Object -Index 0 |
                                Select-Object -Property PrimarySmtpAddress,
                                                        TrusteePSmtp,
                                                        @{  Name = 'RecipientTypeDetails'
                                                            Expression = '0'},
                                                        @{  Name = 'PermissionType'
                                                            Expression = '0'},
                                                        @{  Name = 'AccessRights'
                                                            Expression = '0'},
                                                        @{  Name = 'TrusteeType'
                                                            Expression = '0'}
                            }

    $MainProgress['Status'] = "Applying thresholds"
    Write-Progress @MainProgress -CurrentOperation "Permissive mailbox threshold: $PermissiveMailboxThreshold"

    $UniqueRelationships =  $UniqueRelationships |
                            Group-Object -Property PrimarySmtpAddress |
                            Where-Object {$_.Count -le $PermissiveMailboxThreshold} |
                            Select-Object -ExpandProperty Group

    Write-Progress @MainProgress -CurrentOperation "Power Trustee threshold: $PowerTrusteeThreshold"

    $UniqueRelationships =  $UniqueRelationships |
                            Group-Object -Property TrusteePSmtp |
                            Where-Object {$_.Count -le $PowerTrusteeThreshold} |
                            Select-Object -ExpandProperty Group

    Write-Debug "About to leave process {*}"

} # end process {}

end {

    Write-Output $UniqueRelationships # a.k.a. 'Optimized'

    $MainProgress['Status'] = 'Completed'
    Write-Progress @MainProgress -Completed

}

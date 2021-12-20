<#
    .Synopsis
    ***SQL Edition *** - Faster processing than CSV/PowerShell alone.

    Identify webs of mailbox-trustee relationships, in the following manner:

    1.) Reverse Lookup: Find all mailboxes that one or more trustee users have access to.

    2.) Forward Lookup: For these same users, find all other users who have access to their mailbox.

    3.) Recursively repeat this process against all found mailboxes and trustees until the entire web is known (or set
        a maximum recursion depth).

    .Description
    ***SQL Edition*** Important Note:

    This has only been tested on Windows 10 1803/1809 with SQL Express 2017 ('Basic' installation), where current user
    has access to create tables, import/query/delete data, then drop tables.  At this time, alternate credentials for
    SQL are not accommodated, so current user must have the necessary permissions.

    This is the 2nd of 2 brother scripts to Get-MailboxTrustee.ps1.  Therefore it requires that the passed CSV file(s)
    contain(s) a few headers for matching properties output by Get-MailboxTrustee.ps1.  Any other headers (i.e.
    columns) are ignored.

    Mandatory CSV headers:

    - MailboxPSmtp
    - MailboxType
    - MailboxGuid
    - PermissionType
    - AccessRights
    - TrusteePSmtp
    - TrusteeType
    - TrusteeGuid

    .Parameter GetMailboxTrusteeCsvFilePath
    The full or direct path to the CSV file(s) containing objects that have been output from Get-MailboxTrustee.ps1.

    Since it is common to have run Get-MailboxTrustee.ps1 in multiple passes,
    multiple CSV files can be specified here, to spare the need to combine CSV
    files manually.

    .Parameter StartingPSmtp
    Specify one or more mailboxes and/or trustees to search recursively as described in the synopsis.  Include all SMTP
    addresses that are intended to be in the same web (think 'web' = 'migration batch').

    The reason PrimarySmtpAddress is used for the starting ID property is that this script's original intended purpose
    is to determine mailbox dependencies for migration batch planning.  As such, if a user doesn't have a mailbox of
    their own (i.e. no email address), they will not be impacted by being left out of this exercise, since they
    themselves will not need to be migrated.

    .Parameter SqlServerInstance
    Use the format <ServerName>\<SQLInstance>.  Default if not specified is "localhost\SQLEXPRESS".

    .Parameter SqlDatabase
    Specify another writeable database name if TempDB (the default) is not desired.  Tables and CSV data will be
    created and imported into this database, then dropped/deleted at the end of the script.

    ConfirmImpact is set to 'High' to ensure the user is prompted for confirmation before dropping (if necessary) and
    creating the tables and importing the CSV file(s), overwriting any previously imported data.

    .Parameter SqlNoImportOrPreProccessing
    This switch can be used on subsequent runs to bypass importing the CSV file(s) into SQL again.  In order to enable
    this functionality, first run the script once with the -SqlNoCleanupAfter switch.  Alternatively, respond with 'n'
    for 'No' when prompted for confirmation to perform the cleanup at the end of the script.

    This allows for very quick trial and error using the various filtering (ignore) and threshold parameters to
    fine-tune webs to meet your specific goals.

    .Parameter SqlNoCleanupAfter
    Specifies to leave the tables 'GetMailboxTrusteeWeb1' and 'GetMailboxTrusteeWeb2' in place rather than dropping
    them which is the default behavior.

    **Note: when no cleanup is done, the two tables that were created is left fully in tact for future runs of the
    script (with the -SqlNoImportOrPreProcessing switch).  However, the second table is always recreated when running
    the script, since it is built using the supplied parameters to narrow down the web-identifying process to just the
    intended permission relationships.

    ConfirmImpact is set to 'High' to ensure the user is prompted for confirmation before deleting the tables.

    .Parameter PermissiveMailboxThreshold
    Some mailboxes have very many trustees.  Examples are popular room and equipment mailboxes.  This can plague the
    process of identifying webs that are truly significant.  Use this parameter to specify the maximum # of 1 to 1
    relationships mailboxes can have before they will be excluded from the web.

    - Default: 500
    - Minimum: 1
    - Maximum: Positive [Int32]

    .Parameter PowerTrusteeThreshold
    Some users have access to very many mailboxes.  Examples are administrators and Power Users.  Use this parameter
    to specify the maximum # of 1 to 1 relationships trustees can have before they will be excluded from the web.

    - Default: 500
    - Minimum: 1
    - Maximum: Positive [Int32]

    .Parameter MaximumDepth
    How many times to recurse when performing the forward and reverse trustee lookups.  This can help if the full web
    of permission relationships is a larger group of mailboxes than desired for the task at hand, for example, a
    migration batch.

    - Default: 100
    - Minimum: 1
    - Maximum: Positive [Int32]

    .Parameter IgnoreMailboxPSmtp
    One or more mailboxes to ommit from searches and output.  Expected input is PrimarySmtpAddress (of the mailbox
    (i.e. trust'ING user)), comma-separated or an array variable if supplying multiple.

    .Parameter IgnoreTrusteePSmtp
    One or more trustee users to ommit from searches and output. Expected input is PrimarySmtpAddress (of the trustee
    (i.e. trust'ED user), comma-separated or an array variable if supplying multiple.

    .Parameter IgnorePermissionType
    One or more of the following permissions types to ommit from searches and output:

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


    .Parameter IgnoreMailboxType
    One or more of the following permissions types to ommit from searches and output:

    - UserMailbox
    - AllResources (RoomMailbox & EquipmentMailbox)
    - RoomMailbox
    - EquipmentMailbox
    - SharedMailbox
    - RemoteMailboxes (Remote*Mailbox)
    - RemoteUserMailbox
    - RemoteRoomMailbox
    - RemoteEquipmentMailbox
    - RemoteSharedMailbox

    .Example
    $AllMailboxes = Get-Mailboxes -ResultSize:Unlimited

    PS C:\> $AllMailboxes | .\Get-MailboxTrustee.ps1 -ExpandTrusteeGroups | Export-Csv .\All-Mailbox-Trustees.csv -NTI

    PS C:\> $MigrationBatch1ProposedList = Import-Csv .\MigrationBatch1Users.csv

    PS C:\> $MigBatch1UsersWeb = .\GetMailboxTrusteeWebSQLEdition.ps1 `
                                    -GetMailboxTrusteeCsvFilePath .\All-Mailbox-Trustees.csv `
                                    -StartingPSmtp $MigrationBatch1ProposedList.PrimarySmtpAddress `

    PS C:\> $MigBatch1UsersWeb | Export-Csv .\MigBatchUsers1_MailboxTrusteeWeb.csv -NTI

    .Example
    .\Get-MailboxTrusteeWebSQLEdition -GetMailboxTrusteeCsvFilePath .\AllMbxTrustees.csv `
                                        -StartingPSmtp SalesUser1@contoso.com, SalesUser2@contoso.com `
                                        -SqlNoImportOrPreProcessing `
                                        -PermissiveMailboxThreshold 50 `
                                        -PowerTrusteeThreshold 20 `
                                        -IgnoreMailboxPSmtp AllUsersVacationCalendar@contoso.com `
                                        -IgnoreTrusteePSmtp BusyAdmin@contoso.com `
                                        -OutVariable Web1

    .Example
    .\Get-MailboxTrusteeWebSQLEdition -SqlNoImportOrPreProcessing -SqlNoCleanupAfter:$false
    # ^ this is how to cleanup afterwards gracefully.

    .Example
    cls; [void](.\Get-MailboxTrusteeWebSQLSQLEdition.ps1 -SqlDatabase MigPlanning `
                                                            -SqlNoImportOrPreProcessing `
                                                            -ov o -iv i `
                                                            -StartingPSmtp $ResourceMailboxPSmtps `
                                                            -PermissiveMailboxThreshold 20 `
                                                            -PowerTrusteeThreshold 5)

    PS C:\> $o | ft -auto   # <--: View the web in tabular form.
    PS C:\> $i              # <--: view the After Action Report.

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrusteeWebSQLEdition.ps1
    # ^ Get-MailboxTrusteeWebSQLEdition.ps1

    .Link
    # Get-MailboxTrustee.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrustee.ps1

    # Optimize-MailboxTrusteeWebInput.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Optimize-MailboxTrusteeWebInput.ps1

    # Get-MailboxTrusteeWeb.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxTrusteeWeb.ps1

    # New-MailboxTrusteeReverseLookup.ps1
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/New-MailboxTrusteeReverseLookup.ps1
#>
#Requires -Version 5.1
#Requires -Module SqlServer

[CmdletBinding(
    SupportsShouldProcess = $true,
    ConfirmImpact = 'High')]

param(

    [Parameter(Mandatory = $true)]
    [ValidateScript({
            $_ | ForEach-Object {

                if ($_.Length -gt 320) { throw 'Must be 320 characters or less (maximum for SMTP address)' }
                elseif ($_ -notmatch '^.*\@.*\..*$') { throw "'$($_)' is not a valid SMTP address." }
                else { $true }
            }
        })]
    [string[]]$StartingPSmtp,

    [ValidateRange(1, [int32]::MaxValue)]
    [int]$PermissiveMailboxThreshold = 500,

    [ValidateRange(1, [int32]::MaxValue)]
    [int]$PowerTrusteeThreshold = 500,

    [ValidateRange(1, [int32]::MaxValue)]
    [int]$MaximumDepth = 100,

    [ValidateScript({
            $_ | ForEach-Object {

                if ($_.Length -gt 320) { throw 'Must be 320 characters or less (maximum for SMTP address)' }
                elseif ($_ -notmatch '^.*\@.*\..*$') { throw "'$($_)' is not a valid SMTP address." }
                else { $true }
            }
        })]
    [string[]]$IgnoreMailboxPSmtp,

    [ValidateScript({
            $_ | ForEach-Object {

                if ($_.Length -gt 320) { throw 'Must be 320 characters or less (maximum for SMTP address)' }
                elseif ($_ -notmatch '^.*\@.*\..*$') { throw "'$($_)' is not a valid SMTP address." }
                else { $true }
            }
        })]
    [string[]]$IgnoreTrusteePSmtp,

    [ValidateSet(
        'FullAccess', 'SendAs', 'SendOnBehalf',
        'AllFolders', 'MailboxRoot', 'Inbox', 'Calendar', 'Contacts', 'Tasks', 'SentItems'
    )]
    [string[]]$IgnorePermissionType,

    [Parameter(
        HelpMessage = 'AllResources = Room and Equipment mailboxes; RemoteMailboxes = Remote*Mailbox'
    )]
    [ValidateSet(
        'UserMailbox', 'SharedMailbox',
        'AllResources', 'RoomMailbox', 'EquipmentMailbox',
        'RemoteMailboxes', 'RemoteUserMailbox', 'RemoteRoomMailbox', 'RemoteEquipmentMailbox', 'RemoteSharedMailbox'
    )]
    [string[]]$IgnoreMailboxType,

    [string]$SqlServerInstance = 'localhost\SQLExpress',
    [string]$SqlDatabase = 'TempDB',
    [switch]$SqlNoImportOrPreProcessing,
    [switch]$SqlNoCleanupAfter = $SqlNoImportOrPreProcessing

)

DynamicParam {

    $pdGetMailboxTrusteeCsvFilePath = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary

    $Attributes = New-Object System.Management.Automation.ParameterAttribute
    $Attributes.Mandatory = $true

    $AttributeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
    $AttributeCollection.Add($Attributes)

    if ($SqlNoImportOrPreProcessing -eq $false) {

        $GetMailboxTrusteeCsvFilePath = New-Object -Type System.Management.Automation.RuntimeDefinedParameter(

            'GetMailboxTrusteeCsvFilePath',
            [String[]],
            $AttributeCollection
        )

        $pdGetMailboxTrusteeCsvFilePath.Add('GetMailboxTrusteeCsvFilePath', $GetMailboxTrusteeCsvFilePath)
        $pdGetMailboxTrusteeCsvFilePath
    }
}

begin {
    $StartTime = Get-Date

    $MainProgress = @{

        Activity         = "Get-MailboxTrusteeWeb.ps1 ***SQL Edition*** (Start time: $($StartTime.DateTime))"
        Id               = 0
        ParentId         = -1
        Status           = 'Initializing'
        PercentComplete  = -1
        SecondsRemaining = -1
    }

    Write-Progress @MainProgress

    $InvokeSqlCmdParams = @{

        ServerInstance = "$($SqlServerInstance)"
        ErrorAction    = 'Stop'
    }

    $SqlLookupPSmtpParams = $InvokeSqlCmdParams
    $SqlLookupPSmtpParams['ErrorAction'] = 'Continue'

    $SqlTable1 = 'GetMailboxTrusteeWeb1'
    $SqlTable2 = 'GetMailboxTrusteeWeb2'
    $SqlTableCheck1 = "SELECT * FROM [$($SqlDatabase)].INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '$($SqlTable1)'"
    $SqlTableCheck2 = "SELECT * FROM [$($SqlDatabase)].INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '$($SqlTable2)'"

    $SqlCreateTable2 = "USE [$($SqlDatabase)]`n" +
    "CREATE TABLE $($SqlTable2) (`n" +
    "Id int IDENTITY(1,1) PRIMARY KEY,`n" +
    "MailboxPSmtp varchar(320),`n" +
    "MailboxType varchar(30),`n" +
    "TrusteePSmtp varchar(320),`n" +
    "TrusteeType varchar(30)`n" +
    ")"

    $StandardHeaders = @(

        'MailboxPSmtp',
        'MailboxType',
        'MailboxGuid',
        'PermissionType',
        'AccessRights',
        'TrusteePSmtp',
        'TrusteeType',
        'TrusteeGuid'
    )

    $CsvCounter = 0

    $MainProgress['Status'] = 'Importing CSV file(s)'

    $MailboxTrustees = @()

    $break = $false

    switch ($SqlNoImportOrPreProcessing) {

        $false {
            if ($PSCmdlet.ShouldProcess(
                    "SQL Server instance: $($SqlServerInstance); Database: [$($SqlDatabase)]",
                    'Import and overwrite (create (or drop then re-create) tables and import CSV data to SQL)?')) {

                $GetMailboxTrusteeCsvFilePath.Value |
                ForEach-Object {

                    if ((Test-Path -Path $_) -eq $false) { throw "Can't find file '$($_)'." }
                    if ($_ -notmatch '(\.csv$)') { throw "Only .csv files are accepted." }

                    $CsvCounter++

                    Write-Progress @MainProgress -CurrentOperation "CSV file $($CsvCounter) of $($GetMailboxTrusteeCsvFilePath.Value.Count): $($_)"

                    $CurrentCsv = @()
                    $CurrentCsv += Import-Csv -Path $_

                    $CurrentCsvHeaders = $CurrentCsv |
                    Get-Member -MemberType NoteProperty

                    $StandardHeaders |
                    ForEach-Object {
                        if ($CurrentCsvHeaders.Name -notcontains $_) {

                            $Script:warningMessage = "CSV file '$($GetMailboxTrusteeCsvFilePath.Value[$CurrentCsv-1])' is missing one or more mandatory headers.`n`n" +
                            "See help:`n`n`t" +
                            '.\Get-MailboxTrusteeWebSQLEdition.ps1 -?'

                            $Script:break = $true
                            break
                        }
                    }

                    $MailboxTrustees += $CurrentCsv
                }

                try {
                    $MainProgress['Status'] = 'Preparing SQL / importing CSV data'
                    Write-Progress @MainProgress -CurrentOperation "Creating temporary tables"

                    $SqlCreateTable1 = "USE [$($SqlDatabase)]`n" +
                    "CREATE TABLE $($SqlTable1) (`n" +
                    "Id int IDENTITY(1,1) PRIMARY KEY,`n" +
                    "MailboxPSmtp varchar(320),`n" +
                    "MailboxGuid varchar(36),`n" +
                    "MailboxType varchar(30),`n" +
                    "PermissionType varchar(15),`n" +
                    "AccessRights varchar(1000),`n" +
                    "TrusteePSmtp varchar(320),`n" +
                    "TrusteeGuid varchar(36),`n" +
                    "TrusteeType varchar(30)`n" +
                    ")"

                    if (Invoke-Sqlcmd -Query $SqlTableCheck1 @InvokeSqlCmdParams) {

                        Write-Progress @MainProgress -CurrentOperation "Dropping pre-existing table $($SqlTable1)"

                        Invoke-Sqlcmd -Query "USE [$($SqlDatabase)] DROP TABLE $($SqlTable1)" @InvokeSqlCmdParams
                        Start-Sleep -Milliseconds 250
                    }

                    Write-Progress @MainProgress -CurrentOperation "Creating table $($SqlTable1)"

                    Invoke-Sqlcmd -Query $SqlCreateTable1 @InvokeSqlCmdParams
                    Start-Sleep -Milliseconds 250

                    if (Invoke-Sqlcmd -Query $SqlTableCheck2 @InvokeSqlCmdParams) {

                        Write-Progress @MainProgress -CurrentOperation "Dropping pre-existing table $($SqlTable2)"

                        Invoke-Sqlcmd -Query "USE [$($SqlDatabase)] DROP TABLE $($SqlTable2)" @InvokeSqlCmdParams
                        Start-Sleep -Milliseconds 250
                    }

                    Write-Progress @MainProgress -CurrentOperation "Creating table $($SqlTable2)"

                    Invoke-Sqlcmd -Query $SqlCreateTable2 @InvokeSqlCmdParams
                    Start-Sleep -Milliseconds 250


                    $MainProgress['Status'] = "Processing imported CSV content.  Total CSV data rows: $($MailboxTrustees.Count)."

                    $SqlImportCounter = 0

                    Write-Debug -Message "Stop here to manually perform a BULK INSERT"
                    <#
                        # BULK Insert (storing code here for manual use in -Debug mode):

                        $SqlBulkInsertCsv = "USE [$($SqlDatabase)]`n" +
                        "BULK INSERT $($SqlTable1)`n" +
                        "FROM 'C:\Users\bradshaj\Desktop\NSHA-Exports\.Refreshes\2021-12-09\.processed\MBXTrusteesTMP.csv'`n" +
                        "WITH (FIRSTROW = 2,FIELDTERMINATOR = ',' , ROWTERMINATOR = '\n')"
                        Invoke-Sqlcmd -Query $SqlBulkInsertCsv @InvokeSqlCmdParams
                    #>

                    $MailboxTrustees |
                    Where-Object {
                        ($_.TrusteeType -notmatch '(Not found)|(Expanded).*') -and
                        ([string]::IsNullOrEmpty($_.TrusteePSmtp) -eq $false) -and
                        ($_.TrusteePSmtp -match '^.*\@.*\..*$')
                    } |
                    ForEach-Object {

                        $SqlImportCounter++
                        $MainProgress['PercentComplete'] = (($SqlImportCounter / $MailboxTrustees.Count) * 100)
                        Write-Progress @MainProgress -CurrentOperation "Step 1 of 2: Importing into SQL"

                        $SqlInsertCsvItem = "USE [$($SqlDatabase)]`n" +
                        "INSERT INTO $($SqlTable1) (`n" +
                        "MailboxPSmtp,MailboxType,MailboxGuid,`n" +
                        "PermissionType,AccessRights,`n" +
                        "TrusteePSmtp,TrusteeType,TrusteeGuid)`n" +
                        "VALUES (`n" +
                        "'$($_.MailboxPSmtp -replace ""'"",""''"")','$($_.MailboxType)','$($_.MailboxGuid)',`n" +
                        "'$($_.PermissionType)','$($_.AccessRights)',`n" +
                        "'$($_.TrusteePSmtp -replace ""'"",""''"")','$($_.TrusteeType)','$($_.TrusteeGuid)')"


                        Invoke-Sqlcmd -Query $SqlInsertCsvItem @InvokeSqlCmdParams
                    }

                    $MainProgress['PercentComplete'] = -1
                }
                catch {
                    Write-Warning -Message "A failure occurred while attempting to prepare SQL / import CSV data'.`n`nError:`n"
                    throw $_
                }
            } # end if ($PSCmdlet.ShouldProcess()) {}

            else {
                $WarningMessage = "User canceled CSV -> SQL import."
                $break = $true
                break
            }
        } # end $false

        $true {
            try {
                if (-not (Invoke-Sqlcmd -Query $SqlTableCheck1 @InvokeSqlCmdParams)) {

                    $WarningMessage = "SQL table '$($SqlTable1)' not found.  Exiting script."
                    $break = $true
                    break
                }

                if (Invoke-Sqlcmd -Query $SqlTableCheck2 @InvokeSqlCmdParams) {

                    Write-Progress @MainProgress -CurrentOperation "Dropping pre-existing table $($SqlTable2)"

                    Invoke-Sqlcmd -Query "USE [$($SqlDatabase)] DROP TABLE $($SqlTable2)" @InvokeSqlCmdParams
                    Start-Sleep -Milliseconds 250
                }

                Write-Progress @MainProgress -CurrentOperation "Creating table $($SqlTable2)"

                Invoke-Sqlcmd -Query $SqlCreateTable2 @InvokeSqlCmdParams
                Start-Sleep -Milliseconds 250
            }
            catch {
                Write-Warning -Message "A failure occurred while attempting to prepare SQL ((dropping/re-)creating table '$($SqlTable2)'.`n`nError:`n"
                throw $_
            }
        }# end $true
    } # end switch ($SqlNoImportOrPreprocessing)

    if ($break -eq $true) {

        Write-Warning -Message $WarningMessage
        break
    }

    try {
        $MainProgress['Status'] = "Grouping / Filtering `$IgnoreMailboxPSmtp, `$IgnoreTrusteePSmtp, `$IgnorePermissionType"
        Write-Progress @MainProgress -CurrentOperation "Step 2 of 2: Grouping mailbox-trustee records into unique 1 to 1 relationships"

        $SqlGroupByUniqueRelationships = $null
        $SqlGroupByUniqueRelationships = "USE [$($SqlDatabase)]`n" +
        "INSERT INTO $($SqlTable2) (MailboxPSmtp, MailboxType, TrusteePSmtp, TrusteeType)`n" +
        "SELECT MailboxPSmtp, MailboxType, TrusteePSmtp, TrusteeType`n" +
        "FROM $($SqlTable1)`n"

        if ($PSBoundParameters.ContainsKey('IgnoreMailboxPSmtp')) {

            $SqlGroupByUniqueRelationships += "WHERE MailboxPSmtp NOT IN (`n" +
            "'$((($IgnoreMailboxPsmtp -replace ""'"",""''"") -join ',') -replace ',',""','"")'`n" +
            ")`n"

            if (($PSBoundParameters.ContainsKey('IgnoreTrusteePSmtp')) -or
                ($PSBoundParameters.ContainsKey('IgnorePermissionType')) -or
                ($PSBoundParameters.ContainsKey('IgnoreMailboxType'))) {

                $SqlGroupByUniqueRelationships += "AND "
            }
        }

        if ($PSBoundParameters.ContainsKey('IgnoreTrusteePSmtp')) {

            if ($PSBoundParameters.ContainsKey('IgnoreMailboxPSmtp') -eq $false) {

                $SqlGroupByUniqueRelationships += "WHERE "
            }

            $SqlGroupByUniqueRelationships += "TrusteePSmtp NOT IN (`n" +
            "'$((($IgnoreTrusteePsmtp -replace ""'"",""''"") -join ',') -replace ',',""','"")'`n" +
            ")`n"

            if ($PSBoundParameters.ContainsKey('IgnorePermissionType') -or
            ($PSBoundParameters.ContainsKey('IgnoreMailboxType'))) {

                $SqlGroupByUniqueRelationships += "AND "
            }
        }

        if ($PSBoundParameters.ContainsKey('IgnorePermissionType')) {

            $IgnoredPermissionTypes = @()
            $IgnorePermissionType |
            ForEach-Object {
                if ($_ -eq 'AllFolders') {

                    $IgnoredPermissionTypes += 'Mailbox root', 'Inbox', 'Calendar', 'Contacts', 'Tasks', 'SentItems'
                }
                else {
                    $IgnoredPermissionTypes += $_ -replace 'SendAs', 'Send-As' -replace 'SendOnBehalf', 'Send on behalf' -replace 'MailboxRoot', 'Mailbox root'
                }
            }

            if (($PSBoundParameters.ContainsKey('IgnoreMailboxPSmtp') -eq $false) -and
                ($PSBoundParameters.ContainsKey('IgnoreTrusteePSmtp') -eq $false)) {

                $SqlGroupByUniqueRelationships += "WHERE "
            }

            $SqlGroupByUniqueRelationships += "PermissionType NOT IN (`n" +
            "'$(($IgnoredPermissionTypes -join ',') -replace ',',""','"")'`n" +
            ")`n" +
            "AND`n" +
            "AccessRights NOT IN (`n" +
            "'$(($IgnoredPermissionTypes -join ',') -replace ',',""','"")'`n" +
            ")`n"

            if ($PSBoundParameters.ContainsKey('IgnoreMailboxType')) {

                $SqlGroupByUniqueRelationships += "AND "
            }
        }

        if ($PSBoundParameters.ContainsKey('IgnoreMailboxType')) {

            $IgnoredMailboxTypes = @()
            $IgnoreMailboxType |
            ForEach-Object {
                if ($_ -eq 'AllResources') {

                    $IgnoredMailboxTypes += 'RoomMailbox', 'EquipmentMailbox'
                }
                elseif ($_ -eq 'RemoteMailboxes') {

                    $IgnoredMailboxTypes += 'RemoteUserMailbox', 'RemoteRoomMailbox', 'RemoteEquipmentMailbox', 'RemoteSharedMailbox'
                }
                else {
                    $IgnoredMailboxTypes += $_
                }
            }

            if (($PSBoundParameters.ContainsKey('IgnoreMailboxPSmtp') -eq $false) -and
                ($PSBoundParameters.ContainsKey('IgnoreTrusteePSmtp') -eq $false) -and
                ($PSBoundParameters.ContainsKey('IgnorePermissionType') -eq $false)) {

                $SqlGroupByUniqueRelationships += "WHERE "
            }

            $SqlGroupByUniqueRelationships += "MailboxType NOT IN (`n" +
            "'$(($IgnoredMailboxTypes -join ',') -replace ',',""','"")'`n" +
            ")`n" +
            "AND`n" +
            "TrusteeType NOT IN (`n" +
            "'$(($IgnoredMailboxTypes -join ',') -replace ',',""','"")'`n" +
            ")`n"
        }

        $SqlGroupByUniqueRelationships += "GROUP BY MailboxPSmtp, MailboxType, TrusteePSmtp, TrusteeType"

        Invoke-Sqlcmd -Query $SqlGroupByUniqueRelationships @InvokeSqlCmdParams -QueryTimeout 65535 #<--: Max. allowed observed in the wild.


        $MainProgress['Status'] = "Applying thresholds"
        Write-Progress @MainProgress -CurrentOperation "Permissive mailbox threshold: $PermissiveMailboxThreshold"

        Start-Sleep -Seconds 1

        $SqlPermissiveMailboxWipe = $null
        $SqlPermissiveMailboxWipe = "USE [$($SqlDatabase)]`n" +
        "DELETE FROM $($SqlTable2)`n" +
        "WHERE MailboxPSmtp IN (`n" +
        "SELECT MailboxPSmtp`n" +
        "FROM $($SqlTable2)`n" +
        "GROUP BY MailboxPSmtp`n" +
        "HAVING COUNT(*) > $PermissiveMailboxThreshold`n" +
        ")"

        Invoke-Sqlcmd -Query $SqlPermissiveMailboxWipe @InvokeSqlCmdParams

        $MainProgress['Status'] = "Applying thresholds"
        Write-Progress @MainProgress -CurrentOperation "Power Trustee threshold: $PowerTrusteeThreshold"

        Start-Sleep -Seconds 1

        $SqlPowerTrusteeWipe = $null
        $SqlPowerTrusteeWipe = "USE [$($SqlDatabase)]`n" +
        "DELETE FROM $($SqlTable2)`n" +
        "WHERE TrusteePSmtp IN (`n" +
        "SELECT TrusteePSmtp`n" +
        "FROM $($SqlTable2)`n" +
        "GROUP BY TrusteePSmtp`n" +
        "HAVING COUNT(*) > $PowerTrusteeThreshold`n" +
        ")"

        Invoke-Sqlcmd -Query $SqlPowerTrusteeWipe @InvokeSqlCmdParams
    }
    catch {
        Write-Warning -Message "A failure occurred while performing the grouping / filtering steps with the data in SQL.`n`nError:`n"
        throw $_
    }

    function lookupPSmtp {
        [CmdletBinding()]
        param(
            [string]$Id,
            [string]$SearchProperty
        )
        $SqlLookupQuery = "USE [$($SqlDatabase)]`n" +
        "SELECT MailboxPSmtp, MailboxType, TrusteePSmtp, TrusteeType`n" +
        "FROM $($SqlTable2)`n" +
        "WHERE (`n" +
        "$($SearchProperty) = '$($Id -replace ""'"",""''"")'`n" +
        ")"

        Invoke-Sqlcmd -Query $SqlLookupQuery @SqlLookupPSmtpParams
    }


    $Web = @()
    $MainCounter = 0
    $DepthTracker = 0
    $MemberId = 0

}

process {

    $StartingPSmtp |
    ForEach-Object {

        $MemberId++

        $Web += [PSCustomObject]@{

            Id             = $MemberId
            PSmtp          = $_
            Type           = 'N/A'
            SourceId       = 0
            SourcePSmtp    = 'None'
            SourceRelation = 'Initial web'
            SourceType     = 'None'
            Depth          = 0
        }

        $StartingWeb = $Web
    }

    $StartingWeb |
    ForEach-Object {

        $CurrentPSmtp = $null
        $CurrentPSmtp = $_.PSmtp
        $CurrentId = $null
        $CurrentId = $_.Id

        $MainCounter++

        $MainProgress['Status'] = "Mailbox/Trustee $($MainCounter) of $($StartingPSmtp.Count): $($CurrentPSmtp) (Current recursion depth = 0 (max: $($MaximumDepth)); web size = $($Web.Count))"

        Write-Progress @MainProgress

        $Depth1ForwardLookup = @()
        $Depth1ReverseLookup = @()

        $CurrentForwardLookup = @()
        $CurrentForwardLookup += lookupPSmtp -Id $CurrentPSmtp -SearchProperty MailboxPSmtp
        $CurrentForwardLookup |
        ForEach-Object {
            $Depth1ForwardLookup += [PSCustomObject]@{

                MailboxPSmtp   = $_.MailboxPSmtp
                MailboxType = $_.MailboxType
                TrusteePSmtp         = $_.TrusteePSmtp
                TrusteeType          = $_.TrusteeType
                SourceId             = $CurrentId
                SourceRelation       = 'Mailbox'
            }
        }

        $CurrentReverseLookup = @()
        $CurrentReverseLookup += lookupPSmtp -Id $CurrentPSmtp -SearchProperty TrusteePSmtp
        $CurrentReverseLookup |
        ForEach-Object {
            $Depth1ReverseLookup += [PSCustomObject]@{

                MailboxPSmtp   = $_.MailboxPSmtp
                MailboxType    = $_.MailboxType
                TrusteePSmtp   = $_.TrusteePSmtp
                TrusteeType    = $_.TrusteeType
                SourceId       = $CurrentId
                SourceRelation = 'Trustee'
            }
        }

        $CurrentLookupTrustees = $null
        $CurrentLookupTrusteeFor = $null

        if ($CurrentForwardLookup.Count -gt 0) { $CurrentLookupTrustees = "$(($CurrentForwardLookup.TrusteePSmtp | Select-Object -Unique) -join ';')" }

        if ($CurrentReverseLookup.Count -gt 0) { $CurrentLookupTrusteeFor = "$(($CurrentReverseLookup.MailboxPSmtp | Select-Object -Unique)-join ';')" }

        $Web |
        Where-Object { $_.PSmtp -eq $CurrentPSmtp } |
        Add-Member -NotePropertyName Trustees -NotePropertyValue $CurrentLookupTrustees -PassThru |
        Add-Member -NotePropertyName TrusteeFor -NotePropertyValue $CurrentLookupTrusteeFor

        if ($Depth1ForwardLookupResults) { $Depth1ForwardLookupResults += $Depth1ForwardLookup }
        else { $Depth1ForwardLookupResults = $Depth1ForwardLookup }

        if ($Depth1ReverseLookupResults) { $Depth1ReverseLookupResults += $Depth1ReverseLookup }
        else { $Depth1ReverseLookupResults = $Depth1ReverseLookup }


        $InnerProgress = @{
            Activity = 'Performing lookups'
            Id       = 1
            ParentId = 0
        }

        for (
            $i = 1
            $i -le $MaximumDepth
            $i++
        ) {
            $MainProgress['Status'] = "Mailbox/Trustee $($MainCounter) of $($StartingPSmtp.Count): $($CurrentPSmtp) (Current recursion depth = $($i+1) (max: $($MaximumDepth)); web size = $($Web.Count))"
            Write-Progress @MainProgress

            New-Variable -Name "Depth$($i+1)ForwardLookup" -Value @() -Force
            New-Variable -Name "Depth$($i+1)ReverseLookup" -Value @() -Force

            $FLItemCounter = 0

            $InnerProgress['Status'] = "Processing level $i's forward lookup results (step 1 of 2 (per level))"

            $CurrentLevelFLItems = $null
            $CurrentLevelFLItems = Get-Variable -Name "Depth$($i)ForwardLookup" -ValueOnly
            $CurrentLevelFLItems |
            ForEach-Object {

                $MemberId++
                $FLItemCounter++

                if (($Web.PSmtp -notcontains $_.TrusteePSmtp) -and
                    ($Web.SourcePSmtp -notcontains $_.TrusteePSmtp)) {

                    $InnerProgress['PercentComplete'] = (($FLItemCounter / $CurrentLevelFLItems.Count) * 100)
                    Write-Progress @InnerProgress -CurrentOperation "Forward/reverse lookup $($FLItemCounter) of $($CurrentLevelFLItems.Count): $($_.TrusteePSmtp)"

                    $CurrentForwardLookup = @()
                    $CurrentForwardLookup += lookupPSmtp -Id $_.TrusteePSmtp -SearchProperty MailboxPSmtp
                    $CurrentForwardLookup |
                    ForEach-Object {
                        (Get-Variable -Name "Depth$($i+1)ForwardLookup").Value +=

                        [PSCustomObject]@{

                            MailboxPSmtp   = $_.MailboxPSmtp
                            MailboxType    = $_.MailboxType
                            TrusteePSmtp   = $_.TrusteePSmtp
                            TrusteeType    = $_.TrusteeType
                            SourceId       = $MemberId
                            SourceRelation = 'Mailbox'
                        }
                    }

                    $CurrentReverseLookup = @()
                    $CurrentReverseLookup += lookupPSmtp -Id $_.TrusteePSmtp -SearchProperty TrusteePSmtp
                    $CurrentReverseLookup |
                    ForEach-Object {
                        (Get-Variable -Name "Depth$($i+1)ReverseLookup").Value +=

                        [PSCustomObject]@{

                            MailboxPSmtp   = $_.MailboxPSmtp
                            MailboxType    = $_.MailboxType
                            TrusteePSmtp   = $_.TrusteePSmtp
                            TrusteeType    = $_.TrusteeType
                            SourceId       = $MemberId
                            SourceRelation = 'Trustee'
                        }
                    }

                    $CurrentLookupTrustees = $null
                    $CurrentLookupTrusteeFor = $null

                    if ($CurrentForwardLookup.Count -gt 0) { $CurrentLookupTrustees = "$(($CurrentForwardLookup.TrusteePSmtp | Select-Object -Unique)-join ';')" }

                    if ($CurrentReverseLookup.Count -gt 0) { $CurrentLookupTrusteeFor = "$(($CurrentReverseLookup.MailboxPSmtp | Select-Object -Unique) -join ';')" }

                    $Web += [PSCustomObject]@{

                        Id             = $MemberId
                        PSmtp          = $_.TrusteePSmtp
                        Type           = $_.TrusteeType
                        SourceId       = $_.SourceId
                        SourcePSmtp    = $_.MailboxPSmtp
                        SourceType     = $_.MailboxType
                        SourceRelation = $_.SourceRelation
                        Depth          = $i
                        Trustees       = $CurrentLookupTrustees
                        TrusteeFor     = $CurrentLookupTrusteeFor
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

                    $InnerProgress['PercentComplete'] = (($RLItemCounter / $CurrentLevelRLItems.Count) * 100)
                    Write-Progress @InnerProgress -CurrentOperation "Forward/reverse lookup $($RLItemCounter) of $($CurrentLevelRLItems.Count): $($_.MailboxPSmtp)"

                    $CurrentForwardLookup = @()
                    $CurrentForwardLookup += lookupPSmtp -Id $_.MailboxPSmtp -SearchProperty MailboxPSmtp
                    $CurrentForwardLookup |
                    ForEach-Object {
                        (Get-Variable -Name "Depth$($i+1)ForwardLookup").Value +=

                        [PSCustomObject]@{

                            MailboxPSmtp   = $_.MailboxPSmtp
                            MailboxType    = $_.MailboxType
                            TrusteePSmtp   = $_.TrusteePSmtp
                            TrusteeType    = $_.TrusteeType
                            SourceId       = $MemberId
                            SourceRelation = 'Mailbox'
                        }
                    }

                    $CurrentReverseLookup = @()
                    $CurrentReverseLookup += lookupPSmtp -Id $_.MailboxPSmtp -SearchProperty TrusteePSmtp
                    $CurrentReverseLookup |
                    ForEach-Object {
                        (Get-Variable -Name "Depth$($i+1)ReverseLookup").Value +=

                        [PSCustomObject]@{

                            MailboxPSmtp   = $_.MailboxPSmtp
                            MailboxType    = $_.MailboxType
                            TrusteePSmtp   = $_.TrusteePSmtp
                            TrusteeType    = $_.TrusteeType
                            SourceId       = $MemberId
                            SourceRelation = 'Trustee'
                        }
                    }

                    $CurrentLookupTrustees = $null
                    $CurrentLookupTrusteeFor = $null

                    if ($CurrentForwardLookup.Count -gt 0) {
                        $CurrentLookupTrustees = "$(($CurrentForwardLookup.TrusteePSmtp | Select-Object -Unique)-join ';')"
                    }

                    if ($CurrentReverseLookup.Count -gt 0) {
                        $CurrentLookupTrusteeFor = "$(($CurrentReverseLookup.MailboxPSmtp | Select-Object -Unique) -join ';')"
                    }

                    $Web += [PSCustomObject]@{

                        Id             = $MemberId
                        PSmtp          = $_.MailboxPSmtp
                        Type           = $_.MailboxType
                        SourceId       = $_.SourceId
                        SourcePSmtp    = $_.TrusteePSmtp
                        SourceType     = $_.TrusteeType
                        SourceRelation = $_.SourceRelation
                        Depth          = $i
                        Trustees       = $CurrentLookupTrustees
                        TrusteeFor     = $CurrentLookupTrusteeFor
                    }
                }
            }

            if ($i -eq $MaximumDepth) { $DepthTracker = $i }

            if (((Get-Variable -Name "Depth$($i+1)ForwardLookup").Value.Count -eq 0) -and
                ((Get-Variable -Name "Depth$($i+1)ReverseLookup").Value.Count -eq 0)) {

                if ($i -gt $DepthTracker) { $DepthTracker = $i }
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


        Write-Progress @InnerProgress -CurrentOperation 'Level completed.'

    } # end $StartingPSmtp | ForEach-Object {}
} # end process {}

end {

    switch ($SqlNoCleanupAfter) {
        $false {
            try {
                if ($PSCmdlet.ShouldProcess(
                        "SQL Server instance: $($SqlServerInstance); Database: [$($SqlDatabase)]",
                        'Delete tables and imported CSV data from SQL?')) {

                    if (Invoke-Sqlcmd -Query $SqlTableCheck1 @InvokeSqlCmdParams) {

                        Write-Progress @MainProgress -CurrentOperation "Dropping table $($SqlTable1)"

                        Invoke-Sqlcmd -Query "USE [$($SqlDatabase)] DROP TABLE $($SqlTable1)" @InvokeSqlCmdParams
                        Start-Sleep -Milliseconds 500
                    }
                    else {
                        Write-Warning "Table 'GetMailboxTrusteeWeb1' wasn't found."
                    }
                    if (Invoke-Sqlcmd -Query $SqlTableCheck2 @InvokeSqlCmdParams) {

                        Write-Progress @MainProgress -CurrentOperation "Dropping table $($SqlTable2)"

                        Invoke-Sqlcmd -Query "USE [$($SqlDatabase)] DROP TABLE $($SqlTable2)" @InvokeSqlCmdParams
                        Start-Sleep -Milliseconds 500
                    }
                    else {
                        Write-Warning "Table 'GetMailboxTrusteeWeb2' wasn't found."
                    }
                }
                else {
                    { Write-Warning -Message "SQL cleanup skipped." }
                }
            }
            catch {
                Write-Warning -Message "A failure occurred while attempting to perform SQL cleanup tasks.`n`nError:`n"
                throw $_
            }
        }
    }

    Write-Progress @InnerProgress -Completed
    Write-Progress @MainProgress -Completed

    $Web

    if ([Environment]::GetCommandLineArgs() -notmatch '-noni*') {

        $EndTime = Get-Date

        $ScriptRuntimeDetails = [Ordered]@{

            'Final web size'               = $Web.Count
            'Starting web size'            = $StartingPSmtp.Count
            'Start time'                   = $StartTime.ToLongTimeString()
            'End time'                     = $EndTime.ToLongTimeString()
            'Duration'                     = $EndTime - $StartTime -replace '\..*', ''
            'Depth reached'                = $DepthTracker
            'Maximum depth'                = $MaximumDepth
            'Permissive mailbox threshold' = $PermissiveMailboxThreshold
            'Power trustee threshold'      = $PowerTrusteeThreshold
            'Ignored mailboxes'            = $IgnoreMailboxPSmtp.Count
            'Ignored trustees'             = $IgnoreTrusteePSmtp.Count
            'Ignored permission types'     = $IgnorePermissionType -join ','
            'Ignored mailbox types'        = $IgnoreMailboxType -join ','
        }

        if ($PSBoundParameters.ContainsKey('OutVariable')) {

            $ScriptRuntimeDetails.Add('OutVariable', "`$$($PSBoundParameters.OutVariable)")
        }

        if ($PSBoundParameters.ContainsKey('InformationVariable')) {

            $ScriptRuntimeDetails.Add('InformationVariable', "`$$($PSBoundParameters.InformationVariable)")
        }

        $DepthTallies = Get-Variable -Name Depth*LookupResults

        $LongestDepthVariable = ($DepthTallies.Name | Measure-Object -Property Length -Maximum).Maximum

        $DepthTalliesSortable = @()

        $DepthTallies |
        ForEach-Object {

            $CurrentDepth = $_.Name -split '[A-Za-z]+' -join '' -replace '\s'

            if ($_.Name -match 'Forward') { $LookupType = 'forward' } else { $LookupType = 'reverse' }

            $DepthTalliesSortable += [PSCustomObject]@{

                SortableName = "Depth " + '0' * ($LongestDepthVariable - $_.Name.Length) + "$($CurrentDepth) $($LookupType) lookup results"
                Count        = $_.Value.Count
            }
        }

        $DepthTalliesSortable |
        Sort-Object -Property SortableName |
        ForEach-Object {

            $ScriptRuntimeDetails.Add($_.SortableName, $_.Count)
        }

        $LongestKey = ($ScriptRuntimeDetails.Keys | Measure-Object -Property Length -Maximum).Maximum

        $AfterActionReport = "`n`tGet-MailboxTrusteeWebSQLEdition.ps1 - After Action Report`n"

        $ScriptRuntimeDetails.Keys |
        ForEach-Object {

            $AfterActionReport += "`n`t$($_[0..40] -join '')" +
            "." * (($LongestKey + 4) - ($_[0..40].Length)) +
            ": $($ScriptRunTimeDetails.$($_))"
        }

        Write-Information -MessageData $AfterActionReport -InformationAction:Continue
        Write-Information -MessageData "`nCommand:`n`n$($PSCmdlet.MyInvocation.Line)" -InformationAction:Continue
    }

}

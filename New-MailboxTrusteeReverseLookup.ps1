<#

    .Synopsis

    Find all mailbox & mailbox folder permissions held by a trustee user or group.

    This is the sister script to Get-MailboxTrustee.ps1.  Therefore it requires
    that the passed CSV file(S) contains headers for the properties output by that
    script.  Any other headers (i.e. columns) are ignored.

    There are two distinct accepted sets of headers, corresponding to the
    standard and -MinimizeOutput modes of Get-MailboxTrustee.ps1.

    Standard:

    - Alias
    - DisplayName
    - DistinguishedName
    - Guid
    - PrimarySmtpAddress
    - RecipientTypeDetails
    - SamAccountName
    - PermissionType
    - AccessRights
    - TrusteeDisplayName
    - TrusteeDN
    - TrusteeGuid
    - TrusteePSmtp
    - TrusteeType

    Minimized Output:

    - MailboxGuid
    - PermissionType
    - AccessRights
    - TrusteeGuid


    .Parameter GetMailboxTrusteeCsvFilePath

    The full or direct path to the CSV file(s) containing objects that have been
    output from Get-MailboxTrustee.ps1.  Must contain all headers from either the
    Standard or Minizmed Output sets.

    Since it's common to have run Get-MailboxTrustee.ps1 twice - once against EXO
    and once against Exchange On-Premises - multiple CSV files can be specified
    here, to spare the need to combine CSV files manually.


    .Parameter TrusteeId

    Expected values are any of the Trustee_____ properties from a
    Get-MailboxTrustee.ps1-outputted object:

    - Guid
    - DistinguishedName
    - PrimarySmtpAddress
    - DisplayName


    .Link

    https://github.com/JeremyTBradshaw/PowerShell/blob/master/New-MailboxTrusteeReverseLookup.ps1

#>

#Requires -Version 3

[CmdletBinding()]

param(

    [Parameter(Mandatory = $true)]
    [ValidateScript( {
        if ((Test-Path -Path $_) -eq $false) {throw "Can't find file '$($_)'."}
        if ($_ -notmatch '(\.csv$)') {throw "Only .csv files are accepted."}
        $true
    })]
    [System.IO.FileInfo[]]$GetMailboxTrusteeCsvFilePath,

    [Parameter(Mandatory = $true)]
    [ValidateScript({
        $_ | ForEach-Object {

            if ($_.Length -gt 256) {throw 'TrusteeId must be 256 characters or less (maximum for DisplayName)'}
            else {$true}
        }
    })]
    [string[]]$TrusteeId
)

begin {

    $StartTime = Get-Date

    $MainProgress = @{

        Activity         = "New-MailboxTrusteeReverseLookup.ps1 (Start time: $($StartTime.DateTime))"
        Id               = 0
        ParentId         = -1
        SecondsRemaining = -1
        Status           = 'Initializing'
    }
    Write-Progress @MainProgress

    $StandardHeaders = @(
        'AccessRights',
        'Alias',
        'DisplayName',
        'DistinguishedName',
        'Guid',
        'PermissionType',
        'PrimarySmtpAddress',
        'RecipientTypeDetails',
        'SamAccountName',
        'TrusteeDisplayName',
        'TrusteeDN',
        'TrusteeGuid',
        'TrusteePSmtp',
        'TrusteeType'
    )

    $MinimizedOutputHeaders = @(
        'MailboxGuid',
        'PermissionType',
        'AccessRights',
        'TrusteeGuid'
    )

    $MailboxTrustees = @()

    $CsvCounter = 0
    $GetMailboxTrusteeCsvFilePath |
    ForEach-Object {

        $CsvCounter++
        Write-Progress @MainProgress -CurrentOperation "Importing CSV file $($CsvCounter) of $($GetMailboxTrusteeCsvFilePath.Count): $($_)" -PercentComplete (($CsvCounter / $GetMailboxTrusteeCsvFilePath.Count) * 100)

        $CurrentCsv = @()
        $CurrentCsv += Import-Csv -Path $_

        $CurrentCsvHeaders    = $CurrentCsv |
                                Get-Member -MemberType NoteProperty

        $break = $false

        if ($CurrentCsvHeaders.Name -eq 'MailboxGuid') {

            $CsvHeaderSet = 'MinimizedOutput'

            $MinimizedOutputHeaders |
            ForEach-Object {
                if ($CurrentCsvHeaders.Name -notcontains $_) {$break = $true}
            }
        }

        else {
            $CsvHeaderSet = 'Standard'
            $StandardHeaders |
            ForEach-Object {
                if ($CurrentCsvHeaders.Name -notcontains $_) {$break = $true}
            }
        }

        if ($break -eq $true) {

            $WarningMessage   = "CSV file '$($GetMailboxTrusteeCsvFilePath)' is missing one or more mandatory headers.`n`n" +
                                "See help:`n`n`t" +
                                '.\New-MailboxTrusteeReverseLookup.ps1 -?'
            Write-Warning -Message $WarningMessage
            break
        }

        else {
            $MailboxTrustees += $CurrentCsv
        }
    }

}

process {

    $TrusteeCounter = 0
    $MainProgress['Status'] = 'Working'

    $TrusteeId |
    ForEach-Object {

        $CurrentTrustee = $_

        $TrusteeCounter++
        Write-Progress @MainProgress -CurrentOperation "Validating TrusteeId #$($TrusteeCounter) of $($TrusteeId.Count): $($_)"

        switch ($CsvHeaderSet) {

            Standard {

                if ($CurrentTrustee -match '^([0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12})$') {

                    $TrusteeIdProperty = 'TrusteeGuid'
                }
                elseif ($CurrentTrustee -match '^CN=.*DC=.*') {$TrusteeIdProperty = 'TrusteeDN'}
                elseif ($CurrentTrustee -match '^.*\@.*\..*$') {$TrusteeIdProperty = 'TrusteePSmtp'}
                else {$TrusteeIdProperty = 'DisplayName'}
            }

            MinimizedOutput {

                if ($CurrentTrustee -match '^([0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12})$') {

                    $TrusteeIdProperty = 'TrusteeGuid'
                }
                else {$break = $true}
            }
        }

        if ($break -eq $true) {

            $WarningMessage =
            "TrusteeId '$($_)' failed validation.  Expected inputs are:`n`n" +
            "- Guid`n- DistinguishedName`n- PrimarySmtpAddress`n- Displayname (least preferred due to being non-unique)`n`n" +
            "The TrusteeId format must match one of the Trustee____ properties in the CSV file(s) specified with -GetMailboxTrusteeCsvFilePath."
            Write-Warning -Message $WarningMessage
            break
        }

        Write-Progress @MainProgress -CurrentOperation "Performing reverse lookup of TrusteeId $($TrusteeCounter) of $($TrusteeId.Count): $($TrusteeIdProperty) -eq $($CurrentTrustee)" -PercentComplete (($TrusteeCounter / $TrusteeId.Count) * 100)

        Write-Debug -Message "`$TrusteeId | ForEach-Object {*}"

        $MailboxTrustees |
        Where-Object {$_.$($TrusteeIdProperty) -eq $CurrentTrustee}
    }

}

end {
    $MainProgress['Status'] = 'Completed'
    Write-Progress @MainProgress -Completed
}

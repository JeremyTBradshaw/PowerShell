<#
    .Synopsis
    Get all mailbox-enabled AD users, including some chosen properties.

    .Parameter ADServer
    This is a direct passthrough for the -Server parameter of Get-ADUser.  If not specified, that parameter won't be
    used, so the local computer's domain will be tried by Get-ADUser naturally/instead.

    .Parameter TestMode
    Switch parameter to enable the -ResultSetSize parameter for testing with a smaller number of users.

    .Parameter ResultSetSize
    Default value is 5000, but is only active when the -TestMode switch is used.  Otherwise the script gets all users
    in the forest.

    .Example
    $Start = [datetime]::Now
    $Users = .\Documents\Scripts\Get-ADMailboxUser.ps1
    $Users | Export-Csv .\Desktop\MailboxUsers-2020-05-26.csv -NTI
    "Start: $($Start)"
    "End: $([datetime]::Now)"
    "User count: $($Users.Count)"

    .Example
    .\Documents\Scripts\Get-ADMailboxUser.ps1 -TestMode -ResultSetSize 500 | Export-Csv .\Desktop\Test-Sample.csv -NTI
#>
#Requires -Version 4
#Requires -Module ActiveDirectory

[CmdletBinding()]
param(
    [Parameter(
        HelpMessage = "Enter a domain or domain controller FQDN.  Include ':3268' to target the Global Catalog (i.e. to search the entire forest)."
    )]
    [string]$ADServer,


    [switch]$TestMode,

    [ValidateRange(1, 50000)]
    [int]$ResultSetSize = 5000
)

$StartDateTime = [DateTime]::Now
$StopWatch1 = [System.Diagnostics.Stopwatch]::StartNew()

$Progress = @{

    Activity         = 'Get-ADMailboxUser.ps1'
    CurrentOperation = 'Initializing'
    Status           = "Start time: $($StartDateTime.ToString('yyyy-MM-dd hh:mm:ss tt (zzzz)'))"
}

$msExchRecipientTypeDetails = @{

    1             = 'UserMailbox'
    2             = 'LinkedMailbox'
    4             = 'SharedMailbox'
    8             = 'LegacyMailbox'
    16            = 'RoomMailbox'
    32            = 'EquipmentMailbox'
    8192          = 'SystemAttendantMailbox'
    16384         = 'SystemMailbox'
    8388608       = 'ArbitrationMailbox'
    536870912     = 'DiscoveryMailbox'
    2147483648    = 'RemoteUserMailbox'
    8589934592    = 'RemoteRoomMailbox'
    17179869184   = 'RemoteEquipmentMailbox'
    34359738368   = 'RemoteSharedMailbox'
    549755813888  = 'MonitoringMailbox'
    4398046511104 = 'AuditLogMailbox'
}

Write-Debug -Message 'Quick stop and check before slow Get-ADUser command.'

$Progress['CurrentOperation'] = 'Getting all mailbox-enabled AD users.'
Write-Progress @Progress

$GetADUser = @{

    Filter     = "msExchRecipientTypeDetails -eq $($msExchRecipientTypeDetails.Keys -join ' -or msExchRecipientTypeDetails -eq ')"
    Properties = 'displayName', 'mail', 'canonicalName', 'msExchRecipientTypeDetails', 'LastLogonDate', 'targetAddress'
}
if ($PSBoundParameters.ContainsKey('ADServer')) { $GetADUser['Server'] = $ADServer}
if ($TestMode) { $GetADUser['ResultSetSize'] = $ResultSetSize }

try {
    $AllMailboxUsers = Get-ADUser @GetADUser -ErrorAction Stop
}
catch {
    Write-Warning -Message "Scripted ended prematurely! Failed on getting all AD users.  Error:`n $($_.Exception)"
    break
}

$Progress['CurrentOperation'] = "Processing $($AllMailboxUsers.Count) users"
$StopWatch2 = [System.Diagnostics.Stopwatch]::StartNew()

$UserCounter = 0
$AllMailboxUsers | & {

    process {

        $UserCounter++

        if ($StopWatch2.Elapsed.TotalMilliseconds -ge 500) {

            $Progress['PercentComplete'] = (($UserCounter / $AllMailboxUsers.Count) * 100)
            $Progress['CurrentOperation'] = "Processed $($UserCounter) of $($AllMailboxUsers.Count) users.  Time elapsed: $($StopWatch1.Elapsed.ToString() -replace '\..*')"
            Write-Progress @Progress

            $StopWatch2.Reset(); $StopWatch2.Start()
        }

        $OU = ($_.CanonicalName -split '\/' | Select-Object -SkipLast 1) -join '/'
        $OUClassification = if ($OU -match '(disable)|(inactive)') { 'Disabled/Inactive' } else { 'Standard' }

        $RecipientTypeDetails = ''
        if ($_.msExchRecipientTypeDetails) {

            $RecipientTypeDetails = $msExchRecipientTypeDetails[($msExchRecipientTypeDetails.Keys -eq $_.msExchRecipientTypeDetails)][0]
        }

        $EmailAddressDomain = if ($_.mail) { $_.mail -replace '.*\@' } else { '' }

        $MbxType = if ($RecipientTypeDetails -notmatch 'Mailbox') { 'No Mailbox' }
        elseif ($RecipientTypeDetails -match 'User') { 'User' }
        elseif ($RecipientTypeDetails -match 'Shared') { 'Shared' }
        elseif ($RecipientTypeDetails -match '(Room)|(Equipment)') { 'Resource' }
        elseif ($RecipientTypeDetails -match '(Arbitration)|(AuditLog)|(Discovery)|(Monitoring)|(System)') { 'System' }
        else { 'Unknown' }

        $L150 = 'Not logged in'
        $LLD = ''
        if ($_.LastLogonDate) {

            if (($StartDateTime - $_.LastLogonDate).Days -le 150) { $L150 = 'Logged in' }
            $LLD = $_.LastLogonDate.ToString('yyyy-MM-dd')
        }

        Write-Debug -Message "Currently in the main loop of the script, just before outputting the assembled object."

        [PSCustomObject]@{

            DisplayName                = $_.displayName
            EmailAddress               = $_.mail
            EmailAddressDomain         = $EmailAddressDomain
            Domain                     = $OU -replace '\/.*'
            'Account State'            = if ($_.Enabled -eq $true) { 'Enabled' } else { 'Disabled' }
            'MBX Type'                 = $MbxType
            'Regular / Test'           = if ($_.mail -match 'test') { 'Test' } else { 'Regular' }
            'OU Classification'        = $OUClassification
            'Last 150 Days'            = $L150
            LastLogonDate              = $LLD
            RecipientTypeDetails       = $RecipientTypeDetails
            OrganizationalUnit         = $OU
            ObjectGuid                 = $_.ObjectGuid.Guid
            msExchRecipientTypeDetails = $_.msExchRecipientTypeDetails
            ObjectClass                = $_.ObjectClass
        }
    }
}

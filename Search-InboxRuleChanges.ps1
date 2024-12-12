<#
    .Synopsis
    Search the Office 365 unified audit log for susicious inbox rule activity.

    .Description
    For OWA-based user acitvity and PowerShell-based admin activity, we can
    search for the New-/Set-InboxRule operations.

    For Outlook client based activity, we can search for the UpdateInboxRules
    activity.

    .Parameter StartDate
    Provide a start date (and optionally time) in a System.DateTime-recognized
    format.  Default is to search back 24 hours (i.e. (Get-Date).AddDays(-1)).

    .Parameter EndDate
    Provide an end date (and optionally time) in a System.DateTime-
    recognized format.  Default is current date/time (i.e. (Get-Date)).

    .Parameter ResultSize
    By default, the maximum (5000) is specified.  Valid range is 1-5000

    .Parameter UseClientIPExcludedRanges
    This bool is $true by default.  Update the section of the script:

        if ($UseClientIPExcludedRanges -eq $true) {}

    This allows us to filter output to only changes made from outside the
    corporate network.  The common use case of this script case is to detect
    when an account has been compromised and the attacker creates a rule to
    hide NDR backscatter, allowing them to send spam while delaying the mailbox
    owner becoming aware.

    .Notes
    I have decided to have the script process all results of the search then
    output all entries at the end, rather than outputting each log entry
    individually, directly after processing.  This allows for the discovery of
    all audit log entries' list of properties so that a common PS custom object
    can be output for every log entry.  This helps with ensuring down the line
    cmdlets (e.g. Export-Csv) will work predictably.  There could be more
    efficient ways to accomplish this, but I've settled on this one until I
    find a more favorable method.

    Note that this dynamic list of properties challenge is also felt by Excel,
    as is noted in the following article:
    https://docs.microsoft.com/en-us/microsoft-365/compliance/export-view-audit-log-records

    The method of dealing with nested multi-valued properties (sometimes in
    JSON format) in this script results in many properties (i.e. columns) in
    the output.  This is hoped to be superior to how the same data will be
    presented in Excel if the process from the link above is followed instead.

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Search-InboxRuleChanges.ps1

    .Link
    # [Unified audit log] Audited Activities:
    https://docs.microsoft.com/en-us/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance#audited-activities

    .Link
    # [Unified audit log] Detailed Properties
    https://docs.microsoft.com/en-us/microsoft-365/compliance/detailed-properties-in-the-office-365-audit-log

    .Example
    .\Search-InboxRuleChanges.ps1

    .Example
    .\Search-InboxRuleChanges.ps1 -UseClientIPExcludedRanges $false -ResultSize 100 -StartDate (Get-Date).AddHours(-4)
#>
#Requires -Version 5.1
#Requires -Modules @{ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.6.0'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'}

[CmdletBinding()]
param (
    [datetime]$StartDate = (Get-Date).AddDays(-1),
    [datetime]$EndDate = (Get-Date),
    [ValidateRange(1, 5000)][int]$ResultSize = 5000,
    [bool]$UseClientIPExcludedRanges = $true
)

Write-Verbose -Message "Verifying there's an active connection to EXO PowerShell."
$exoConnectionInfo = Get-ConnectionInformation -ErrorAction Stop
if (-not ($exoConnectionInfo | Where-Object { ($_.TokenStatus -eq 'Active') -and ($_.ConnectionUri -like '*outlook.office365.com') })) {
    Write-Warning -Message "Please be connected to EXO via Connect-ExchangeOnline before running this script."
    break
}

if ($UseClientIPExcludedRanges -eq $true) {

    Write-Warning -Message 'Using predefined ClientIP excluded IPRanges.'
    Write-Warning -Message 'To avoid this, use -UseClientIPExcludedRanges $false'

    # Ensure to use outside-facing IP's (e.g. NAT'd, external).
    # Since we're searching in EXO, all ClientIP's will be public IP addresses.

    $ClientIPExcludedIPRanges = @()
    foreach ($i in (1..254)) { $ClientIPExcludedIPRanges += "192.168.1.$i" } # <--: Example (but don't actually use private/internal IP's).
    foreach ($i in (1..254)) { $ClientIPExcludedIPRanges += "192.168.2.$i" }
    foreach ($i in (80..90)) { $ClientIPExcludedIPRanges += "10.10.10.$i" }
}

Write-Verbose -Message "Performing search..."

$SearchResults = Search-UnifiedAuditLog -Operations New-InboxRule, Set-InboxRule, UpdateInboxRules -StartDate $StartDate -EndDate $EndDate -ResultSize:$ResultSize
$SearchResultsProcessed = @()

foreach ($sr in $SearchResults) {

    Write-Verbose -Message "Processing log entry $($sr.ResultIndex) of $($sr.ResultCount)"

    $AuditData = $null
    $AuditData = $sr.AuditData | ConvertFrom-Json

    if ($UseClientIPExcludedRanges -eq $true) {

        if ($ClientIPExcludedIPRanges -notcontains "$($AuditData.ClientIP -replace '\[' -replace '\].*' -replace ':.*')") { $ContinueProcessing = $true }
        else { $ContinueProcessing = $false }
    }
    else { $ContinueProcessing = $true }

    if ($ContinueProcessing -eq $true) {

        $ProcessedLogEntry = $null
        $ProcessedLogEntry = [PSCustomObject]@{

            RecordType      = $sr.RecordType
            CreationDateUTC = $sr.CreationDate.ToString('yyyy-MM-dd hh:mm:ss tt')
            UserIds         = $sr.UserIds
            Operations      = $sr.Operations
            ResultIndex     = $sr.ResultIndex
            ResultCount     = $sr.ResultCount
            ClientIP        = $AuditData.ClientIP
            UserId          = $AuditData.UserId
            ExternalAccess  = $AuditData.ExternalAccess
        }

        $InboxRule = @()

        if ($sr.Operations -eq 'UpdateInboxRules') {

            $ProcessedLogEntry |
            Add-Member -NotePropertyName ClientInfoString -NotePropertyValue $AuditData.ClientInfoString -PassThru |
            Add-Member -NotePropertyName ClientProcessName -NotePropertyValue $AuditData.ClientProcessName -PassThru |
            Add-Member -NotePropertyName ClientVersion -NotePropertyValue $AuditData.ClientVersion -PassThru |
            Add-Member -NotePropertyName LogonUserSid -NotePropertyValue $AuditData.LogonUserSid -PassThru |
            Add-Member -NotePropertyName MailboxOwnerSid -NotePropertyValue $AuditData.MailboxOwnerSid -PassThru |
            Add-Member -NotePropertyName MailboxOwnerUPN -NotePropertyValue $AuditData.MailboxOwnerUPN -PassThru |
            Add-Member -NotePropertyName MailboxGuid -NotePropertyValue $AuditData.MailboxGuid

            $OperationProperties = $null
            $OperationProperties = $AuditData | Select-Object -ExpandProperty OperationProperties

            foreach ($opn in $OperationProperties.Name) {

                if ($opn -match 'RuleActions') {

                    $RuleActions = $null
                    $RuleActions = $OperationProperties.Value[$OperationProperties.Name.IndexOf($opn)] | ConvertFrom-Json
                    $RAProps = $RuleActions | Get-Member -MemberType NoteProperty

                    foreach ($rap in $RAProps.Name) {

                        $ProcessedLogEntry |
                        Add-Member -NotePropertyName "RuleAction_$($rap)" -NotePropertyValue $RuleActions.$rap
                    }
                }
                else {
                    $ProcessedLogEntry |
                    Add-Member -NotePropertyName $opn -NotePropertyValue $OperationProperties.Value[$OperationProperties.Name.IndexOf($opn)]
                }
            }

            if (($ProcessedLogEntry.RuleOperation -notmatch 'RemoveMailboxRule') -and ($ProcessedLogEntry.RuleName)) {

                $InboxRule += Get-InboxRule "$($AuditData.UserId)\$($ProcessedLogEntry.RuleName)" -ErrorAction:SilentlyContinue
            }
        }
        elseif ($sr.Operations -like '*-InboxRule') {

            $ProcessedLogEntry |
            Add-Member -NotePropertyName ResultStatus -NotePropertyValue $AuditData.ResultStatus -PassThru |
            Add-Member -NotePropertyName ObjectId -NotePropertyValue $AuditData.ObjectId

            $ParametersProperties = $null
            $ParametersProperties = $AuditData | Select-Object -ExpandProperty Parameters

            foreach ($ppn in $ParametersProperties.Name) {

                Write-Debug "Inspect `$ppn, `$ParametersProperties(.name)"
                $ProcessedLogEntry |
                Add-Member -NotePropertyName CmdletParameter_$ppn -NotePropertyValue $ParametersProperties.Value[$ParametersProperties.Name.IndexOf($ppn)]
            }

            $InboxRule += Get-InboxRule $AuditData.ObjectId -ErrorAction:SilentlyContinue
        }
        else {
            $ProcessedLogEntry |
            Add-Member -NotePropertyName LogEntryProblem -NotePropertyValue "'Operations' is not one of New-InboxRule, Set-InboxRule, or UpdateInboxRules"
        }

        if ($ProcessedLogEntry.RuleOperation -notmatch 'RemoveMailboxRule') {

            if ($InboxRule.Count -eq 1) {

                $ProcessedLogEntry |
                Add-Member -NotePropertyName InboxRule_Description -NotePropertyValue $InboxRule.Description
            }
            elseif ($InboxRule.Count -gt 1) {

                $ProcessedLogEntry |
                Add-Member -NotePropertyName InboxRule_Description -NotePropertyValue "Multiple matching rules found - check manually."
            }
            else {
                $ProcessedLogEntry |
                Add-Member -NotePropertyName InboxRule_Description -NotePropertyValue "Rule not found - check manually."
            }
        }

        $SearchResultsProcessed += $ProcessedLogEntry

    } # end: if ($ContinueProcessing -eq $true) {}

} # end: foreach ($sr in $SearchResults) {}

Write-Debug "`$SearchResultsProcessed <--: Check it out!"

if ($SearchResultsProcessed.Count -ge 1) {

    Write-Verbose -Message "Preparing final output..."

    $FinalOutputProperties = @()
    $FinalOutputProperties += $SearchResultsProcessed[0] | Get-Member -MemberType NoteProperty

    foreach ($srp in $SearchResultsProcessed[1..$SearchResultsProcessed.Count]) {

        $FinalOutputProperties += $srp |
        Get-Member -MemberType NoteProperty |
        Where-Object { $FinalOutputProperties.Name -notcontains $_.Name }
    }

    Write-Output $SearchResultsProcessed |
    Select-Object -Property $FinalOutputProperties.Name
}

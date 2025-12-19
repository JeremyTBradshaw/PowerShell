<#
    .Synopsis
    Get a quick digest of the messages in EOP's quarantine.  By default for yesterday's date.

    .Description
    Messages in the EOP Quarantine are summarized in the output as follows:
    - Number of quarantined messages by category (i.e., phish, high-confidence phish, spam, high-confidence spam...)
    - Top 10 (by default) categories.
    - Number of quarantined messages by sender domain.
    - Top 10 (by default) sender domains.
    - Top 10 (by default) recipients.
    - Number of quarantined messages by release status (i.e., Needs Review, Approved, Denied, etc.)

    .Parameter StartReceivedDate
    The start date for which to summarize, based on the received date of the quarantined messages.

    .Parameter EndReceivedDate
    The end date for which to summarize, based on the received date of the quarantined messages.

    .Parameter TopNOverride
    Use the paramter to adjust the digest's 'Top ****' items (e.g., 'Top 10 Sender Domains').  Default is 10.

    .Parameter CSVFilePathForRawQuarantineResults
    Specifies a file path to export a CSV of all found Quarantine messages in the selected time range (previous day by
    default).

    .Example
    .\Get-EOPQuarantineDigest.ps1

    .Example
    .\Get-EOPQuarantineDigest -StartReceivedDate (Get-Date).AddDays(-7) -EndReceivedDate (Get-Date) -TopNOverride 50 -CSVFilePathForRawQuarantineResults .\QuarantineMessages-Last1Week.csv | Out-File .\QuarantineDigest.json

    .Outputs
    The primary output from this script is formatted in JSON (using ConvertTo-Json cmdlet).  Optionally, the
    -CSVFilePathForRawQuarantineResults parameter can be used to export a CSV file.
#>
#Requires -Modules @{ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.0.0'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'}

[CmdletBinding()]
param (
    [datetime]$StartReceivedDate = [datetime]::Today.AddDays(-1),
    [datetime]$EndReceivedDate = [datetime]::Today,
    [ValidateRange(1, 100)]
    [int]$TopNOverride = 10,
    [System.IO.FileInfo]$CSVFilePathForRawQuarantineResults
)

if (-not (Get-Command Get-QuarantineMessage -ErrorAction SilentlyContinue)) {

    throw "This script requires an active connection to Exchange Online PowerShell (using v3.0.0 module or newer), and access to the Get-QuarantineMessage cmdlet."
}

#======#-----------#
#region# Functions #
#======#-----------#

# no Functions needed (yet)

#=========#-----------#
#endregion# Functions #
#=========#-----------#



#======#-----------------------------#
#region#- Initialization and Variables #
#======#------------------------------#

$dtNow = [datetime]::Now

$ht_policyTypes = @{

    'HostedContentFilterPolicy' = 'Anti-Spam'
    'AntiMalwarePolicy'         = 'Anti-Malware'
    'AntiPhishPolicy'           = 'Anti-Phish'
    'SafeAttachmentPolicy'      = 'Safe Attachments'
    'ExchangeTransportRule'     = 'Exchange Transport Rule'
}

$Progress = @{

    Activity = "$($PSCmdlet.MyInvocation.MyCommand.Name) - Start time: $($dtNow)"
    PercentComplete = -1
}

#=========#------------------------------#
#endregion# Initialization and Variables #
#=========#------------------------------#



#======#----------------#
#region# Data Retrieval #
#======#----------------#

$QuarantineResults = @()
$PageNumber = 1
$InLoopResults = @(1)
do {
    # Next line is to get around EXO v3.0.0 issue which sets $ProgressPreference to SilentlyContinue globally:
    $ProgressPreference = 'Continue'
    Write-Progress @Progress -Status "Getting quarantine messages between $($StartReceivedDate) and $($EndReceivedDate)..."

    $InLoopResults = @()
    $InLoopResults += Get-QuarantineMessage -StartReceivedDate $StartReceivedDate -EndReceivedDate $EndReceivedDate -PageSize 1000 -Page $PageNumber
    if ($InLoopResults.Count -ge 1) {

        $QuarantineResults += $InLoopResults
    }
    $PageNumber++
}
until ($InLoopResults.Count -eq 0)

#=========#----------------#
#endregion# Data Retrieval #
#=========#----------------#



#======#-----------------#
#region# Post-Processing #
#======#-----------------#

# Next line is to get around EXO v3.0.0 issue which sets $ProgressPreference to SilentlyContinue globally:
$ProgressPreference = 'Continue'
Write-Progress @Progress -Status "Summarizing $($QuarantineResults.Count) quarantine messages..."

$ResultsForProcessing = $QuarantineResults |
Select-Object Organization, Identity, ReceivedTime, SenderAddress, Subject, Type, ReleaseStatus, SystemReleased, Reported,
@{Name = 'Hour'; Expression = { $_.ReceivedTime.Hour } },
@{Name = 'SenderDomain'; Expression = { $_.SenderAddress -replace '.*\@' } },
@{Name = 'RecipientDomain'; Expression = { $_.RecipientAddress -replace '.*\@' } },
@{Name = 'RecipientAddress'; Expression = { $_.RecipientAddress -join '; ' } }, 
@{Name = 'Policy'; Expression = { "$($ht_policyTypes[$($_.PolicyType)]): $($_.PolicyName)" } }

# Start the report object with the high-level totals:
$ReportCustomObject = [PSCustomObject]@{
    
    'Total Messages' =  $ResultsForProcessing.Count
    'Total Reported' = ($ResultsForProcessing | Where-Object {$_.Reported -eq $true}).Count
    'Total System-Released' = ($ResultsForProcessing | Where-Object {$_.SystemReleased -eq $true}).Count
} 

# Add other totals:
foreach ($property in @('Policy', 'Type', 'ReleaseStatus')) {

    $_nestedDetails = [PSCustomObject]@{}
    foreach ($p in ($ResultsForProcessing | Group-Object $property -NoElement | Sort-Object Name | Select-Object Name, Count)) {
    
        $_nestedDetails | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Count
    }
    $ReportCustomObject |
    Add-Member -NotePropertyName "Total by $($property)" -NotePropertyValue $_nestedDetails
}

# Add the Top 10's (or Top NN):
foreach ($property in @('SenderDomain','SenderAddress','RecipientDomain','RecipientAddress')) {

    $_nestedDetails = [PSCustomObject]@{}
    foreach ($p in ($ResultsForProcessing | Group-Object $property -NoElement | Sort-Object Count -Descending | Select-Object -First $TopNOverride -Property Name, Count)) {
    
        $_nestedDetails | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Count
    }
    $ReportCustomObject |
    Add-Member -NotePropertyName "Top $($TopNOverride) $($property)" -NotePropertyValue $_nestedDetails
}

# Add the Top 10 (or Top NN) Subject, SenderAddress:
$_nestedDetails = [PSCustomObject]@{}
foreach ($p in ($ResultsForProcessing | Group-Object Subject, SenderAddress -NoElement | Sort-Object Count -Descending | Select-Object -First $TopNOverride -Property Name, Count)) {

    $_nestedDetails | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Count
}
$ReportCustomObject |
Add-Member -NotePropertyName "Top $($TopNOverride) Subject/Sender" -NotePropertyValue $_nestedDetails

#=========#-----------------#
#endregion# Post-Processing #
#=========#-----------------#



#======#--------#
#region# Output #
#======#--------#

# Keeping it very simple, and easy (enough) to read, send report object out to JSON:
$ReportCustomObject | ConvertTo-Json

if ($PSBoundParameters.ContainsKey('CSVFilePathForRawQuarantineResults')) {

    try {
        $QuarantineResults | Export-Csv -Path $CSVFilePathForRawQuarantineResults -NTI -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        "Failed to Quarantine results to file path: $($CSVFilePathForRawQuarantineResults).  Error to follow."
        throw
    }
}

#=========#--------#
#endregion# Output #
#=========#--------#

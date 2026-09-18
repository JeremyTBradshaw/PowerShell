<#
.SYNOPSIS
    Converts a Unified Audit Log CSV export into normalized Exchange item records.

.DESCRIPTION
    Caters only the item-related operations:

        Create, Update, Copy, Move
        Send, SendAs, SendOnBehalf
        MailItemsAccessed
        MoveToDeletedItems, SoftDelete, HardDelete

        *MailItemsAccessed are reports as either:
            MailItemsAccessed_Bind (i.e., user read the message), or
            MailItemsAccessed_Sync (i.e,. client synced the item)
    
    Those operations funnel up into these Exchange audit record types:

        ExchangeItem             RecordType 2
        ExchangeItemGroup        RecordType 3
        ExchangeItemAggregated   RecordType 50

    This script gathers pertinent item details (e.g., Subject, Attachment/Recipient counts, folder path(s), etc.) from any of the *4 possible item container columns:

        Item
        AffectedItems
        Folders[].FolderItems
        *Folder

    *Note:  Sometimes no item details are included but there is a Folder column that shows the parent folder of the item(s) being acted on.
            These are returned with ItemSource = 'Folder' and all item-related columns are left blank.

.EXAMPLE
    .\Convert-ExchangeAuditLog.ps1 -RawUnifiedAuditLogCSVFile .\UnifiedAuditLog.csv -OutputCSVFileNamePrefix 5551212

.EXAMPLE
    .\Convert-ExchangeAuditLog.ps1 -RawUnifiedAuditLogCSVFile .\UAL.csv -OutputCSVFileNamePrefix 123456 -OutputFolder c:\reports

.NOTES
    - v1.0.0 (2026-09-10): Initial version, tested and working on 100s of 1000s of audit log rows.
    - v1.0.1 (2026-09-18): Updated CSV import logic to support different columns depending where UAL CSV was exported from (PowerShell vs Purview website).

.OUTPUTS
    CSV file named <OutputCSVFileNamePrefix>_ExchangeItem-Operations.csv, which by default is placed into $HOME\Downloads.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [System.IO.FileInfo]$RawUnifiedAuditLogCSVFile,

    [string]$OutputCSVFileNamePrefix = '5551212',
    [ValidateScript(
        {
            if (Test-Path $_ -PathType Container -ErrorAction SilentlyContinue) { $true } else {
                throw "Folder '$($_)' could not be found."
            }
        }
    )]
    [System.IO.FileInfo]$OutputFolder = "$HOME\Downloads\"
)

$ErrorActionPreference = 'Stop'

function Convert-RecordType {
    param (
        $RecordType
    )

    switch ([string]$RecordType) {
        '2' { 'ExchangeItem' }
        '3' { 'ExchangeItemGroup' }
        '50' { 'ExchangeItemAggregated' }
        default { [string]$RecordType }
    }
}

function Get-AttachmentsCount {
    param (
        $Attachments
    )

    if ([string]::IsNullOrWhiteSpace([string]$Attachments)) { return $null }
    @([string]$Attachments -split '; ').Count
}

function Get-ItemEntries {
    param (
        [Parameter(Mandatory)]
        $Record
    )

    $hasItem = $null -ne $Record.Item
    $hasAffectedItems = $null -ne $Record.AffectedItems
    $hasFolders = $null -ne $Record.Folders
    $hasFolder = $null -ne $Record.Folder

    $containerCount = @(
        $hasItem
        $hasAffectedItems
        $hasFolders
        $hasFolder
    ) | Where-Object { $_ } | Measure-Object | Select-Object -ExpandProperty Count

    if ($containerCount -gt 1 -and -not $hasFolders -and -not $hasFolder) {
        # Audit Logs sometimes have Folders and/ Folder columns with only the folder info, for items that are already
        # covered in matching detail within the AffectedItems column, so we ignore and don't warn about those.
        # Otherwise we warn so that we can review the raw log and potentially update the script for a new and un-
        # foreseen scenario.  Only the 1 container's item(s) will be returned.
        Write-Warning (
            "Multiple item containers found for operation '{0}' (type '{1}', record ID '{2}'.  Only the 1 container's item(s) will be returned." -f
            $Record.Operation,
            (Convert-RecordType $Record.RecordType),
            $Record.AuditRecordId
        )
    }
    elseif ($containerCount -eq 0) {
        Write-Warning (
            "No item containers found for operation '{0}' (type '{1}', record ID '{2}'." -f
            $Record.Operation,
            (Convert-RecordType $Record.RecordType),
            $Record.AuditRecordId
        )
        [PSCustomObject]@{
            Item         = $null
            ParentFolder = $null
            Source       = '[NO ITEM CONTAINER FOUND]'
        }
    }

    if ($hasItem) {
        foreach ($item in @($Record.Item)) {
            [PSCustomObject]@{
                Item         = $item
                ParentFolder = $item.ParentFolder.Path
                Source       = 'Item'
            }
        }
    }
    elseif ($hasAffectedItems) {
        foreach ($item in @($Record.AffectedItems)) {
            [PSCustomObject]@{
                Item         = $item
                ParentFolder = $item.ParentFolder.Path
                Source       = 'AffectedItems'
            }
        }
    }
    elseif ($hasFolders) {
        foreach ($folder in @($Record.Folders)) {
            foreach ($item in @($folder.FolderItems)) {
                [PSCustomObject]@{
                    Item         = $item
                    ParentFolder = $folder.Path
                    Source       = 'Folders.FolderItems'
                }
            }
        }
    }
    elseif ($hasFolder) {
        [PSCustomObject]@{
            Item         = $null
            ParentFolder = $Record.Folder.Path
            Source       = 'Folder'
        }
    }
}

# Import and validate the raw UAL CSV file:
try { $inputUAL = @(Import-Csv $RawUnifiedAuditLogCSVFile -ErrorAction Stop) }
catch { throw "Unable to import '$($RawUnifiedAuditLogCSVFile)': $($_.Exception.Message)" }

if ($inputUAL.Count -eq 0) { throw "The CSV file is empty." }

$columnAliases = @{
    Identity     = @('Id', 'Identity', 'RecordId')
    CreationDate = @('CreationDate')
    RecordType   = @('RecordType')
    Operations   = @('Operation', 'Operations')
    UserIds      = @('UserId', 'UserIds')
    AuditData    = @('AuditData')
}

$inputUALProperties = @($inputUAL[0].PSObject.Properties.Name)
$resolvedColumns = @{}
$missingColumns = @()

foreach ($requiredColumn in $columnAliases.Keys) {
    $matchingColumns = @($inputUALProperties | Where-Object { $_ -in $columnAliases[$requiredColumn] })

    if ($matchingColumns.Count -eq 0) { $missingColumns += $requiredColumn }
    else { $resolvedColumns[$requiredColumn] = $matchingColumns[0] }
}

if ($missingColumns.Count -gt 0) {
    throw "CSV file is missing required columns: $($missingColumns -join ', '). Found columns: $($inputUALProperties -join ', ')"
}

$inputUAL = foreach ($row in $inputUAL) {
    [pscustomobject]@{
        Identity     = $row.($resolvedColumns['Identity'])
        AuditData    = $row.($resolvedColumns['AuditData'])
        CreationDate = $row.($resolvedColumns['CreationDate'])
        Operations   = $row.($resolvedColumns['Operations'])
        RecordType   = $row.($resolvedColumns['RecordType'])
        UserIds      = $row.($resolvedColumns['UserIds'])
    }
}

# Expand AuditData:
$records = foreach ($row in $inputUAL) {
    try { $record = ConvertFrom-Json -InputObject $row.AuditData -Depth 20 -ErrorAction Stop }
    catch {
        Write-Warning "Skipping row with Id '$($row.Id)' because the AuditData column failed to parse with ConvertFrom-Json: $($_.Exception.Message)"
        continue
    }

    # Retain CSV row's Operations value if AuditData does not contain 'Operation' (or it is empty).
    if ([string]::IsNullOrWhiteSpace([string]$record.Operation)) {
        $record | Add-Member -MemberType NoteProperty -Name Operation -Value ([string]$row.Operations)
    }

    # Retain the CSV RecordType if AuditData does not contain it.
    if ($null -eq $record.RecordType -and $null -ne $row.RecordType) {
        $record | Add-Member -MemberType NoteProperty -Name RecordType -Value $row.RecordType
    }
    $record | Add-Member -MemberType NoteProperty -Name AuditRecordId -Value $row.Identity

    $record
}

$operationsToProcess = @(
    'Create',
    'Update',
    'Copy',
    'Move',
    'Send',
    'SendAs',
    'SendOnBehalf',
    'MailItemsAccessed',
    'MoveToDeletedItems',
    'SoftDelete',
    'HardDelete'
)

$itemOperations = foreach ($record in $records) {
    if ($record.Operation -notin $operationsToProcess) {
        continue
    }

    $operation = [string]$record.Operation

    if ($operation -eq 'MailItemsAccessed') {
        $operation = "MailItemsAccessed_$($record.OperationProperties.Value)"
    }

    foreach ($entry in @(Get-ItemEntries -Record $record)) {

        $item = $entry.Item

        [PSCustomObject]@{
            DateTimeADT       = ([datetime]$record.CreationTime).ToLocalTime()
            MailboxOwnerUPN   = $record.MailboxOwnerUPN
            UserId            = $record.UserId
            Operation         = $operation
            ParentFolder      = $entry.ParentFolder
            Subject           = $item.Subject
            RecipientsCount   = if ($item.RecipientsCount) { $item.RecipientsCount } else { $null }
            AttachmentsCount  = Get-AttachmentsCount $item.Attachments
            SizeInBytes       = $item.SizeInBytes
            MoveOrCopySource  = if ($operation -in @('Copy', 'Move')) { $record.Folder.Path } else { $null }
            MoveOrCopyDest    = if ($operation -in @('Copy', 'Move')) { $record.DestFolder.Path } else { $null }
            InternetMessageId = $item.InternetMessageId
            ClientIPAddress   = $record.ClientIPAddress
            ClientInfoString  = $record.ClientInfoString
            Workload          = $record.Workload
            RecordType        = Convert-RecordType $record.RecordType
            ItemSource        = $entry.Source
            CreationTime      = $record.CreationTime
            AuditRecordId     = $record.AuditRecordId
        }
    }
}

$outputPath = "$($OutputFolder)\$($OutputCSVFileNamePrefix)_ExchangeItem-Operations.csv"

$itemOperations |
Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "Exported Exchange item operations to:"
Write-Host $outputPath

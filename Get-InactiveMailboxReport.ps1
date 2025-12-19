<#
    .SYNOPSIS
    Review key details about one or more inactive mailboxes in Exchange Online.

    .DESCRIPTION
    This script will retrieve a list of inactive mailboxes in Exchange Online, for a particular Identity, and display
    key details about each mailbox, for the purpose of helping to determine which mailbox should be removed, and which
    should be kept as inactive for recovery/restore/other purposes.

    The main reason for this script is to help identify inactive mailboxes that are no longer needed, and can be
    removed, as a form of housekeeping.

    .PARAMETER Identity
    Passthrough paramter for Get-Mailbox and Get-MailUser cmdlets.

    .PARAMETER NoGridView
    Switch to supress Out-GridView output.

    .PARAMETER GridViewOnly
    Switch to supress all output except for Out-GridView.

    .NOTES
    When using M365 Retention Policies to retain inactive mailboxes, EXO mailboxes and MailUsers will be retained for
    the effective duration of the hold(s).  This commonly leads to a build-up of redundant and/or invalid inactive
    mailboxes and/or MailUsers that can and should be removed.  This script is intended to help identify those objects.

    To force delete a soft-deleted/orphaned MailUser:
    Set-MailUser USER_ID -ExcludeFromAllOrgHolds

    To force delete an Inactive Mailbox:
    Set-Mailbox MAILBOX_ID -InactiveMailbox -ExcludeFromAllOrgHolds
#>
#Requires -Version 5.1
#Requires -Modules @{ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.1.0'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'}
[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [object[]]$Identity,

    [switch]$NoGridView,
    [switch]$GridViewOnly
)
begin {
    if (-not (Get-Command Get-Mailbox -ParameterName IncludeInactiveMailbox)) {
        throw 'This script requires an active connection to EXO PowerShell (using ExchangeOnlineManagement v3.1.0 or newer), and access to the Get-Mailbox cmdlet.'
    }
}
process {
    $RcptObjects = @()
    $Mailboxes = Get-Mailbox -Identity $Identity[0] -IncludeInactiveMailbox -ErrorAction SilentlyContinue

    $MbxStats = @{}
    foreach ($mbx in $Mailboxes) {
        $RcptObjects += $mbx
        Get-MailboxStatistics -Identity $mbx.Guid.Guid -IncludeSoftDeletedRecipients -ErrorAction SilentlyContinue | ForEach-Object {
            $MbxStats[$mbx.Guid.Guid] = [pscustomobject]@{
                PrimaryMailboxItemCount = $_.ItemCount
                PrimaryMailboxSizeMB    = ([Int64]($_.TotalItemSize.Value -replace '.*\(' -replace '\s.*' -replace ',') / 1MB) -replace '\..*'
            }
        }
    }

    $RcptObjects += Get-MailUser -Identity $Identity[0] -ErrorAction SilentlyContinue
    $RcptObjects += Get-MailUser -Identity $Identity[0] -SoftDeletedMailUser -ErrorAction SilentlyContinue

    $CommonOutputObjects = @($RcptObjects |
        Select-Object -Property ExternalDirectoryObjectId, RecipientTypeDetails, DisplayName, Name,
        Alias, UserPrincipalName, PrimarySmtpAddress,
        @{
            Name       = 'PrimaryMailboxItemCount'
            Expression = { $MbxStats[$_.Guid.Guid].PrimaryMailboxItemCount }
        },
        @{
            Name       = 'PrimaryMailboxSizeMB'
            Expression = { $MbxStats[$_.Guid.Guid].PrimaryMailboxSizeMB }
        },
        IsInactiveMailbox, IsSoftDeletedByRemove, IsSoftDeletedByDisable, IsDirSynced,
        WhenCreated, WhenMailboxCreated, WhenSoftDeleted, WhenChanged,
        Guid, ExchangeGuid, ArchiveGuid)

    Write-Debug "STOP"
    if ($CommonOutputObjects.Count -gt 0) {

        if (-not $NoGridView) {
            $CommonOutputObjects | Out-GridView -Title "Mailboxes/MailUsers for Identity '$($Identity[0])'"
        }
        if (-not $GridViewOnly) { $CommonOutputObjects }
    }
    else {
        Write-Warning "No Mailboxes nor MailUsers were found for Identity '$($Identity[0])'."
    }
}
end {}

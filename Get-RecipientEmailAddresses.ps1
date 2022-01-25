<#
    .Synopsis
    Get all email addresses present on the specified recipient(s).

    .Description
    For various reasons, it can be helpful to export email addresses to individual 'rows' or output objects.  For
    example, old email addresses that use no-longer-preset Accepted Domains preventing mailboxes from being migrated to
    Exchange Online.  This script outputs every email address onto its own line (or as its own object in PS), allowing
    for easy inspection of any and all addresses.

    .Parameter Identity
    One or more recipient ID's.  For which property to use as the ID, follow the guidance from Microsoft for
    Get-Recipient's -Identity parameter:
    - https://docs.microsoft.com/en-us/powershell/module/exchange/get-recipient

    .Parameter InputObject
    One or more objects with the EmailAddresses property included (e.g., from Get-Mailbox, Get-Recipient, etc.).

    .Parameter CompareWithAcceptedDomains
    Adds the 'IsAcceptedDomain' property with True/False value.

    .Example
    Get-Recipient | .\Get-RecipientEmailAddresses.ps1 -CompareWithAcceptedDomains

    .Example
    $Mailboxes = Get-Mailbox; $Mailboxes | .\Get-RecipientEmailAddresses.ps1 -PrefixesToIgnore x400, GWISE

    .Example
    .\Get-RecipientEmailAddresses.ps1 -Identity SalesGroup-DL@GroupsWorkToo.local

    .Example
    $Groups = Get-DistributionGroup; .\Get-RecipientEmailAddresses.ps1 -InputObject $Groups -CompareWithAcceptedDomains

    .Example
    $MbxArray | .\Get-RecipientEmailAddresses.ps1 -CompareWithAcceptedDomains -PrefixesToIgnore x400,x500 | where {$_.IsAcceptedDomain -ne $true}
#>
#Requires -PSEdition Desktop
#Requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'stringInput')]
param (
    [Parameter(
        ParameterSetName = 'stringInput',
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName
    )]
    [string[]]$Identity,

    [Parameter(
        ParameterSetName = 'objectInput',
        ValueFromPipeline
    )]
    [ValidateScript(
        {
            if ($_ | Get-Member -Name EmailAddresses) { $true } else {

                throw 'Input objects need to include the EmailAddresses property.'
            }
        }
    )]
    [object[]]$InputObject,

    [string[]]$PrefixesToIgnore,
    [switch]$CompareWithAcceptedDomains
)
begin {
    if (-not (Get-Command Get-AcceptedDomain)) {

        throw 'Command not found: Get-Accepted.  Please be connected to an Exchange remote PS session before calling this script.'
    }
    else {
        try { $Script:AcceptedDomains = Get-AcceptedDomain -ErrorAction Stop } catch { throw }
    }
}
process {
    try {
        $Script:Recipients = if ($PSCmdlet.ParameterSetName -eq 'stringInput') {

            if (-not (Get-Command Get-Recipient)) {

                throw 'Command not found: Get-Recipient.  Please be connected to an Exchange remote PS session before calling this script.'
            }

            foreach ($id in $Identity) { Get-Recipient -Identity "$($id)" -ErrorAction Stop }
        }
        else { $InputObject }
    }
    catch { throw }

    foreach ($rcpt in $Recipients) {
        foreach ($addr in $rcpt.EmailAddresses) {

            $addressObject = [PSCustomObject]@{

                PrimarySmtpAddress = $rcpt.PrimarySmtpAddress
                EmailAddress       = $addr
                Domain             = if ($addr -match '(^smtp:)|(^sip:)') { $addr -replace '.*@' } else { $null }
                Prefix             = $addr -replace '(.*):.*', '$1'
                Guid               = $rcpt.Guid.ToString()
            } |
            Where-Object {$PrefixesToIgnore -notcontains $_.Prefix}

            if ($CompareWithAcceptedDomains) {

                $addressObject |
                Add-Member -NotePropertyName IsAcceptedDomain -NotePropertyValue ($Script:AcceptedDomains.DomainName -contains $addressObject.Domain)
            }

            $addressObject
        }
    }
}

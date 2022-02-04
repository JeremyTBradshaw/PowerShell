<#
    .Synopsis
    !!!! Script is in development, not ready for use !!!!
    Helper script for EXO Inactive Mailbox recovery/restore, or no recovery/restore, but rather net-new mailbox.

    .Description
    This script assumes a Hybrid Exchange environment and that the on-premises user is a *Remote* mailbox recipient
    object.  It also requires that the on-premises user is currently *NOT* synced to Azure AD via AAD Connect, and that
    there is no user in Azure AD, not even in the Recycle Bin, with the ImmutableId property matching the on-premises
    user's ObjectGuid.

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/blob/main/Restore-InactiveMailbox.ps1

    .Parameter Recover
    Switch parameter to specify we want to recover the mailbox.

    .Parameter Restore
    Switch parameter to specify we want to restore the mailbox.  Indicates that a destination mailbox for the restore
    exists and is ready to receive the restored data.

    .Parameter ExcludeFolders
    Essentially a passthrough parameter for New-MailboxRestoreRequest.  Allows for excluding certain folders from the
    restore.  The input for this parameter isn't checked, so be sure to review the MS Docs article for
    New-MailboxRestoreRequest, particularly the -ExcludeFolders parameter:
    https://docs.microsoft.com/en-us/powershell/module/exchange/new-mailboxrestorerequest

    .Parameter TargetRootFolder
    Another passthrough parameter for New-MailboxRestoreRequest.  Again, follow the logic of MS Docs article:
    https://docs.microsoft.com/en-us/powershell/module/exchange/new-mailboxrestorerequest

    .Parameter NetNewMailbox
    Switch parameter to specify we neither want to recover nor restore an existing Inactive Mailbox, rather want to
    enable the user with a new, empty mailbox.

    .Parameter AdminUPN
    Optional parameter which serves as a passthrough for Connect-ExchangeOnline's -UserPrincipalName parameter.
    Supplying this does well to prevent re-prompting to authenticate if there's an existing token cached for the
    specified UserPrincipalName.
#>
#Requires -Version 5.1
#Requires -Modules @{ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '2.0.5'; Guid = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'}
[CmdletBinding(
    DefaultParameterSetName = 'Recover',
    SupportsShouldProcess,
    ConfirmImpact = 'High'
)]
param (
    [Parameter(ParameterSetName = 'Recover')]
    [switch]$Recover,

    [Parameter(ParameterSetName = 'Restore')]
    [switch]$Restore,

    [Parameter(ParameterSetName = 'Restore')]
    [string]$TargetRootFolder,

    [Parameter(ParameterSetName = 'NetNewMailbox')]
    [switch]$NetNewMailbox,

    [string]$AdminUPN
)

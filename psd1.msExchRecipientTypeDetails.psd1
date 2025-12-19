<#
    .Synopsis
    msExchRecipientTypeDetails to RecipientTypeDetails lookup table.

    .Notes
    Sources:
      - https://answers.microsoft.com/en-us/msoffice/forum/msoffice_o365admin-mso_exchon-mso_o365b/recipient-type-values/7c2620e5-9870-48ba-b5c2-7772c739c651
      - https://www.undocumented-features.com/2020/05/06/every-last-msexchrecipientdisplaytype-and-msexchrecipienttypedetails-value/
#>
@{
    0               = 'None'
    1               = 'UserMailbox'
    2               = 'LinkedMailbox'
    4               = 'SharedMailbox'
    8               = 'LegacyMailbox'
    16              = 'RoomMailbox'
    32              = 'EquipmentMailbox'
    64              = 'MailContact'
    128             = 'MailUser'
    256             = 'MailUniversalDistributionGroup'
    512             = 'MailNonUniversalGroup'
    1024            = 'MailUniversalSecurityGroup'
    2048            = 'DynamicDistributionGroup'
    4096            = 'PublicFolder'
    8192            = 'SystemAttendantMailbox'
    16384           = 'SystemMailbox'
    32768           = 'MailForestContact'
    65536           = 'User'
    131072          = 'Contact'
    262144          = 'UniversalDistributionGroup'
    524288          = 'UniversalSecurityGroup'
    1048576         = 'NonUniversalGroup'
    2097152         = 'Disable User'
    4194304         = 'MicrosoftExchange'
    8388608         = 'ArbitrationMailbox'
    16777216        = 'MailboxPlan'
    33554432        = 'LinkedUser'
    268435456       = 'RoomList'
    536870912       = 'DiscoveryMailbox'
    1073741824      = 'RoleGroup'
    2147483648      = 'RemoteUserMailbox'
    4294967296      = 'Computer'
    8589934592      = 'RemoteRoomMailbox'
    17179869184     = 'RemoteEquipmentMailbox'
    34359738368     = 'RemoteSharedMailbox'
    68719476736     = 'PublicFolderMailbox'
    137438953472    = 'Team Mailbox'
    274877906944    = 'RemoteTeamMailbox'
    549755813888    = 'MonitoringMailbox'
    1099511627776   = 'GroupMailbox'
    2199023255552   = 'LinkedRoomMailbox'
    4398046511104   = 'AuditLogMailbox'
    8796093022208   = 'RemoteGroupMailbox'
    17592186044416  = 'SchedulingMailbox'
    35184372088832  = 'GuestMailUser'
    70368744177664  = 'AuxAuditLogMailbox'
    140737488355328 = 'SupervisoryReviewPolicyMailbox'
}

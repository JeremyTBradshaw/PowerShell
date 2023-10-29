<#
    .SYNOPSIS
    Using MS Graph Advanced Hunting API, find messages which fail authentication, pulling back pertinent details.

    .NOTES
    Script is incomplete / not ready / draft, as of 2023-10-29.
#>
#Requires -Modules @{ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.8.0'; Guid = '883916f2-9184-46ee-b1f8-b6a2fb784cee'}
#Requires -Modules @{ModuleName = 'Microsoft.Graph.Mail'; ModuleVersion = '2.8.0'; Guid = '6e4d36b5-7ff2-454b-8572-674b3ab0362b'}
#Requires -Modules @{ModuleName = 'Microsoft.Graph.Security'; ModuleVersion = '2.8.0'; Guid = '06b0769e-2c63-4d60-9fb4-9ca0ec87e0d7'}
[CmdletBinding()]
param ()

$kql_EmailEvents = @"
EmailEvents
| where EmailDirection == "Inbound"
| extend DMARC = parse_json(AuthenticationDetails).DMARC
| where DMARC =~ 'fail'
| extend SPF = parse_json(AuthenticationDetails).SPF
| extend DKIM = parse_json(AuthenticationDetails).DKIM
| extend CompAuth = parse_json(AuthenticationDetails).CompAuth
| project Timestamp, EmailDirection,RecipientEmailAddress, SenderDisplayName, SenderFromAddress, SenderFromDomain, SenderMailFromDomain, SenderIPv4,
DeliveryAction, SPF, DKIM, DMARC, CompAuth, BCL = BulkComplaintLevel, SCL = parse_json(ConfidenceLevel).Spam,PCL = parse_json(ConfidenceLevel).Phish,
EmailAction, SpamDetectionMethod = parse_json(DetectionMethods).Spam, PhishDetectionMethod = parse_json(DetectionMethods).Phish,
RecipientObjectId, AdditionalFields, InternetMessageId
"@

$messages = Start-MgSecurityHuntingQuery -Query $kql_EmailEvents

<#
We end up with this:
    $messages.Results[0].AdditionalProperties
        Key                   Value
        ---                   -----
        Timestamp             2023-10-29T18:58:55Z
        EmailDirection        Inbound
        RecipientEmailAddress internaluser@demo_x12345.onmicrosoft.com
        SenderDisplayName     Bradshaw, Jeremy
        SenderFromAddress     jeremy.bradshaw@contoso.com
        SenderFromDomain      contoso.com
        SenderMailFromDomain  mailerDaemon02.contoso.com
        SenderIPv4            192.168.0.218
        DeliveryAction        Junked
        SPF                   fail
        DKIM                  none
        DMARC                 fail
        CompAuth              pass
        EmailAction           Moved to Junk folder
        RecipientObjectId     f182bdcb-dbd9-4186-b097-a70dbf4f2baf
        AdditionalFields
#>

$output = foreach ($msg in $messages.Results) {

    $mailMsg = Get-MgUserMessage -UserId $msg.AdditionalProperties.RecipientObjectId -Filter "internetMessageId eq '$($msgAddtionalProperties.InternetMessageId)'" -Property InternetMessageHeaders
    $msg.AdditionalProperties | Select-Object -Property *, @{Name = 'InternetMessageHeaders'; Expression = { $mailMsg.InternetMessageHeaders } }
}


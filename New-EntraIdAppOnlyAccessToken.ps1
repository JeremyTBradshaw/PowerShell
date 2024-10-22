using namespace System
using namespace System.Management.Automation.Host
using namespace System.Runtime.InteropServices
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ if (([guid]::TryParse($_, [ref]([guid]::NewGuid()))) -or ($_ -like '*.onmicrosoft.com')) { $true } else { throw "Invalid TenantId: $($_)" } })]
    [Object]$TenantId,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ if ([guid]::TryParse($_, [ref]([guid]::NewGuid()))) { $true } else { throw "Invalid ApplicationId: $($_)" } })]
    [Object]$ApplicationId,

    [Parameter(Mandatory = $true, ParameterSetName = 'Certificate')]
    [X509Certificate]$Certificate,

    [Parameter(Mandatory = $true, ParameterSetName = 'ClientSecret')]
    [ValidatePattern('^[a-zA-Z0-9_\-\.\~]{1,40}$')]
    [string]$ClientSecret,

    [Parameter(HelpMessage = 'Default scope is for Microsoft Graph (https://graph.microsoft.com/.default).  Other supported scopes are: EWS, IMAP, POP, and SMTP')]
    [ValidateSet('EWS', 'SMTP', 'IMAP', 'POP')]
    [string]$AltScope
)

function ConvertTo-Base64Url {
    param (
        [ValidatePattern('^(?:[A-Za-z0-9+\/]{4})*(?:[A-Za-z0-9+\/]{2}==|[A-Za-z0-9+\/]{3}=|[A-Za-z0-9+\/]{4})$')]
        [string[]]$String
    )
    $String -replace '\+', '-' -replace '/', '_' -replace '='
}

try {
    $Scope = switch ($AltScope) {
        default {
            # https://learn.microsoft.com/en-us/entra/identity-platform/scopes-oidc#the-default-scope
            'https://graph.microsoft.com/.default'
        }
        { @('EWS', 'IMAP', 'POP', 'SMTP') -contains $_ } {
            # https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth#get-a-token-with-app-only-auth
            # https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#use-client-credentials-grant-flow-to-authenticate-smtp-imap-and-pop-connections
            'https://outlook.office365.com/.default'
        }
    }
    $trBody = @{
        client_id  = $ApplicationId
        scope      = $Scope
        grant_type = "client_credentials"
    }
    $trParams = @{
        Method      = 'POST'
        Uri         = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        ContentType = 'application/x-www-form-urlencoded'
        UserAgent   = 'PowerShell'
        ErrorAction = 'Stop'
    }

    if ($PSCmdlet.ParameterSetName -eq 'Certificate') {
        $NowUTC = [datetime]::UtcNow
        $EncodedHeader = [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes(
                (
                    ConvertTo-Json -InputObject (
                        @{
                            alg = 'RS256'
                            typ = 'JWT'
                            x5t = ConvertTo-Base64Url -String ([Convert]::ToBase64String($Certificate.GetCertHash()))
                        }
                    )
                )
            )
        )
        $EncodedPayload = [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes(
                (
                    ConvertTo-Json -InputObject (
                        @{
                            aud = "https://login.microsoftonline.com/$TenantId/oauth2/token"
                            exp = (Get-Date $NowUTC.AddMinutes(5) -UFormat '%s') -replace '\..*'
                            iss = $ApplicationId
                            jti = [Guid]::NewGuid()
                            nbf = (Get-Date $NowUTC -UFormat '%s') -replace '\..*'
                            sub = $ApplicationId
                        }
                    )
                )
            )
        )
        $JWT = (ConvertTo-Base64Url -String $EncodedHeader, $EncodedPayload) -join '.'
        $Signature = ConvertTo-Base64Url -String (
            [Convert]::ToBase64String(
                $Script:Certificate.PrivateKey.SignData(
                    [Text.Encoding]::UTF8.GetBytes($JWT),
                    [HashAlgorithmName]::SHA256,
                    [RSASignaturePadding]::Pkcs1
                )
            )
        )
        $ClientAssertion = $JWT + '.' + $Signature
        $trBody['client_assertion'] = $ClientAssertion
        $trBody['client_assertion_type'] = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        $trParams['Headers'] = @{ Authorization = "Bearer $($ClientAssertion)" }
    }
    else { $trBody['client_secret'] = $ClientSecret }
    $trParams['Body'] = $trBody

    Invoke-RestMethod @trParams
}
catch { throw }

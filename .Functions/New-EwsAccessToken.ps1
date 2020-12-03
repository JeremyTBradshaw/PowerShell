#Requires -Version 5.1
using Namespace System.Security.Cryptography.X509Certificates
using Namespace System.Management.Automation.Host

<# v0.0.0 (incomplete and unpublished) #>

function New-EwsAccessToken {

    [CmdletBinding(
        DefaultParameterSetName = 'Certificate'
    )]
    param (
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [Alias('ClientId')]
        [Guid]$ApplicationId,

        [Parameter(
            Mandatory,
            ParameterSetName = 'Certificate',
            HelpMessage = 'E.g. Use $Certificate, where `$Certificate = Get-ChildItem cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [X509Certificate2]$Certificate,

        [Parameter(
            Mandatory,
            ParameterSetName = 'CertificateStorePath',
            HelpMessage = 'E.g. cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317; E.g. cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [ValidateScript(
            {
                if (Test-Path -Path $_) { $true } else {
                
                    throw "An example proper path would be 'cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'."
                }
            }
        )]
        [string]$CertificateStorePath,

        [ValidateRange(1, 10)]
        [int16]$JWTExpMinutes = 2
    )

    if ($PSCmdlet.ParameterSetName -eq 'CertificateStorePath') {

        try {
            $Script:Certificate = Get-ChildItem -Path $CertificateStorePath -ErrorAction Stop
        }
        catch { throw $_ }
    }
    else { $Script:Certificate = $Certificate }

    if (-not (Test-CertificateProvider -Certificate $Script:Certificate)) {

        $ErrorMessage = "The supplied certificate does not use the provider 'Microsoft Enhanced RSA and AES Cryptographic Provider'.  " +
        "For best luck, use a certificate generated using New-SelfSignedMSGraphApplicationCertificate."

        throw $ErrorMessage
    }

    $NowUTC = [datetime]::UtcNow

    $JWTHeader = @{

        alg = 'RS256'
        typ = 'JWT'
        x5t = ConvertTo-Base64UrlFriendly -String ([System.Convert]::ToBase64String($Script:Certificate.GetCertHash()))
    }

    $JWTClaims = @{

        aud = "https://login.microsoftonline.com/$TenantId/oauth2/token"
        exp = (Get-Date $NowUTC.AddMinutes($JWTExpMinutes) -UFormat '%s') -replace '\..*'
        iss = $ApplicationId.Guid
        jti = [Guid]::NewGuid()
        nbf = (Get-Date $NowUTC -UFormat '%s') -replace '\..*'
        sub = $ApplicationId.Guid
    }

    $EncodedJWTHeader = [System.Convert]::ToBase64String(
        
        [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $JWTHeader))
    )
    
    $EncodedJWTClaims = [System.Convert]::ToBase64String(
        
        [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $JWTClaims))
    )

    $JWT = ConvertTo-Base64UrlFriendly -String ($EncodedJWTHeader + '.' + $EncodedJWTClaims)

    $Signature = ConvertTo-Base64UrlFriendly -String ([System.Convert]::ToBase64String(
        
            $Script:Certificate.PrivateKey.SignData(
            
                [System.Text.Encoding]::UTF8.GetBytes($JWT),
                [Security.Cryptography.HashAlgorithmName]::SHA256,
                [Security.Cryptography.RSASignaturePadding]::Pkcs1
            )
        )
    )

    $JWT = $JWT + '.' + $Signature

    $Body = @{

        client_id             = $ApplicationId
        client_assertion      = $JWT
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        scope                 = 'https://outlook.office365.com/.default'
        grant_type            = "client_credentials"
    }

    $TokenRequestParams = @{

        Method      = 'POST'
        Uri         = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        Body        = $Body
        Headers     = @{ Authorization = "Bearer $($JWT)" }
        ContentType = 'application/x-www-form-urlencoded'
        ErrorAction = 'Stop'
    }

    try {
        Invoke-RestMethod @TokenRequestParams
    }
    catch { throw $_ }
}

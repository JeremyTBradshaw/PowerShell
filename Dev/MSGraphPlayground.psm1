#Requires -Version 5.1
using namespace System
using namespace System.Management.Automation.Host
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates

# ParamterSetName = 'DeviceCode'
function New-DeviceCodeAccessToken {

    [CmdletBinding()]
    param (
        [ValidateSet('Common', 'Consumers', 'Organizations')]
        [string]$Endpoint = 'Common',

        [Parameter(Mandatory)]
        [Alias('ClientId')]
        [Guid]$ApplicationId,

        [string[]]$Scopes
    )

    try {
        $dcrBody = @(
            "client_id=$($ApplicationId)",
            "scope=$($Scopes -join ' ')"
        ) -join '&'
    
        $dcrParams = @{
    
            Method      = 'POST'
            Uri         = "https://login.microsoftonline.com/$($Endpoint)/oauth2/v2.0/devicecode"
            Body        = $dcrBody
            ContentType = 'application/x-www-form-urlencoded'
            ErrorAction = 'Stop'
        }
        $dcrResponse = Invoke-RestMethod @dcrParams
    }
    catch { throw $_ }
    
    $dtNow = [datetime]::Now
    $sw1 = [Diagnostics.Stopwatch]::StartNew()
    $dcExpiration = "$($dtNow.AddSeconds($dcrResponse.expires_in).ToString('yyyy-MM-dd hh:mm:ss tt'))"

    $trBody = @(
        "grant_type=urn:ietf:params:oauth:grant-type:device_code",
        "client_id=$($ApplicationId)",
        "device_code=$($dcrResponse.device_code)"
    ) -join '&'

    # Wait for user to enter code before starting to poll token endpoint:
    switch (
        $host.UI.PromptForChoice(
    
            "Authorization started (expires at $($dcExpiration)",
            "$($dcrResponse.message)",
            [ChoiceDescription]('&Done'),
            0
        )
    ) { 0 { <##> } }

    Write-Debug 'Inspect $dcrResponse.'
    if ($sw1.Elapsed.Minutes -lt 15) {

        $sw2 = [Diagnostics.Stopwatch]::StartNew()
        $successfulResponse = $false
        $pollCount = 0
        do {
            if ($sw2.Elapsed.Seconds -ge $dcrResponse.interval) {

                $sw2.Restart()
                $pollCount++

                try {
                    $trParams = @{
    
                        Method      = 'POST'
                        Uri         = "https://login.microsoftonline.com/$($Endpoint)/oauth2/v2.0/token"
                        Body        = $trBody
                        ContentType = 'application/x-www-form-urlencoded'
                        ErrorAction = 'Stop'
                    }
                    $trResponse = Invoke-RestMethod @trParams
                    $successfulResponse = $true
                }
                catch {
                    if ($_.ErrorDetails.Message) {
                        
                        $badResponse = ConvertFrom-Json -InputObject $_.ErrorDetails.Message
    
                        if ($badResponse.error -eq 'authorization_pending') {
                        
                            if ($pollCount -eq 1) {

                                "The user hasn't finished authenticating, but hasn't canceled the flow (error: authorization_pending).  " +
                                "Continuing to poll the token endpoint at the requested interval ($($dcrResponse.interval) seconds)." |
                                Write-Warning
                            }
                        }
                        elseif ($badResponse.error -match '^(authorization_declined)|(bad_verification_code)|(expired_token)$') {

                            # https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code#expected-errors
                            throw "Authorization failed due to foreseeable error: $($badResponse.error)."
                        }
                        else {
                            Write-Warning 'Authorization failed due to an unexpected error.'
                            throw $badResponse.error_description
                        }
                    }
                    else {
                        Write-Warning 'An error was encountered with the Invoke-RestMethod command.  Authorization request did not complete.'
                        throw $_
                    }
                }
            }
            if (-not $successfulResponse) { Start-Sleep -Seconds 1 }
        }
        while ($sw1.Elapsed.Minutes -lt 15 -and -not $successfulResponse)

        # Output the token request response:
        $trResponse
    }
    else {
        throw "Authorization request expired at $($dcExpiration), please try again."
    }
}

# ParameterSetName = 'RefreshToken'
function Get-RefreshedAcessToken {
    [CmdletBinding()]
    param (
        [ValidateSet('Common', 'Consumers', 'Organizations')]
        [string]$Endpoint = 'Common',

        [Parameter(Mandatory)]
        [Alias('ClientId')]
        [Guid]$ApplicationId,

        [Parameter(
            Mandatory,
            HelpMessage = 'Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken -DeviceCodeFlow'
        )]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and
                    $_.scope -and
                    $_.access_token -match '^[-\w]+\.[-\w]+\.[-\w]+$' -and
                    $_.refreh_token -match '^[-\w]+\.[-\w]+\.[-\w]+$') { $true } else {

                    throw 'Invalid token object.  Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken ...'
                }
            }
        )]
        [Object]$TokenObject
    )

    $trBody = @(
        "grant_type=refresh_token",
        "refresh_token=$($TokenObject.refresh_token)"
    ) -join '&'

    $trParams = @{
    
        Method      = 'POST'
        Uri         = "https://login.microsoftonline.com/$($Endpoint)/oauth2/v2.0/token"
        Body        = $trBody
        ContentType = 'application/x-www-form-urlencoded'
        ErrorAction = 'Stop'
    }
    $trResponse = Invoke-RestMethod @trParams

    # Output the token request response:
    $trResponse
}

# ParameterSetName = 'CertificateCredentials'
function New-AppOnlyAccessToken {

    [CmdletBinding(
        DefaultParameterSetName = 'Certificate'
    )]
    param (
        [Parameter(Mandatory)]
        [Alias('Tenant', 'TenantName', 'TenantDomain', 'TenantDomainName')]
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
        x5t = ConvertTo-Base64UrlFriendly -String ([Convert]::ToBase64String($Script:Certificate.GetCertHash()))
    }

    $JWTClaims = @{

        aud = "https://login.microsoftonline.com/$TenantId/oauth2/token"
        exp = (Get-Date $NowUTC.AddMinutes($JWTExpMinutes) -UFormat '%s') -replace '\..*'
        iss = $ApplicationId.Guid
        jti = [Guid]::NewGuid()
        nbf = (Get-Date $NowUTC -UFormat '%s') -replace '\..*'
        sub = $ApplicationId.Guid
    }

    $EncodedJWTHeader = [Convert]::ToBase64String(
        
        [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $JWTHeader))
    )
    
    $EncodedJWTClaims = [Convert]::ToBase64String(
        
        [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $JWTClaims))
    )

    $JWT = ConvertTo-Base64UrlFriendly -String ($EncodedJWTHeader + '.' + $EncodedJWTClaims)

    $Signature = ConvertTo-Base64UrlFriendly -String ([Convert]::ToBase64String(
        
            $Script:Certificate.PrivateKey.SignData(
            
                [Text.Encoding]::UTF8.GetBytes($JWT),
                [HashAlgorithmName]::SHA256,
                [RSASignaturePadding]::Pkcs1
            )
        )
    )

    $JWT = $JWT + '.' + $Signature

    $trBody = @{

        client_id             = $ApplicationId
        client_assertion      = $JWT
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        scope                 = 'https://graph.microsoft.com/.default'
        grant_type            = "client_credentials"
    }

    $trParams = @{

        Method      = 'POST'
        Uri         = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        Body        = $trBody
        Headers     = @{ Authorization = "Bearer $($JWT)" }
        ContentType = 'application/x-www-form-urlencoded'
        ErrorAction = 'Stop'
    }

    $trResponse = Invoke-RestMethod @trParams

    # Output the token request response:
    $trResponse
}

function Test-CertificateProvider ([X509Certificate2]$Certificate) {

    if ($PSVersionTable.PSEdition -eq 'Desktop') {

        $certProvider = $Certificate.PrivateKey.CspKeyContainerInfo.ProviderName
    }
    else { $certProvider = $Certificate.PrivateKey.Key.Provider }

    if ($certProvider -eq 'Microsoft Enhanced RSA and AES Cryptographic Provider') {

        $true
    }
    else { $false }
}

function ConvertTo-Base64Url ([string]$String) {

    $String -replace '\+', '-' -replace '/', '_' -replace '='
}

function ConvertFrom-Base64Url ([string]$String) {

    while ($String.Length % 4) { $String += '=' }
    $String -replace '-', '\+' -replace '_', '/'
}

function ConvertFrom-JWT {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_ -match '^eyJ0[-\w]+\.[-\w]+\.[-\w]+$') { $true } else { throw 'Invalid JWT.' }
            }
        )]
        [Object]$JWT
    )

    $Headers, $Claims = ($JWT -split '\.')[0,1]
    
    [PSCustomObject]@{

        Headers = [System.Text.Encoding]::ASCII.GetString(

            [Convert]::FromBase64String((ConvertFrom-Base64Url $Headers))
        ) |
        ConvertFrom-Json

        Claims = [System.Text.Encoding]::ASCII.GetString(
        
            [Convert]::FromBase64String((ConvertFrom-Base64Url $Claims))
        ) |
        ConvertFrom-Json
    }
}

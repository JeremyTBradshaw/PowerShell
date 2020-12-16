#Requires -Version 5.1
using namespace System
using namespace System.Management.Automation.Host
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates

function New-MSGraphAccessToken {

    [CmdletBinding(
        DefaultParameterSetName = 'DeviceCode_Endpoint'
    )]
    param (
        [Parameter(Mandatory, ParameterSetName = 'ClientCredentials_Certificate')]
        [Parameter(Mandatory, ParameterSetName = 'ClientCredentials_CertificateStorePath')]
        [Parameter(Mandatory, ParameterSetName = 'DeviceCode_TenantId')]
        [Parameter(Mandatory, ParameterSetName = 'RefreshToken_TenantId')]
        [string]$TenantId, # Guid / FQDN

        [Parameter(Mandatory)]
        [Guid]$ApplicationId,

        [Parameter(
            Mandatory,
            ParameterSetName = 'ClientCredentials_Certificate',
            HelpMessage = 'E.g. Use $Certificate, where `$Certificate = Get-ChildItem cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [X509Certificate2]$Certificate,

        [Parameter(
            Mandatory,
            ParameterSetName = 'ClientCredentials_CertificateStorePath',
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

        [Parameter(ParameterSetName = 'ClientCredentials_Certificate')]
        [Parameter(ParameterSetName = 'ClientCredentials_CertificateStorePath')]
        [ValidateRange(1, 10)]
        [int16]$JWTExpMinutes = 2,

        [Parameter(ParameterSetName = 'DeviceCode_Endpoint')]
        [Parameter(ParameterSetName = 'RefreshToken_Endpoint')]
        [ValidateSet('Common', 'Consumers', 'Organizations')]
        [string]$Endpoint = 'Common',

        [Parameter(Mandatory, ParameterSetName = 'DeviceCode_TenantId')]
        [Parameter(Mandatory, ParameterSetName = 'DeviceCode_Endpoint')]
        [Parameter(ParameterSetName = 'RefreshToken_TenantId')]
        [Parameter(ParameterSetName = 'RefreshToken_Endpoint')]
        [string[]]$Scopes,
    
        [Parameter(Mandatory, ParameterSetName = 'RefreshToken_TenantId')]
        [Parameter(Mandatory, ParameterSetName = 'RefreshToken_Endpoint')]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and $_.refresh_token) { $true } else {

                    throw 'Invalid access token.  Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken -Scopes offline_access...'
                }
            }
        )]
        [Object]$RefreshToken
    )

    #region Initialization
    if ($PSCmdlet.ParameterSetName -eq 'ClientCredentials_CertificateStorePath') {
        try {
            $Script:Certificate = Get-ChildItem -Path $CertificateStorePath -ErrorAction Stop
        }
        catch { throw $_ }
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'ClientCredentials_Certificate') {
        
        $Script:Certificate = $Certificate
    }

    if ($PSCmdlet.ParameterSetName -like 'ClientCredentials_*') {

        if (-not (Test-SigningCertificate -Certificate $Script:Certificate)) {

            throw = "The supplied certificate must use the provider 'Microsoft Enhanced RSA and AES Cryptographic Provider', " +
            'and the SHA-256 hashing algorithm.  ' +
            'For best luck, use a certificate generated using New-SelfSignedMSGraphApplicationCertificate.'
        }
    }

    if ($PSCmdlet.ParameterSetName -like '*_TenantId') { $Script:Endpoint = $TenantId } else { $Script:Endpoint = $Endpoint }
    #endregion Initialization

    #region Functions
    function New-AppOnlyAccessToken ($TenantId, $ApplicationId, $Certificate, $JWTExpMinutes) {
        try {
            $NowUTC = [datetime]::UtcNow

            $EncodedHeader = [Convert]::ToBase64String(
                [Text.Encoding]::UTF8.GetBytes(
                    (ConvertTo-Json -InputObject (
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
                    (ConvertTo-Json -InputObject (
                            @{
                                aud = "https://login.microsoftonline.com/$TenantId/oauth2/token"
                                exp = (Get-Date $NowUTC.AddMinutes($JWTExpMinutes) -UFormat '%s') -replace '\..*'
                                iss = $ApplicationId.Guid
                                jti = [Guid]::NewGuid()
                                nbf = (Get-Date $NowUTC -UFormat '%s') -replace '\..*'
                                sub = $ApplicationId.Guid
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
        catch { throw $_ }
    }

    function New-DeviceCodeAccessToken ($Endpoint, $ApplicationId, $Scopes) {
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
        catch { throw $_ }
    }

    function Get-RefreshedAcessToken ($Endpoint, $ApplicationId, $RefreshToken, $Scopes) {
        try {
            $trBody = @(
                "client_id=$($ApplicationId)",
                'grant_type=refresh_token',
                "refresh_token=$($RefreshToken.refresh_token)"
            )
            if ($Scopes) { $trBody += "scope=$($Scopes -join ' ')" }
            
            $trBody = $trBody -join '&'
    
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
        catch { throw $_ }
    }
    #endregion Functions

    #region Main
    switch -Wildcard ($PSCmdlet.ParameterSetName) {

        'ClientCredentials_*' {

            New-AppOnlyAccessToken $TenantId $ApplicationId $Script:Certificate $JWTExpMinutes
        }

        'DeviceCode_*' {

            New-DeviceCodeAccessToken $Script:Endpoint $ApplicationId $Scopes
        }

        'RefreshToken_*' {

            Get-RefreshedAcessToken $Script:Endpoint $ApplicationId $RefreshToken $Scopes
        }
    }
    #endregion Main
}

function Test-SigningCertificate ([X509Certificate2]$Certificate) {

    if ($PSVersionTable.PSEdition -eq 'Desktop') {

        $Provider = $Certificate.PrivateKey.CspKeyContainerInfo.ProviderName
    }
    else { $Provider = $Certificate.PrivateKey.Key.Provider }

    if (
        $Provider -eq 'Microsoft Enhanced RSA and AES Cryptographic Provider' -and
        $Certificate.SignatureAlgorithm.FriendlyName -match '(sha256)'
    ) {
        $true
    }
    else { $false }
}

function ConvertTo-Base64Url {
    param (
        [ValidatePattern('^(?:[A-Za-z0-9+\/]{4})*(?:[A-Za-z0-9+\/]{2}==|[A-Za-z0-9+\/]{3}=|[A-Za-z0-9+\/]{4})$')]
        [string[]]$String
    )

    $String -replace '\+', '-' -replace '/', '_' -replace '='
}
function ConvertFrom-Base64Url ([string[]]$String) {

    foreach ($s in $String) {

        while ($s.Length % 4) { $s += '=' }
        $s -replace '-', '\+' -replace '_', '/'
    }
}

function ConvertFrom-JWTAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_ -match '^eyJ[-\w]+\.[-\w]+\.[-\w]+$') { $true } else { throw 'Invalid JWT.' }
            }
        )]
        [Object]$JWT
    )

    $Headers, $Claims = ($JWT -split '\.')[0, 1]

    [PSCustomObject]@{
        Headers = ConvertFrom-Json (
            [Text.Encoding]::ASCII.GetString(
                [Convert]::FromBase64String((ConvertFrom-Base64Url $Headers))
            )
        )
        Payload = ConvertFrom-Json(
            [Text.Encoding]::ASCII.GetString(
                [Convert]::FromBase64String((ConvertFrom-Base64Url $Claims))
            )
        )
    }
}

function New-MSGraphRequest {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Request,

        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and $_.access_token) { $true } else {

                    throw 'Invalid access token.  Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken ...'
                }
            }
        )]
        [Object]$AccessToken,

        [Alias('API', 'Version', 'Endpoint')]
        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion = 'v1.0',

        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')]
        [string]$Method = 'GET',

        [string]$Body,

        [ValidateSet('Warn', 'Inquire', 'Continue', 'SilentlyContinue')]
        [string]$nextLinkAction = 'Warn'
    )

    try {
        $RequestParams = @{

            Headers     = @{ Authorization = "Bearer $($AccessToken.access_token)" }
            Uri         = "https://graph.microsoft.com/$($ApiVersion)/$($Request)"
            Method      = $Method
            ContentType = 'application/json'
            ErrorAction = 'Stop'
        }
    
        if ($PSBoundParameters.ContainsKey('Body')) {
    
            if ($Method -notmatch '(POST)|(PATCH)') {
    
                throw "Body is not allowed when the method is $($Method), only POST or PATCH."
            }
            else { $RequestParams['Body'] = $Body }
        }
    
        Invoke-RestMethod @RequestParams -OutVariable requestResponse
    
        if ($requestResponse.'@odata.nextLink') {
    
            $Script:Continue = $true
    
            switch ($nextLinkAction) {
    
                Warn {
                    Write-Warning -Message "There are more results available. Next page: $($requestResponse.'@odata.nextLink')"
                    $Script:Continue = $false
                }
                Continue {
                    Write-Information -MessageData 'There are more results available.  Getting the next page...' -InformationAction Continue
    
                }
                Inquire {
                    switch (
                        $host.UI.PromptForChoice(
    
                            'There are more results available (i.e. response included @odata.nextLink).',
                            'Get more results?',
                            [ChoiceDescription[]]@('&Yes', 'Yes to &All', '&No'),
                            2
                        )
                    ) {
                        0 { <# Will prompt for choice again if the next response includes another @odata.nextLink.#> }
                        1 { $nextLinkAction = 'SilentlyContinue' }
                        2 { $Script:Continue = $false }
                    }
                }
            }
    
            if ($Script:Continue) {
    
                $nextLinkRequestParams = @{
                    AccessToken    = $AccessToken
                    ApiVersion     = $ApiVersion
                    Request        = "$($requestResponse.'@odata.nextLink' -replace 'https://graph.microsoft.com/(v1\.0|beta)/')"
                    nextLinkAction = $nextLinkAction
                    ErrorAction    = 'Stop'
                }
    
                New-MSGraphRequest @nextLinkRequestParams
            }
        }
    }
    catch { throw $_ }
}

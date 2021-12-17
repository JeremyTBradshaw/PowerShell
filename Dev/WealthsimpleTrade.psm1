#Requires -Version 7.2.1
#Requires -PSEdition Core
using namespace System
using namespace System.Management.Automation.Host
using namespace System.Runtime.InteropServices
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates

<# Release Notes for v0.0.0 (2021-12-17):

    - Just starting to work on this now.  It's not ready for reliable use.  Looking at porting all of the examples here:
    https://github.com/MarkGalloway/wealthsimple-trade/blob/master/API.md
    - Enforcing option for OTP (i.e., MFA) to keep it somewhat safe.
#>

#======#---------------------------#
#region# Main (exported) Functions #
#======#---------------------------#

function Connect-WealthsimpleTrade {
    [CmdletBinding()]
    param (
        [PSCredential]$Credential,
        [Uri]$UriOverride
    )

    $loginBody = if ($PSBoundParameters.ContainsKey('Credential')) {
        @{
            email    = $Credential.UserName
            password = ConvertFrom-SecureStringToPlainText $Credential.Password
        }
    }
    else {
        @{
            email    = 'Enter username (email)'
            password = Read-Host -AsSecureString -Prompt 'Enter password'
        }
    }
    $loginBody['otp'] = Read-Host -Prompt 'Enter passcode'

    $reqParams = @{
        Method                  = 'POST'
        ContentType             = 'application/json'
        Body                    = ConvertTo-Json -InputObject $loginBody
        ResponseHeadersVariable = 'responseHeaders'
    }

    $reqParams['Uri'] = if ($PSBoundParameters.ContainsKey('UriOverride')) {

        $UriOverride
    }
    else { 'https://trade-service.wealthsimple.com/auth/login' }

    $response = Invoke-RestMethod @reqParams

    $Global:WSTokens = [PSCustomObject]@{

        responseHeaders = $responseHeaders
        responseBody    = $response
    }
}

function Update-WealthsimpleTokens {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [Object]$TokenObject,
        [Uri]$UriOverride
    )

    $reqParams = @{
        Method                  = 'POST'
        ContentType             = 'application/json'
        Body                    = ConvertTo-Json @{ refresh_token = "$($TokenObject.responseHeaders.'X-Refresh-Token')" }
        ResponseHeadersVariable = 'responseHeaders'
    }

    $reqParams['Uri'] = if ($PSBoundParameters.ContainsKey('UriOverride')) {

        $UriOverride
    }
    else { 'https://trade-service.wealthsimple.com/auth/refresh' }

    $response = Invoke-RestMethod @reqParams

    $Global:WSTokens = [PSCustomObject]@{

        responseHeaders = $responseHeaders
        responseBody    = $response
    }
}

function Get-WealthsimpleAccount {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [Object]$TokenObject,
        [Uri]$UriOverride
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
    }

    $reqParams['Uri'] = if ($PSBoundParameters.ContainsKey('UriOverride')) {

        $UriOverride
    }
    else { 'https://trade-service.wealthsimple.com/account/list' }

    Invoke-RestMethod @reqParams
}

#=========#---------------------------#
#endregion# Main (exported) Functions #
#=========#---------------------------#



#======#--------------------#
#region# Internal Functions #
#======#--------------------#

function ConvertFrom-SecureStringToPlainText ([SecureString]$SecureString) {

    [Marshal]::PtrToStringAuto(
        [Marshal]::SecureStringToBSTR($SecureString)
    )
}

#=========#--------------------#
#endregion# Internal Functions #
#=========#--------------------#
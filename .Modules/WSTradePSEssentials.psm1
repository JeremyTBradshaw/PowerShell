#Requires -Version 7.2.1
#Requires -PSEdition Core
using namespace System
using namespace System.Management.Automation.Host
using namespace System.Runtime.InteropServices
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates

<# Release Notes for v0.0.0 (2022-01-26):

    - Just starting to work on this now.  It's not ready for reliable use.  Looking at porting all of the examples here:
    https://github.com/MarkGalloway/wealthsimple-trade/blob/master/API.md
    - Enforcing option for OTP (i.e., MFA) to keep it somewhat safe.
#>

#======#---------------------------#
#region# Main (exported) Functions #
#======#---------------------------#

function Connect-WSTrade {
    <#
        .Synopsis
        Login to Wealthsimple Trade's unofficial API.

        .Example
        $UserAndPass = Get-Credential; Connect-WSTrade -Credential $UserAndPass

        .Example
        Connect-WSTrade
    #>
    [CmdletBinding()]
    param (
        [PSCredential]$Credential
    )

    $loginBody = if ($PSBoundParameters.ContainsKey('Credential')) {
        @{
            email    = $Credential.UserName
            password = ConvertFrom-SecureStringToPlainText $Credential.Password
        }
    }
    else {
        @{
            email    = Read-Host 'Enter username (email)'
            password = Read-Host -AsSecureString -Prompt 'Enter password'
        }
    }
    $loginBody['otp'] = Read-Host -Prompt 'Enter passcode'

    $reqParams = @{
        Method                  = 'POST'
        ContentType             = 'application/json'
        Uri                     = 'https://trade-service.wealthsimple.com/auth/login'
        Body                    = ConvertTo-Json -InputObject $loginBody
        ResponseHeadersVariable = 'responseHeaders'
    }
    $response = Invoke-RestMethod @reqParams

    $Global:WSTokens = [PSCustomObject]@{

        responseHeaders = $responseHeaders
        responseBody    = $response
    }
}

function Update-WSTradeTokens {
    <#
        .Synopsis
        Exchange refresh token for a new access token (i.e., re-login).

        .Example
        Update-WSTradeTokens
    #>
    [CmdletBinding()]
    param (
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method                  = 'POST'
        ContentType             = 'application/json'
        Uri                     = 'https://trade-service.wealthsimple.com/auth/refresh'
        Body                    = ConvertTo-Json @{ refresh_token = "$($TokenObject.responseHeaders.'X-Refresh-Token')" }
        ResponseHeadersVariable = 'responseHeaders'
    }
    $response = Invoke-RestMethod @reqParams

    $Global:WSTokens = [PSCustomObject]@{

        responseHeaders = $responseHeaders
        responseBody    = $response
    }
}

function Get-WSTradeAccount {
    <#
        .Synopsis
        Get all Wealthsimple Trade accounts (e.g., TFSA, RRSP, Crypto).

        .Example
        Get-WSTradeAccount
    #>
    [CmdletBinding()]
    param (
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/account/list'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeHistoricalAccountData {
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens,
        [ValidateSet('1d', '1w', '1m', '1y', 'all')]
        [String]$Timeframe = '1w',
        [string]$AccountId
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = "https://trade-service.wealthsimple.com/account/history/$($Timeframe)?account_id=$($AccountId)"
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeOrders {
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/orders'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeSecurity {
    [CmdletBinding(DefaultParameterSetName = 'TickerSymbol')]
    param(
        [Object]$TokenObject = $Global:WSTokens,

        [Parameter(ParameterSetName = 'SecurityId', Mandatory, ValueFromPipeline)]
        [string]$SecurityId,

        [Parameter(ParameterSetName = 'TickerSymbol', Mandatory)]
        [string]$TickerSymbol
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/securities'
    }
    if ($PSBoundParameters.ContainsKey('TickerSymbol')) {

        $reqParams['Uri'] = "$($reqParams['Uri'])?query=$($TickerSymbol)"
    }
    else {
        $reqParams['Uri'] = "$($reqParams['Uri'])/$($SecurityId)"
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradePositions {
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/account/positions'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeActivities {
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
    }

    $reqParams['Uri'] = 'https://trade-service.wealthsimple.com/account/activities'

    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeMe {
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/me'
    }
    Invoke-RestMethod @reqParams
}

function Get-WSTradePerson {
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/person'
    }
    Invoke-RestMethod @reqParams
}

function Get-WSTradeBankAccounts {
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/bank-accounts'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeDeposits {
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/deposits'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeForeignExchangeRate {
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.responseHeaders.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/forex'
    }
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

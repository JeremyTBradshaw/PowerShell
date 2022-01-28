#Requires -Version 5.1
using namespace System
using namespace System.Management.Automation.Host
using namespace System.Runtime.InteropServices
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates

<# Release Notes for v0.0.0 (2022-01-27):

    - Just starting to work on this now.  It's not ready for reliable use.  Looking at porting all of the examples here:
    https://github.com/MarkGalloway/wealthsimple-trade/blob/master/API.md
    - Enforcing option for OTP (i.e., MFA) to keep it somewhat safe.
#>

#======#---------------------------#
#region# Main (exported) Functions #
#======#---------------------------#

function Connect-WSTrade {
    <#
        .SYNOPSIS
        Login to Wealthsimple Trade's unofficial API.  This should be run before trying any of the other commands in
        the WSTradePSEssentials module.

        .PARAMETER Credential
        Optional parameter which accepts a PSCredential object.  If this parameter is not specified, the caller will be
        prompted for their email and password.  Whether specified or not, the caller will be prompted for their
        passcode, which is the OTP from their authenticator app.

        .EXAMPLE
        $UserAndPass = Get-Credential; Connect-WSTrade -Credential $UserAndPass

        .EXAMPLE
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
    }
    $response = Invoke-WebRequest @reqParams

    $Global:WSTokens = [PSCustomObject]@{

        'X-Access-Token'  = $response.Headers.'X-Access-Token'
        'X-Refresh-Token' = $response.Headers.'X-Refresh-Token'
    }
}

function Update-WSTradeTokens {
    <#
        .SYNOPSIS
        Exchange refresh token for a new access token (i.e., re-login).

        .EXAMPLE
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
        Body                    = ConvertTo-Json @{ refresh_token = "$($TokenObject.'X-Refresh-Token')" }
    }
    $response = Invoke-WebRequest @reqParams

    $Global:WSTokens = [PSCustomObject]@{

        'X-Access-Token'  = $response.Headers.'X-Access-Token'
        'X-Refresh-Token' = $response.Headers.'X-Refresh-Token'
    }
}

function Get-WSTradeAccount {
    <#
        .SYNOPSIS
        Get all Wealthsimple Trade accounts (e.g., TFSA, RRSP, Crypto).

        .EXAMPLE
        Get-WSTradeAccount
    #>
    [CmdletBinding()]
    param (
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/account/list'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeHistoricalAccountData {
    <#
        .SYNOPSIS
        Get historical account data from a WS Trade account, including point in time readings of net deposits, value,
        equity value, etc.

        .PARAMETER Timeframe
        The period of time (relative to now) for which to retrieve historical data.  Options: 1d, 1w, 1m, 1y, all.

        .PARAMETER AccountId
        Account ID value which can be found in the output from Get-WSTradeAccount.

        .EXAMPLE
        $Accounts = Get-WSTradeAccount; Get-WSTradeHistoricalAccountData -Timeframe all -AccountId $Accounts[1].id
    #>
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
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
        Uri         = "https://trade-service.wealthsimple.com/account/history/$($Timeframe)?account_id=$($AccountId)"
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeOrders {
    <#
        .SYNOPSIS
        Get all orders.

        .EXAMPLE
        Get-WSTradeOrders
    #>
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/orders'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeSecurity {
    <#
        .SYNOPSIS
        Get security (i.e., stock, ETF, crypto, etc.) by security ID or by search term (e.g., ticker symbol).

        .PARAMETER Search
        Specifies a string to search for, such as 'AAPL' if search for Apple, Inc.'s stock, or 'Microsoft' to lookup
        Microsoft's available securities.

        .PARAMETER SecurityId
        Specifies the exact security ID to get.

        .EXAMPLE
        Get-WSTradeSecurity -Search Microsoft

        .EXAMPLE
        Get-WSTradeSecurity -SecurityId sec-s-2b07d13e1dee4f418afe10d3ffeb5b9c
    #>
    [CmdletBinding(DefaultParameterSetName = 'Search')]
    param(
        [Object]$TokenObject = $Global:WSTokens,

        [Parameter(ParameterSetName = 'SecurityId', Mandatory, ValueFromPipeline)]
        [string]$SecurityId,

        [Parameter(ParameterSetName = 'Search', Mandatory)]
        [string]$Search
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/securities'
    }
    if ($PSBoundParameters.ContainsKey('Search')) {

        $reqParams['Uri'] = "$($reqParams['Uri'])?query=$($Search)"
    }
    else {
        $reqParams['Uri'] = "$($reqParams['Uri'])/$($SecurityId)"
    }
    $response = Invoke-RestMethod @reqParams
    if ($PSBoundParameters.ContainsKey('Search')) {

        $response | Select-Object -ExpandProperty Results
    }
    else { $response }
}

function Get-WSTradePositions {
    <#
        .SYNOPSIS
        Get all positions (holdings) including details like number of shares, current value, and more.

        .EXAMPLE
        Get-WSTradePositions
    #>
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/account/positions'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeActivities {
    <#
        .SYNOPSIS
        Get all activities from your WS Trade account.  Some of the activities are redundant to those which can be
        retrieved using the other Get-WSTrade*** functions.

        .EXAMPLE
        Get-WSTradeActivies
    #>
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
    }

    $reqParams['Uri'] = 'https://trade-service.wealthsimple.com/account/activities'

    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeMe {
    <#
        .SYNOPSIS
        Get your WS Trade user account's basic info.

        .EXAMPLE
        Get-WSTradeMe
    #>
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/me'
    }
    Invoke-RestMethod @reqParams
}

function Get-WSTradePerson {
    <#
        .SYNOPSIS
        Get your WS Trade account holder information.  CAREFUL, this one returns sensitive info like Social Insurance
        Number.

        .EXAMPLE
        Get-WSTradePerson
    #>
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/person'
    }
    Invoke-RestMethod @reqParams
}

function Get-WSTradeBankAccounts {
    <#
        .SYNOPSIS
        Get bank accounts which have been added to your WS Trade account.

        .EXAMPLE
        Get-WSTradeBankAccounts
    #>
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/bank-accounts'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeDeposits {
    <#
        .SYNOPSIS
        Get all deposits made into your WS Trade accounts.

        .EXAMPLE
        Get-WSTradeDeposits
    #>
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
        Uri         = 'https://trade-service.wealthsimple.com/deposits'
    }
    Invoke-RestMethod @reqParams | Select-Object -ExpandProperty Results
}

function Get-WSTradeForeignExchangeRate {
    <#
        .SYNOPSIS
        Get the foreign exchange rate for USD, relative to your currency (i.e., CAD).

        .EXAMPLE
        Get-WSTradeForeignExchangeRate
    #>
    [CmdletBinding()]
    param(
        [Object]$TokenObject = $Global:WSTokens
    )

    $reqParams = @{
        Method      = 'GET'
        ContentType = 'application/json'
        Headers     = @{ Authorization = "$($TokenObject.'X-Access-Token')" }
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

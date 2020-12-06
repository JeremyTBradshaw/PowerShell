using namespace System
using namespace System.Management.Automation.Host

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

    $B1 = @(
        "client_id=$($ApplicationId)",
        "scope=$($Scopes -join ' ')"
    ) -join '&'

    $DeviceCodeRequest = @{

        Method      = 'POST'
        Uri         = "https://login.microsoftonline.com/$($Endpoint)/oauth2/v2.0/devicecode"
        Body        = $B1
        ContentType = 'application/x-www-form-urlencoded'
        ErrorAction = 'Stop'
    }
    $DeviceCodeResponse = Invoke-RestMethod @DeviceCodeRequest

    $Stopwatch = [Diagnostics.Stopwatch]::StartNew()

    $DeviceCode = $DeviceCodeResponse.device_code
    $B2 = @(
        "grant_type=urn:ietf:params:oauth:grant-type:device_code",
        "client_id=$($ApplicationId)",
        "device_code=$($DeviceCode)"
    ) -join '&'

    switch (
        $host.UI.PromptForChoice(
    
            'Device Code Flow Started $($)',
            "$($DeviceCodeResponse.message)",
            [ChoiceDescription]('&Done'),
            0
        )
    ) {0 { <##> }}

    if ($Stopwatch.Elapsed.Minutes -lt 15) {

        $TokenParams = @{

            Method      = 'POST'
            Uri         = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
            Body        = $B2
            ContentType = 'application/x-www-form-urlencoded'
        }
        $TokenResponse = Invoke-RestMethod @TokenParams
    
        Write-Debug 'Inspect $DeviceCodeResponse, $TokenResponse'
    
        $TokenResponse
    }
    else { throw '15 minutes has passed, unable to request access token, please try again.'}
}

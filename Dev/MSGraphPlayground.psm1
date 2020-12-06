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
        $trResponse = $null
        do {
            if ($sw2.Elapsed.Seconds -ge $dcrResponse.interval) {
                try {
                    $trParams = @{
    
                        Method      = 'POST'
                        Uri         = "https://login.microsoftonline.com/$($Endpoint)/oauth2/v2.0/token"
                        Body        = $trBody
                        ContentType = 'application/x-www-form-urlencoded'
                        ErrorAction = 'Stop'
                    }
                    $trResponse = Invoke-RestMethod @trParams

                    $sw2.Restart()
                }
                catch {
                    if ($_.ErrorDetails.Message) {
                        
                        $trResponse = ConvertFrom-Json -InputObject $_.ErrorDetails.Message
    
                        if ($trResponse.error -eq 'authorization_pending') {
                        
                            Write-Warning "The user hasn't finished authenticating, but hasn't canceled the flow (error: authorization_pending)."
                            Write-Warning "Continuing to poll the token endpoint at the requested interval ($($dcrReponse.interval) seconds)"
                        }
                        elseif ($trResponse.error -match '^(authorization_declined)|(bad_verification_code)|(expired_token)$') {

                            Write-Debug 'Inspect $_, $_.ErrorDetails.'
                            throw "Authorization failed due to error: $($trResponse.error)."
                        }
                        else {
                            Write-Debug 'Inspect $_, $_.ErrorDetails.'
                            Write-Warning 'Authorization failed due to an unexpected error.'
                            throw $trResponse.error_description
                        }
                    }
                    else {
                        Write-Debug 'Inspect $_.'
                        Write-Warning 'Authorization failed due to an unexpected error.'
                        throw $_
                    }
                }
            }
            if (-not $trResponse) { Start-Sleep -Seconds 1 }
        }
        while ($sw1.Elapsed.Minutes -lt 15 -and -not $trResponse)

        # Output the token request response:
        $trResponse
    }
    else {
        throw "Authorization request expired at $($dcExpiration), please try again."
    }
}

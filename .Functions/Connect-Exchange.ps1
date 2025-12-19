function Connect-Exchange {
    [CmdletBinding(DefaultParameterSetName = 'ConnectionUri')]
    param (
        [Parameter(Mandatory)]
        [PSCredential]$Credential,

        [Parameter(Mandatory, ParameterSetName = 'FQDN')]
        [string]$ServerFQDN,

        [Parameter(ParameterSetName = 'FQDN')]
        [switch]$UseHttps,

        [Parameter(Mandatory, ParameterSetName = 'ConnectionUri')]
        [uri]$ConnectionUri,

        [ValidateSet('Basic', 'Default', 'Kerberos', 'Digest')]
        [string]$Authentication = 'Default',

        [string]$Prefix,
        [switch]$TrustAllCertificates
    )

    if ($PSCmdlet.ParameterSetName -eq 'FQDN') {

        $Script:ConnectionUri = "http://$($ServerFQDN)/PowerShell"

        if ($UseHttps) { $Script:ConnectionUri = $Script:ConnectionUri -replace 'http', 'https' }
    }
    else { $Script:ConnectionUri = $ConnectionUri }

    if ($Authentication -eq 'Basic' -and $Script:ConnectionUri -match '(http:)') { $Script:ConnectionUri = $Script:ConnectionUri -replace 'http:', 'https:' }

    $PSSessionParams = @{

        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = $Script:ConnectionUri
        Authentication    = $Authentication
        AllowRedirect     = $true
        Credential        = $Credential
    }

    if ($TrustAllCertificates) {

        $PSSessionParams['SessionOption'] = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    }

    try {
        $Session = New-PSSession @PSSessionParams -ErrorAction Stop

        $ImportSessionParams = @{

            Session             = $Session
            DisableNameChecking = $true
            ErrorAction         = 'Stop'
        }
        if ($Prefix) {
            $ImportSessionParams['Prefix'] = $Prefix
        }
        Import-PSSession @ImportSessionParams
    }
    catch { throw $_ }
}

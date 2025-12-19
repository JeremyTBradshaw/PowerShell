<#
    .Synopsis
    Quick and dirty get Exchange Certificates from all Exchange servers.
    Will fail if servers aren't reachable from the working location.

    .Outputs
    This script outputs to a CSV file on CurrentUser's Desktop, named:
    'ExchangeCertificates_yyyy-MM-dd_hh-mm-ss_tt.csv'

    e.g. ExchangeCertificates_2020-01-09_03-04-25_PM.csv
#>
#Requires -Version 4

if ($exscripts) {

    Get-ExchangeServer |
    ForEach-Object {
        $server = "$($_.Name)"
        Get-ExchangeCertificate -Server $_.Name |
        Select-Object -Property @{Name='Server';Expression={$server}},
        FriendlyName,
        Thumbprint,
        Services,
        NotBefore,
        NotAfter,
        Subject,
        @{Name='DnsNameList';Expression={$_.DnsNameList.Unicode -join ', '}}
    } | 
    Export-Csv "$($HOME)\Desktop\ExchangeCertificates_$([datetime]::Now.ToString('yyyy-MM-dd_hh-mm-ss_tt')).csv" -NTI
}
else {
    Write-Warning -Message "This script needs to be run from Exchange 2016 Management Shell."
}

<#
    .Synopsis
    Get-ExchangeServerVersion: Obtain Exchange server version, including:

      - Service pack and Rollup Update (Exchange 2010)
      - Cumulative Update (Exchange 2013/2016/2019)

    .Description
    Obtain Exchange server version info from one or more Exchange servers.  Supported Exchange versions are 2010, 2013,
    2016, and 2019.  For Exchange 2010, the guidance from Microsoft is to run the local command:
    
    Get-Command ExSetup | ForEach {$_.FileVersionInfo}

    In order to run this command remotely on Exchange 2010 servers, we can *try the following:

        Invoke-Command -ComputerName <server FQDN> -ScriptBlock {Get-Command ExSetup | ForEach {$_.FileVersionInfo}}

        *This requires the server to have WinRm configured (winrm quickconfig[enter])
    
    .Parameter Credential
    Can be a PSCredential object stored in a $variable (e.g. -Credential $MyCred) or (Get-Credential).  If not
    specified, the script will not prompt.  If an Exchange 2010 server is found (when using -AllServers), a warning
    will be shown, advising to re-run the script with -Credential, and the basic version info (Edition,
    AdminDisplayVersion) will be outputted.

    .Parameter ServerName
    One or more server names (e.g. -ServerName EX1, EX2.contoso.ca, EX4).  If no parameters are specified, this one
    will prompt for input.

    .Parameter AllServers
    Self-explanatory.  Will only get servers that are the same version or earlier than the one from the current
    PSSession.

    .Link
    https://github.com/JeremyTBradshaw/PowerShell/origin/main/Get-ExchangeServerVersion.ps1

    .Notes
    Eventually, I would like for this script to be able to dynamically parse Microsoft's docs page - 
    https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates - to obtain
    the latest build numbers, etc.  Until then, the information from that page is manually entered into
    this script.  See the script on GitHub to determine when the information was last updated.
#>

#Requires -Version 5.1
using Namespace System.Management.Automation.Runspaces
using Namespace System.Management.Automation

[CmdletBinding(
    DefaultParameterSetName = 'Specified'
)]
param(
    [PSCredential]$Credential,

    [Parameter(
        ParameterSetName = 'Specified',
        Mandatory
    )]
    [string[]]$ServerName,

    [Parameter(ParameterSetName = 'All')]
    [switch]$AllServers
)

# begin block has been moved to after the process block (for cosmetic reasons only).

process {

    $ExchangeServers = @()
    try {
        if ($PSCmdlet.ParameterSetName -eq 'Specified') {

            foreach ($Server in $ServerName) { $ExchangeServers += Get-ExchangeServer -Identity "$($Server)" -ErrorAction Stop }
        }
        else {
            $ExchangeServers += Get-ExchangeServer -ErrorAction Stop
        }
    }
    catch {
        Write-Warning -Message 'Failed on Get-ExchangeServer.  Error message:'
        Write-Warning -Message $_.Exception
        break
    }

    foreach ($Server in $ExchangeServers) {

        $ExVersionOutput = [PSCustomObject]@{
                        
            Server              = $Server.Name
            ServerRole          = $Server.ServerRole -join ', '
            Fqdn                = $Server.Fqdn
            Edition             = $Server.Edition
            AdminDisplayVersion = $Server.AdminDisplayVersion
        }

        $AdminDisplayVersion = $Server.AdminDisplayVersion.ToString() -replace '\(|\)' -split '\s'

        if ($AdminDisplayVersion[1] -like '14*') {

            if (-not ($PSBoundParameters.ContainsKey('Credential'))) {

                Write-Warning -Message 'Exchange 2010 server detected.  Re-run the script with -Credential parameter (required for Invoke-Command to work).'

                $ExVersionOutput |
                    Add-Member -NotePropertyName ProductName  -NotePropertyValue Unknown -PassThru |
                    Add-Member -NotePropertyName BuildNumber_Long  -NotePropertyValue Unknown -PassThru |
                    Add-Member -NotePropertyName BuildNumber_Short -NotePropertyValue Unknown 
            }
            else {
                try {
                    $exsetup = Invoke-Command -ComputerName $Server.Fqdn -Credential $Credential -ScriptBlock {

                        Get-Command exsetup | ForEach-Object { $_.FileVersionInfo }

                    } -ErrorAction Stop

                    $ProductName = ($ExchangeServerBuildNumbers.GetEnumerator() |
                            Select-Object -Index ($ExchangeServerLongBuildNumbers.IndexOf("$($exsetup.FileVersion)"))).Name

                    $ExVersionOutput |
                        Add-Member -NotePropertyName ProductName  -NotePropertyValue $ProductName -PassThru |
                        Add-Member -NotePropertyName BuildNumber_Long  -NotePropertyValue $ExchangeServerBuildNumbers["$($ProductName)"].BuildNumber_Long -PassThru |
                        Add-Member -NotePropertyName BuildNumber_Short -NotePropertyValue $ExchangeServerBuildNumbers["$($ProductName)"].BuildNumber_Short
                }
                catch {
                    Write-Warning -Message "Failed on Invoke-Command to server $($Server.Name)."

                    $ExVersionOutput |
                        Add-Member -NotePropertyName ProductName  -NotePropertyValue Unknown -PassThru |
                        Add-Member -NotePropertyName BuildNumber_Long  -NotePropertyValue Unknown -PassThru |
                        Add-Member -NotePropertyName BuildNumber_Short -NotePropertyValue Unknown 
                }
            }
        }
        elseif ($AdminDisplayVersion[1] -match '(15)|(16)') {

            $ProductName = ($ExchangeServerBuildNumbers.GetEnumerator() |
                    Select-Object -Index ($ExchangeServerShortBuildNumbers.IndexOf("$($AdminDisplayVersion[1,3] -join '.')"))).Name

            $ExVersionOutput |
                Add-Member -NotePropertyName ProductName  -NotePropertyValue $ProductName -PassThru |
                Add-Member -NotePropertyName BuildNumber_Long  -NotePropertyValue $ExchangeServerBuildNumbers["$($ProductName)"].BuildNumber_Long -PassThru |
                Add-Member -NotePropertyName BuildNumber_Short -NotePropertyValue $ExchangeServerBuildNumbers["$($ProductName)"].BuildNumber_Short
        }
        else {
            Write-Warning -Message "Server '$Server.Name' couldn't be matched to a known Exchange version (AdminDisplayVersion: $($Server.AdminDisplayVersion))."
            $ExVersionOutput |
                Add-Member -NotePropertyName ProductName  -NotePropertyValue Unknown -PassThru |
                Add-Member -NotePropertyName BuildNumber_Long  -NotePropertyValue Unknown -PassThru |
                Add-Member -NotePropertyName BuildNumber_Short -NotePropertyValue Unknown
        }

        if ($ExVersionOutput) { Write-Output -InputObject $ExVersionOutput }
    }
}
end { }

begin {
    Write-Verbose -Message "Checking for Exchange PSSession."

    [PSSession[]]$ExPSSession = Get-PSSession |
        Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' -and $_.ComputerName -ne 'outlook.office365.com' }

    if ($ExPSSession.Count -eq 1) {

        # Set scope to entire forest (important for multi-domain forests).
        Set-ADServerSettings -ViewEntireForest:$true -WarningAction SilentlyContinue
    }
    else {
        Write-Warning -Message "Requires a *single* remote session to a on-premises Exchange server (i.e. not EXO)."
        break
    }

    # Define Exchange server build numbers hashtable
    # Source: https://github.com/JeremyTBradshaw/PowerShell/blob/main/psd.ExchangeServerBuildNumbers.psd1
    $ExchangeServerBuildNumbers = @{

        <#
            I've purposely indexed the product name instead of the short or long build numbers.
            The reason for this is that when PS remoting, Exchange 2010 reports the long format while 2013 and newer report the short format.
        #>
    
        #Region 2019
    
        'Exchange Server 2019 CU5'                                   = @{
            'BuildNumber_Long'  = '15.02.0595.003'
            'BuildNumber_Short' = '15.2.595.3'
            'ReleaseDate'       = '3/17/2020'
        }
        'Exchange Server 2019 CU4'                                   = @{
            'BuildNumber_Long'  = '15.02.0529.005'
            'BuildNumber_Short' = '15.2.529.5'
            'ReleaseDate'       = '12/17/2019'
        }
        'Exchange Server 2019 CU3'                                   = @{
            'BuildNumber_Long'  = '15.02.0464.005'
            'BuildNumber_Short' = '15.2.464.5'
            'ReleaseDate'       = '9/17/2019'
        }
        'Exchange Server 2019 CU2'                                   = @{
            'BuildNumber_Long'  = '15.02.0397.003'
            'BuildNumber_Short' = '15.2.397.3'
            'ReleaseDate'       = '6/18/2019'
        }
        'Exchange Server 2019 CU1'                                   = @{
            'BuildNumber_Long'  = '15.02.0330.005'
            'BuildNumber_Short' = '15.2.330.5'
            'ReleaseDate'       = '2/12/2019'
        }
        'Exchange Server 2019 RTM'                                   = @{
            'BuildNumber_Long'  = '15.02.0221.012'
            'BuildNumber_Short' = '15.2.221.12'
            'ReleaseDate'       = '10/22/2018'
        }
        'Exchange Server 2019 Preview'                               = @{
            'BuildNumber_Long'  = '15.02.0196.000'
            'BuildNumber_Short' = '15.2.196.0'
            'ReleaseDate'       = '7/24/2018'
        }
    
        #EndRegion 2019
        #Region 2016
    
        'Exchange Server 2016 CU16'                                  = @{
            'BuildNumber_Long'  = '15.01.1979.003'
            'BuildNumber_Short' = '15.1.1979.3'
            'ReleaseDate'       = '3/17/2020'
        }
        'Exchange Server 2016 CU15'                                  = @{
            'BuildNumber_Long'  = '15.01.1913.005'
            'BuildNumber_Short' = '15.1.1913.5'
            'ReleaseDate'       = '12/17/2019'
        }
        'Exchange Server 2016 CU14'                                  = @{
            'BuildNumber_Long'  = '15.01.1847.003'
            'BuildNumber_Short' = '15.1.1847.3'
            'ReleaseDate'       = '9/17/2019'
        }
        'Exchange Server 2016 CU13'                                  = @{
            'BuildNumber_Long'  = '15.01.1779.002'
            'BuildNumber_Short' = '15.1.1779.2'
            'ReleaseDate'       = '6/18/2019'
        }
        'Exchange Server 2016 CU12'                                  = @{
            'BuildNumber_Long'  = '15.01.1713.005'
            'BuildNumber_Short' = '15.1.1713.5'
            'ReleaseDate'       = '2/12/2019'
        }
        'Exchange Server 2016 CU11'                                  = @{
            'BuildNumber_Long'  = '15.01.1591.010'
            'BuildNumber_Short' = '15.1.1591.10'
            'ReleaseDate'       = '10/16/2018'
        }
        'Exchange Server 2016 CU10'                                  = @{
            'BuildNumber_Long'  = '15.01.1531.003'
            'BuildNumber_Short' = '15.1.1531.3'
            'ReleaseDate'       = '6/19/2018'
        }
        'Exchange Server 2016 CU9'                                   = @{
            'BuildNumber_Long'  = '15.01.1466.003'
            'BuildNumber_Short' = '15.1.1466.3'
            'ReleaseDate'       = '3/20/2018'
        }
        'Exchange Server 2016 CU8'                                   = @{
            'BuildNumber_Long'  = '15.01.1415.002'
            'BuildNumber_Short' = '15.1.1415.2'
            'ReleaseDate'       = '12/19/2017'
        }
        'Exchange Server 2016 CU7'                                   = @{
            'BuildNumber_Long'  = '15.01.1261.035'
            'BuildNumber_Short' = '15.1.1261.35'
            'ReleaseDate'       = '9/19/2017'
        }
        'Exchange Server 2016 CU6'                                   = @{
            'BuildNumber_Long'  = '15.01.1034.026'
            'BuildNumber_Short' = '15.1.1034.26'
            'ReleaseDate'       = '6/27/2017'
        }
        'Exchange Server 2016 CU5'                                   = @{
            'BuildNumber_Long'  = '15.01.0845.034'
            'BuildNumber_Short' = '15.1.845.34'
            'ReleaseDate'       = '3/21/2017'
        }
        'Exchange Server 2016 CU4'                                   = @{
            'BuildNumber_Long'  = '15.01.0669.032'
            'BuildNumber_Short' = '15.1.669.32'
            'ReleaseDate'       = '12/13/2016'
        }
        'Exchange Server 2016 CU3'                                   = @{
            'BuildNumber_Long'  = '15.01.0544.027'
            'BuildNumber_Short' = '15.1.544.27'
            'ReleaseDate'       = '9/20/2016'
        }
        'Exchange Server 2016 CU2'                                   = @{
            'BuildNumber_Long'  = '15.01.0466.034'
            'BuildNumber_Short' = '15.1.466.34'
            'ReleaseDate'       = '6/21/2016'
        }
        'Exchange Server 2016 CU1'                                   = @{
            'BuildNumber_Long'  = '15.01.0396.030'
            'BuildNumber_Short' = '15.1.396.30'
            'ReleaseDate'       = '3/15/2016'
        }
        'Exchange Server 2016 RTM'                                   = @{
            'BuildNumber_Long'  = '15.01.0225.042'
            'BuildNumber_Short' = '15.1.225.42'
            'ReleaseDate'       = '10/1/2015'
        }
        'Exchange Server 2016 Preview'                               = @{
            'BuildNumber_Long'  = '15.01.0225.016'
            'BuildNumber_Short' = '15.1.225.16'
            'ReleaseDate'       = '7/22/2015'
        }
    
        #EndRegion 2016
        #Region 2013
    
        'Exchange Server 2013 CU23'                                  = @{
            'BuildNumber_Long'  = '15.00.1497.002'
            'BuildNumber_Short' = '15.0.1497.2'
            'ReleaseDate'       = '6/18/2019'
        }
        'Exchange Server 2013 CU22'                                  = @{
            'BuildNumber_Long'  = '15.00.1473.003'
            'BuildNumber_Short' = '15.0.1473.3'
            'ReleaseDate'       = '2/12/2019'
        }
        'Exchange Server 2013 CU21'                                  = @{
            'BuildNumber_Long'  = '15.00.1395.004'
            'BuildNumber_Short' = '15.0.1395.4'
            'ReleaseDate'       = '6/19/2018'
        }
        'Exchange Server 2013 CU20'                                  = @{
            'BuildNumber_Long'  = '15.00.1367.003'
            'BuildNumber_Short' = '15.0.1367.3'
            'ReleaseDate'       = '3/20/2018'
        }
        'Exchange Server 2013 CU19'                                  = @{
            'BuildNumber_Long'  = '15.00.1365.001'
            'BuildNumber_Short' = '15.0.1365.1'
            'ReleaseDate'       = '12/19/2017'
        }
        'Exchange Server 2013 CU18'                                  = @{
            'BuildNumber_Long'  = '15.00.1347.002'
            'BuildNumber_Short' = '15.0.1347.2'
            'ReleaseDate'       = '9/19/2017'
        }
        'Exchange Server 2013 CU17'                                  = @{
            'BuildNumber_Long'  = '15.00.1320.004'
            'BuildNumber_Short' = '15.0.1320.4'
            'ReleaseDate'       = '6/27/2017'
        }
        'Exchange Server 2013 CU16'                                  = @{
            'BuildNumber_Long'  = '15.00.1293.002'
            'BuildNumber_Short' = '15.0.1293.2'
            'ReleaseDate'       = '3/21/2017'
        }
        'Exchange Server 2013 CU15'                                  = @{
            'BuildNumber_Long'  = '15.00.1263.005'
            'BuildNumber_Short' = '15.0.1263.5'
            'ReleaseDate'       = '12/13/2016'
        }
        'Exchange Server 2013 CU14'                                  = @{
            'BuildNumber_Long'  = '15.00.1236.003'
            'BuildNumber_Short' = '15.0.1236.3'
            'ReleaseDate'       = '9/20/2016'
        }
        'Exchange Server 2013 CU13'                                  = @{
            'BuildNumber_Long'  = '15.00.1210.003'
            'BuildNumber_Short' = '15.0.1210.3'
            'ReleaseDate'       = '6/21/2016'
        }
        'Exchange Server 2013 CU12'                                  = @{
            'BuildNumber_Long'  = '15.00.1178.004'
            'BuildNumber_Short' = '15.0.1178.4'
            'ReleaseDate'       = '3/15/2016'
        }
        'Exchange Server 2013 CU11'                                  = @{
            'BuildNumber_Long'  = '15.00.1156.006'
            'BuildNumber_Short' = '15.0.1156.6'
            'ReleaseDate'       = '12/15/2015'
        }
        'Exchange Server 2013 CU10'                                  = @{
            'BuildNumber_Long'  = '15.00.1130.007'
            'BuildNumber_Short' = '15.0.1130.7'
            'ReleaseDate'       = '9/15/2015'
        }
        'Exchange Server 2013 CU9'                                   = @{
            'BuildNumber_Long'  = '15.00.1104.005'
            'BuildNumber_Short' = '15.0.1104.5'
            'ReleaseDate'       = '6/17/2015'
        }
        'Exchange Server 2013 CU8'                                   = @{
            'BuildNumber_Long'  = '15.00.1076.009'
            'BuildNumber_Short' = '15.0.1076.9'
            'ReleaseDate'       = '3/17/2015'
        }
        'Exchange Server 2013 CU7'                                   = @{
            'BuildNumber_Long'  = '15.00.1044.025'
            'BuildNumber_Short' = '15.0.1044.25'
            'ReleaseDate'       = '12/9/2014'
        }
        'Exchange Server 2013 CU6'                                   = @{
            'BuildNumber_Long'  = '15.00.0995.029'
            'BuildNumber_Short' = '15.0.995.29'
            'ReleaseDate'       = '8/26/2014'
        }
        'Exchange Server 2013 CU5'                                   = @{
            'BuildNumber_Long'  = '15.00.0913.022'
            'BuildNumber_Short' = '15.0.913.22'
            'ReleaseDate'       = '5/27/2014'
        }
        'Exchange Server 2013 SP1'                                   = @{
            'BuildNumber_Long'  = '15.00.0847.032'
            'BuildNumber_Short' = '15.0.847.32'
            'ReleaseDate'       = '2/25/2014'
        }
        'Exchange Server 2013 CU3'                                   = @{
            'BuildNumber_Long'  = '15.00.0775.038'
            'BuildNumber_Short' = '15.0.775.38'
            'ReleaseDate'       = '11/25/2013'
        }
        'Exchange Server 2013 CU2'                                   = @{
            'BuildNumber_Long'  = '15.00.0712.024'
            'BuildNumber_Short' = '15.0.712.24'
            'ReleaseDate'       = '7/9/2013'
        }
        'Exchange Server 2013 CU1'                                   = @{
            'BuildNumber_Long'  = '15.00.0620.029'
            'BuildNumber_Short' = '15.0.620.29'
            'ReleaseDate'       = '4/2/2013'
        }
        'Exchange Server 2013 RTM'                                   = @{
            'BuildNumber_Long'  = '15.00.0516.032'
            'BuildNumber_Short' = '15.0.516.32'
            'ReleaseDate'       = '12/3/2012'
        }
    
        #EndRegion 2013
        #Region 2010
    
        'Update Rollup 30 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0496.000'
            'BuildNumber_Short' = '14.3.496.0'
            'ReleaseDate'       = '2/11/2020'
        }
        'Update Rollup 29 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0468.000'
            'BuildNumber_Short' = '14.3.468.0'
            'ReleaseDate'       = '7/9/2019'
        }
        'Update Rollup 28 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0461.001'
            'BuildNumber_Short' = '14.3.461.1'
            'ReleaseDate'       = '6/7/2019'
        }
        'Update Rollup 27 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0452.000'
            'BuildNumber_Short' = '14.3.452.0'
            'ReleaseDate'       = '4/9/2019'
        }
        'Update Rollup 26 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0442.000'
            'BuildNumber_Short' = '14.3.442.0'
            'ReleaseDate'       = '2/12/2019'
        }
        'Update Rollup 25 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0435.000'
            'BuildNumber_Short' = '14.3.435.0'
            'ReleaseDate'       = '1/8/2019'
        }
        'Update Rollup 24 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0419.000'
            'BuildNumber_Short' = '14.3.419.0'
            'ReleaseDate'       = '9/5/2018'
        }
        'Update Rollup 23 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0417.001'
            'BuildNumber_Short' = '14.3.417.1'
            'ReleaseDate'       = '8/13/2018'
        }
        'Update Rollup 22 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0411.000'
            'BuildNumber_Short' = '14.3.411.0'
            'ReleaseDate'       = '6/19/2018'
        }
        'Update Rollup 21 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0399.002'
            'BuildNumber_Short' = '14.3.399.2'
            'ReleaseDate'       = '5/7/2018'
        }
        'Update Rollup 20 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0389.001'
            'BuildNumber_Short' = '14.3.389.1'
            'ReleaseDate'       = '3/5/2018'
        }
        'Update Rollup 19 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0382.000'
            'BuildNumber_Short' = '14.3.382.0'
            'ReleaseDate'       = '12/19/2017'
        }
        'Update Rollup 18 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0361.001'
            'BuildNumber_Short' = '14.3.361.1'
            'ReleaseDate'       = '7/11/2017'
        }
        'Update Rollup 17 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0352.000'
            'BuildNumber_Short' = '14.3.352.0'
            'ReleaseDate'       = '3/21/2017'
        }
        'Update Rollup 16 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0336.000'
            'BuildNumber_Short' = '14.3.336.0'
            'ReleaseDate'       = '12/13/2016'
        }
        'Update Rollup 15 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0319.002'
            'BuildNumber_Short' = '14.3.319.2'
            'ReleaseDate'       = '9/20/2016'
        }
        'Update Rollup 14 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0301.000'
            'BuildNumber_Short' = '14.3.301.0'
            'ReleaseDate'       = '6/21/2016'
        }
        'Update Rollup 13 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0294.000'
            'BuildNumber_Short' = '14.3.294.0'
            'ReleaseDate'       = '3/15/2016'
        }
        'Update Rollup 12 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0279.002'
            'BuildNumber_Short' = '14.3.279.2'
            'ReleaseDate'       = '12/15/2015'
        }
        'Update Rollup 11 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0266.002'
            'BuildNumber_Short' = '14.3.266.2'
            'ReleaseDate'       = '9/15/2015'
        }
        'Update Rollup 10 for Exchange Server 2010 SP3'              = @{
            'BuildNumber_Long'  = '14.03.0248.002'
            'BuildNumber_Short' = '14.3.248.2'
            'ReleaseDate'       = '6/17/2015'
        }
        'Update Rollup 9 for Exchange Server 2010 SP3'               = @{
            'BuildNumber_Long'  = '14.03.0235.001'
            'BuildNumber_Short' = '14.3.235.1'
            'ReleaseDate'       = '3/17/2015'
        }
        'Update Rollup 8 v2 for Exchange Server 2010 SP3'            = @{
            'BuildNumber_Long'  = '14.03.0224.002'
            'BuildNumber_Short' = '14.3.224.2'
            'ReleaseDate'       = '12/12/2014'
        }
        'Update Rollup 8 v1 for Exchange Server 2010 SP3 (recalled)' = @{
            'BuildNumber_Long'  = '14.03.0224.001'
            'BuildNumber_Short' = '14.3.224.1'
            'ReleaseDate'       = '12/9/2014'
        }
        'Update Rollup 7 for Exchange Server 2010 SP3'               = @{
            'BuildNumber_Long'  = '14.03.0210.002'
            'BuildNumber_Short' = '14.3.210.2'
            'ReleaseDate'       = '8/26/2014'
        }
        'Update Rollup 6 for Exchange Server 2010 SP3'               = @{
            'BuildNumber_Long'  = '14.03.0195.001'
            'BuildNumber_Short' = '14.3.195.1'
            'ReleaseDate'       = '5/27/2014'
        }
        'Update Rollup 5 for Exchange Server 2010 SP3'               = @{
            'BuildNumber_Long'  = '14.03.0181.006'
            'BuildNumber_Short' = '14.3.181.6'
            'ReleaseDate'       = '2/24/2014'
        }
        'Update Rollup 4 for Exchange Server 2010 SP3'               = @{
            'BuildNumber_Long'  = '14.03.0174.001'
            'BuildNumber_Short' = '14.3.174.1'
            'ReleaseDate'       = '12/9/2013'
        }
        'Update Rollup 8 for Exchange Server 2010 SP2'               = @{
            'BuildNumber_Long'  = '14.02.0390.003'
            'BuildNumber_Short' = '14.2.390.3'
            'ReleaseDate'       = '12/9/2013'
        }
        'Update Rollup 3 for Exchange Server 2010 SP3'               = @{
            'BuildNumber_Long'  = '14.03.0169.001'
            'BuildNumber_Short' = '14.3.169.1'
            'ReleaseDate'       = '11/25/2013'
        }
        'Update Rollup 2 for Exchange Server 2010 SP3'               = @{
            'BuildNumber_Long'  = '14.03.0158.001'
            'BuildNumber_Short' = '14.3.158.1'
            'ReleaseDate'       = '8/8/2013'
        }
        'Update Rollup 7 for Exchange Server 2010 SP2'               = @{
            'BuildNumber_Long'  = '14.02.0375.000'
            'BuildNumber_Short' = '14.2.375.0'
            'ReleaseDate'       = '8/3/2013'
        }
        'Update Rollup 1 for Exchange Server 2010 SP3'               = @{
            'BuildNumber_Long'  = '14.03.0146.000'
            'BuildNumber_Short' = '14.3.146.0'
            'ReleaseDate'       = '5/29/2013'
        }
        'Exchange Server 2010 SP3'                                   = @{
            'BuildNumber_Long'  = '14.03.0123.004'
            'BuildNumber_Short' = '14.3.123.4'
            'ReleaseDate'       = '2/12/2013'
        }
        'Update Rollup 6 Exchange Server 2010 SP2'                   = @{
            'BuildNumber_Long'  = '14.02.0342.003'
            'BuildNumber_Short' = '14.2.342.3'
            'ReleaseDate'       = '2/12/2013'
        }
        'Update Rollup 5 v2 for Exchange Server 2010 SP2'            = @{
            'BuildNumber_Long'  = '14.02.0328.010'
            'BuildNumber_Short' = '14.2.328.10'
            'ReleaseDate'       = '12/10/2012'
        }
        'Update Rollup 8 for Exchange Server 2010 SP1'               = @{
            'BuildNumber_Long'  = '14.01.0438.000'
            'BuildNumber_Short' = '14.1.438.0'
            'ReleaseDate'       = '12/10/2012'
        }
        'Update Rollup 5 for Exchange Server 2010 SP2'               = @{
            'BuildNumber_Long'  = '14.03.0328.005'
            'BuildNumber_Short' = '14.3.328.5'
            'ReleaseDate'       = '11/13/2012'
        }
        'Update Rollup 7 v3 for Exchange Server 2010 SP1'            = @{
            'BuildNumber_Long'  = '14.01.0421.003'
            'BuildNumber_Short' = '14.1.421.3'
            'ReleaseDate'       = '11/13/2012'
        }
        'Update Rollup 7 v2 for Exchange Server 2010 SP1'            = @{
            'BuildNumber_Long'  = '14.01.0421.002'
            'BuildNumber_Short' = '14.1.421.2'
            'ReleaseDate'       = '10/10/2012'
        }
        'Update Rollup 4 v2 for Exchange Server 2010 SP2'            = @{
            'BuildNumber_Long'  = '14.02.0318.004'
            'BuildNumber_Short' = '14.2.318.4'
            'ReleaseDate'       = '10/9/2012'
        }
        'Update Rollup 4 for Exchange Server 2010 SP2'               = @{
            'BuildNumber_Long'  = '14.02.0318.002'
            'BuildNumber_Short' = '14.2.318.2'
            'ReleaseDate'       = '8/13/2012'
        }
        'Update Rollup 7 for Exchange Server 2010 SP1'               = @{
            'BuildNumber_Long'  = '14.01.0421.000'
            'BuildNumber_Short' = '14.1.421.0'
            'ReleaseDate'       = '8/8/2012'
        }
        'Update Rollup 3 for Exchange Server 2010 SP2'               = @{
            'BuildNumber_Long'  = '14.02.0309.002'
            'BuildNumber_Short' = '14.2.309.2'
            'ReleaseDate'       = '5/29/2012'
        }
        'Update Rollup 2 for Exchange Server 2010 SP2'               = @{
            'BuildNumber_Long'  = '14.02.0298.004'
            'BuildNumber_Short' = '14.2.298.4'
            'ReleaseDate'       = '4/16/2012'
        }
        'Update Rollup 1 for Exchange Server 2010 SP2'               = @{
            'BuildNumber_Long'  = '14.02.0283.003'
            'BuildNumber_Short' = '14.2.283.3'
            'ReleaseDate'       = '2/13/2012'
        }
        'Exchange Server 2010 SP2'                                   = @{
            'BuildNumber_Long'  = '14.02.0247.005'
            'BuildNumber_Short' = '14.2.247.5'
            'ReleaseDate'       = '12/4/2011'
        }
        'Update Rollup 6 for Exchange Server 2010 SP1'               = @{
            'BuildNumber_Long'  = '14.01.0355.002'
            'BuildNumber_Short' = '14.1.355.2'
            'ReleaseDate'       = '10/27/2011'
        }
        'Update Rollup 5 for Exchange Server 2010 SP1'               = @{
            'BuildNumber_Long'  = '14.01.0339.001'
            'BuildNumber_Short' = '14.1.339.1'
            'ReleaseDate'       = '8/23/2011'
        }
        'Update Rollup 4 for Exchange Server 2010 SP1'               = @{
            'BuildNumber_Long'  = '14.01.0323.006'
            'BuildNumber_Short' = '14.1.323.6'
            'ReleaseDate'       = '7/27/2011'
        }
        'Update Rollup 3 for Exchange Server 2010 SP1'               = @{
            'BuildNumber_Long'  = '14.01.0289.007'
            'BuildNumber_Short' = '14.1.289.7'
            'ReleaseDate'       = '4/6/2011'
        }
        'Update Rollup 5 for Exchange Server 2010'                   = @{
            'BuildNumber_Long'  = '14.00.0726.000'
            'BuildNumber_Short' = '14.0.726.0'
            'ReleaseDate'       = '12/13/2010'
        }
        'Update Rollup 2 for Exchange Server 2010 SP1'               = @{
            'BuildNumber_Long'  = '14.01.0270.001'
            'BuildNumber_Short' = '14.1.270.1'
            'ReleaseDate'       = '12/9/2010'
        }
        'Update Rollup 1 for Exchange Server 2010 SP1'               = @{
            'BuildNumber_Long'  = '14.01.0255.002'
            'BuildNumber_Short' = '14.1.255.2'
            'ReleaseDate'       = '10/4/2010'
        }
        'Exchange Server 2010 SP1'                                   = @{
            'BuildNumber_Long'  = '14.01.0218.015'
            'BuildNumber_Short' = '14.1.218.15'
            'ReleaseDate'       = '8/23/2010'
        }
        'Update Rollup 4 for Exchange Server 2010'                   = @{
            'BuildNumber_Long'  = '14.00.0702.001'
            'BuildNumber_Short' = '14.0.702.1'
            'ReleaseDate'       = '6/10/2010'
        }
        'Update Rollup 3 for Exchange Server 2010'                   = @{
            'BuildNumber_Long'  = '14.00.0694.000'
            'BuildNumber_Short' = '14.0.694.0'
            'ReleaseDate'       = '4/13/2010'
        }
        'Update Rollup 2 for Exchange Server 2010'                   = @{
            'BuildNumber_Long'  = '14.00.0689.000'
            'BuildNumber_Short' = '14.0.689.0'
            'ReleaseDate'       = '3/4/2010'
        }
        'Update Rollup 1 for Exchange Server 2010'                   = @{
            'BuildNumber_Long'  = '14.00.0682.001'
            'BuildNumber_Short' = '14.0.682.1'
            'ReleaseDate'       = '12/9/2009'
        }
        'Exchange Server 2010 RTM'                                   = @{
            'BuildNumber_Long'  = '14.00.0639.021'
            'BuildNumber_Short' = '14.0.639.21'
            'ReleaseDate'       = '11/9/2009'
        }
    
        #EndRegion 2010
    }
    $ExchangeServerLongBuildNumbers = $ExchangeServerBuildNumbers.Values.BuildNumber_Long
    $ExchangeServerShortBuildNumbers = $ExchangeServerBuildNumbers.Values.BuildNumber_Short
}

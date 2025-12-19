<#
    .Synopsis
    Get Azure MFA status and details for users in Azure AD.

    .Parameter UserPrincipalName
    UPN of user to query for MFA details.  Accepts pipeline input.

    .Parameter MsolUser
    MsolUser objects from Get-MsolUser. Accepts objects in the pipeline or stored as variables.

    .Parameter All
    Specifies to get and process all MsolUser's.

    .Example
    .\Get-MsolUserMFADetails.ps1 -UserPrincipalName User1@jb365.ca
    PS C:\> .\Get-MsolUserMFADetails.ps1 User1@jb365.ca
    PS C:\> "User1@jb365.ca" | .\Get-MsolUserMFADetails.ps1

    .Example
    $HQUsers = Get-MsolUser -City 'Quispamsis'
    PS C:\> .\Get-MsolUserMFADetails.ps1 -MsolUser $HQUsers
    PS C:\> .\Get-MsolUserMFADetails.ps1 $HQUsers
    PS C:\> $HQUsers | .\Get-MsolUserMFADetails.ps1

    .Example
    .\Get-MsolUserMFADetails.ps1 -All | Export-csv MsolUserMFADetails.csv

    .Outputs
    [PSCustomObject] as follows:

    UserPrincipalName      : User1@jb365.ca
    DisplayName            : User1
    MfaState               : Disabled
    DefaultMethod          : PhoneAppNotification
    ConfiguredMethods      : OneWaySMS, TwoWayVoiceMobile, PhoneAppOTP, PhoneAppNotification
    AuthenticationPhone    : +1 8005551212
    AltAuthenticationPhone :
    PhoneAppAuthMethod     : Notification, OTP
    PhoneAppDeviceName     : ONEPLUS A5010
    UserType               : Member
    ObjectId               : 04eb85e2-e0bf-490b-81d2-e5559ad35d19
#>

#Requires -Version 5.1
#Requires -Module MSOnline

[CmdletBinding(DefaultParameterSetName = 'UserPrincipalName')]
param (
    [Parameter(
        ParameterSetName = 'UserPrincipalName',
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        Position = 0
    )]
    [ValidatePattern('.*\@.*\..*')]
    [string]$UserPrincipalName,
    
    [Parameter(
        ParameterSetName = 'MsolUser',
        ValueFromPipeline,
        Position = 0
    )]
    [Microsoft.Online.Administration.User[]]$MsolUser,

    [Parameter(
        ParameterSetName = 'All'
    )]
    [switch]$All
)

begin {

    try {
        Get-MsolAccountSku -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Warning -Message "Connect with Connect-MsolService before running this script."
        break
    }

    if ($Script:All) { $Script:MsolUser = Get-MsolUser -All }
}

process {

    if ($UserPrincipalName) {
        try {
            $Script:MsolUser = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction:Stop
        }
        catch {
            Write-Warning -Message "Failed to find MsolUser with UserPrincipalName ""$($UserPrincipalName)""."
            break
        }
    }

    foreach ($m in $MsolUser) {

        [PSCustomObject]@{

            UserPrincipalName      = $m.UserPrincipalName
            DisplayName            = $m.Displayname
            MfaState               = if ($m.StrongAuthenticationRequirements.State) { $m.StrongAuthenticationRequirements.State }
            else { "Disabled" }
            DefaultMethod          = ($m.StrongAuthenticationMethods | Where-Object { $_.IsDefault -eq $true }).MethodType
            ConfiguredMethods      = $m.StrongAuthenticationMethods.MethodType -join ", "
            AuthenticationPhone    = $m.StrongAuthenticationUserDetails.PhoneNumber
            AltAuthenticationPhone = $m.StrongAuthenticationUserDetails.AlternatePhoneNumber
            PhoneAppAuthMethod     = $m.StrongAuthenticationPhoneAppDetails.AuthenticationType -join ', '
            PhoneAppDeviceName     = $m.StrongAuthenticationPhoneAppDetails.DeviceName -join ', '
            UserType               = $m.UserType
            ObjectId               = $m.ObjectId
        }
    }
}

<#
    .Synopsis
    Generate a report all assigned licenses in Office 365 using AzureAD PS modules
    v1 and v2.

    .Description
    The main report which is generated consists item/object per unique user/product
    combination.  This enables easy Pivot Table usage in Excel, allowing to group
    license assignments by Product.

    .Parameter MsolUser
    MsolUser object, from Get-MsolUser

    .Parameter ObjectId
    ObjectId of an Azure AD user.

    .Parameter All
    Find all MsolUser's with licenses assigned (caution: slow/taxing).

    .Parameter OutputPath
    Folder path for output CSV file(s).

    .Parameter OutputToConsoleOnly
    Just as its name implies.  An example use case is when just testing out the
    script or quickly checking one or more users' license assignments.

    .Parameter GetLicensedGroupsMembers
    The script targets users, but collects all groups from which licenses are
    inherited by the target users. An example use case for this is to later compare
    group memberships with these groups in on-premises AD.
#>

#Requires -Version 5.1
#Requires -Modules MSOnline, AzureAD

[CmdletBinding(
    DefaultparameterSetName = 'ObjectId',
    SupportsShouldProcess = $true,
    ConfirmImpact = 'High'
)]
param(
    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'ObjectId',
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [guid[]]$ObjectId,

    [Parameter(
        ParameterSetName = 'MsolUser',
        ValueFromPipeline = $true
    )]
    [Microsoft.Online.Administration.User[]]$MsolUser,

    [Parameter(ParameterSetName = 'All')]
    [switch]$All,
    
    [ValidateScript ( {
            if (Test-Path -Path $_) { $true }
            else { throw "Couldn't validate `$OutputPath with Test-Path -Path '$($_)'" }
        })]
    [System.IO.DirectoryInfo]$OutputPath = $PSScriptRoot,
    [switch]$GetLicensedGroupsMembers,
    [switch]$OutputToConsoleOnly
)

process {

    #Region Main Script

    if ($PSBoundParameters.ContainsKey('MsolUser')) { $LicensedMsolUsers = $MsolUser }
    elseif ($PSBoundParameters.ContainsKey('ObjectId')) {

        $LicensedMsolUsers = foreach ($o in $ObjectId) {

            try { Get-MsolUser -ObjectId $o -ErrorAction:Stop }
            catch { Write-Warning -Message "Error from Get-MsolUser: $($_.Exception)" }
        }
    }
    elseif ($PSBoundParameters.ContainsKey('All')) {

        if ($PSCmdlet.ShouldProcess(
                "Get-MsolUser -All | Where-Object {`$_.Licenses}",
                'Execute long-running, taxing command')
        ) {
            # Toggle between these two options (via comment/#) to switch between using a pre-existing variable (e.g. for testing), or actually executing Get-MsolUser -All....
            
            # Option 1:
            # try { ([void]($LicensedMsolUsers += $global:AllLicensedMsolUsers[11100..11250])) }
            
            #Option 2 (currently selected):
            try { $LicensedMsolUsers = Get-MsolUser -All | Where-Object { $_.Licenses } }
            catch { Write-Warning -Message "Error from Get-MsolUser: $($_.Exception)" }
        }
    }

    if (@($LicensedMsolUsers).Count -eq 0) {

        Write-Warning 'No user license assignments to report.'
    }
    $LicensedMsolUsers | ForEach-Object {

        $thisUser = $null; $thisUser = $_
        $accountSource = if ($null -eq $_.ImmutableId) { 'Cloud-only' } else { 'Synced from on-premises AD' }
        $emailAddress = if ($thisUser.ProxyAddresses) { $thisUser.ProxyAddresses -cmatch 'SMTP:' -replace 'SMTP:' } else { '' }
        $recipientTypeDetails = if ($null -ne $thisUser.msExchRecipientTypeDetails) { $RecipientTypeDetailsFriendlyNames[$thisUser.msExchRecipientTypeDetails] } else { 'User' }
        $onPremisesObjectGuid = if ($accountSource -eq 'Cloud-only') { '' } else { [System.Guid]::New([System.Convert]::FromBase64String($thisUser.ImmutableId)) }

        foreach ($a in $_.LicenseAssignmentDetails) {

            $LicenseAssignmentDetail = [PSCustomObject]@{

                AccountSource        = $accountSource
                DisplayName          = $thisUser.DisplayName
                EmailAddress         = $emailAddress
                UserPrincipalName    = $thisUser.UserPrincipalName
                RecipientTypeDetails = $recipientTypeDetails
                Title                = $thisUser.Title
                Department           = $thisUser.Department
                Office               = $thisUser.Office
                UsageLocation        = $thisUser.UsageLocation
                
                Product              = $SkuPartNumberFriendlyNames[$a.AccountSku.SkuPartNumber]
                AssignmentPaths      = (

                    $a.Assignments.ReferencedObjectId.Guid | ForEach-Object {

                        if ($_ -eq $thisUser.ObjectId.Guid) { "Direct" }
                        elseif ($LicensedAzureADGroups.ContainsKey($_)) { "Inherited:$($LicensedAzureADGroups[$($_)].DisplayName) ($($_))" }
                        else {
                            try {
                                $AzureADGroup = @()
                                $AzureADGroup += Get-AzureADGroup -ObjectId $_ -ErrorAction:Stop
                            }
                            catch {
                                Write-Warning -Message "Error from Get-AzureADGroup: $($_.Exception)"
                            }
                            
                            if ($AzureADGroup.Count -eq 1) {
                                
                                $LicensedAzureADGroups["$($_)"] = $AzureADGroup
                                "Inherited:$($LicensedAzureADGroups[$($_)].DisplayName) ($($_))"
                            }
                            else { "Inherited:GROUP_NOT_FOUND_OR_AMBIGUOUS($($_))" }
                        }
                    }
                ) -join '; '

                ImmutableId          = $thisUser.ImmutableId
                OnPremisesObjectGuid = $onPremisesObjectGuid
                ObjectId             = $thisUser.ObjectId.Guid
            }

            switch ($OutputToConsoleOnly) {
                
                $true {
                    Write-Output $LicenseAssignmentDetail
                }
                $false { $MainReport += $LicenseAssignmentDetail }
            }
        }
    }
}

end {
    # MsolUser License Assignment Report:
    if ($MainReport.Count -ge 1) {

        try {
            [void]( New-Item -Path $MainReportCSV -ItemType File -ErrorAction:Stop )
            $MainReport | Export-Csv -Path $MainReportCSV -NoTypeInformation -Encoding:UTF8 -ErrorAction:Stop
        }
        catch { Write-Warning -Message "Problem creating / exporting to main report CSV file ($($MainReportCSV)).`r`nException: $($_.Exception)" }
    }
    else { Write-Warning -Message 'No user license assignments to report.' }


    # Licensed AzureADGroups' Members Report:
    if (($PSBoundParameters.ContainsKey('GetLicensedGroupsMembers')) -and ($MainReport.Count -ge 1)) {

        $GroupMemberReport += foreach ($g in $LicensedAzureADGroups.Keys) {
        
            foreach ($m in (Get-AzureADGroupMember -ObjectId $g -All:$true | where-object { $_.ObjectType -eq 'User' })) {

                [PSCustomObject]@{

                    GroupDisplayName           = $LicensedAzureADGroups[$g].DisplayName
                    GroupObjectId              = $LicensedAzureADGroups[$g].ObjectId
                    GroupOnPremiseSid          = $LicensedAzureADGroups[$g].OnPremisesSecurityIdentifier

                    MemberDisplayname          = $m.DisplayName
                    MemberObjectId             = $m.ObjectId
                    MemberOnPremisesObjectGuid = [System.Guid]::New([System.Convert]::FromBase64String($m.ImmutableId))
                }
            }
        }

        if ($GroupMemberReport.Count -ge 1) {

            if ($PSBoundParameters.ContainsKey('OutputToConsoleOnly')) {

                if ($PSCmdlet.ShouldProcess(
                        "Output licensed groups' members to console",
                        'Proceed')
                ) {
                    Write-Output $GroupMemberReport
                }
            }
            else {
                try {
                    [void]( New-Item -Path $GroupMemberReportCSV -ItemType File -ErrorAction:Stop )
                    $GroupMemberReport | Export-Csv -Path $GroupMemberReportCSV -NoTypeInformation -Encoding:UTF8 -ErrorAction:Stop
                }
                catch { Write-Warning -Message "Problem creating / exporting to main report CSV file ($($MainReportCSV)).`r`nException: $($_.Exception)" }
            }
        }
        else { Write-Warning -Message 'No group members to report.' }
    }
}

begin {

    #EndRegion Main Script
    
    
    #Region out-of-loop variables
    
    $FileNameDate = [datetime]::Now.ToString('yyyy-MM-dd_hhmmsstt')
    [System.IO.FileInfo]$MainReportCSV = Join-Path -Path $OutputPath -ChildPath "MsolUser-License-Assignment-Report_$($FileNameDate).csv"
    [System.IO.FileInfo]$GroupMemberReportCSV = Join-Path -Path $OutputPath -ChildPath "AzureAD-Licensed-Groups-Members_$($FileNameDate).csv"

    $MainReport = @()
    $GroupMemberReport = @()
    $LicensedAzureADGroups = @{ }

    #EndRegion out-of-loop variables


    #Region Reference Material

    # Load hashtable of SkuPartNumber's to product names and service plan friendly names (e.g. ENTERPRISEPACK -> Office 365 E3)
    # Source: https://github.com/JeremyTBradshaw/PowerShell/blob/master/psd.MSOnlineServicesSkuPartNumberFriendlyNames.psd1

    $SkuPartNumberFriendlyNames = @{

        # Provisioning ID                                     = Offer Display Name
        'AAD_BASIC_FACULTY'                                   = 'Azure Active Directory Basic for Faculty'
        'AAD_BASIC_STUDENT'                                   = 'Azure Active Directory Basic for Students'
        'AAD_PREMIUM'                                         = 'Azure Active Directory Premium P1'
        'AAD_PREMIUM_FACULTY'                                 = 'Azure Active Directory Premium P1 for Faculty'
        'AAD_PREMIUM_P2'                                      = 'Azure Active Directory Premium P2'
        'AAD_PREMIUM_P2_FACULTY'                              = 'Azure Active Directory Premium P2 for Faculty'
        'AAD_PREMIUM_P2_STUDENT'                              = 'Azure Active Directory Premium P2 for Students'
        'AAD_PREMIUM_STUDENT'                                 = 'Azure Active Directory Premium P1 for Students'
        'ADALLOM_STANDALONE'                                  = 'Microsoft Cloud App Security'
        'ATA'                                                 = 'Azure Advanced Threat Protection for Users'
        'ATP_ENTERPRISE'                                      = 'Office 365 Advanced Threat Protection (Plan 1)'
        'ATP_ENTERPRISE_FACULTY'                              = 'Office 365 Advanced Threat Protection (Plan 1) for faculty'
        'ATP_ENTERPRISE_STUDENT'                              = 'Office 365 Advanced Threat Protection (Plan 1) for students'
        'AX_DATABASE_STORAGE'                                 = 'Dynamics 365 Unified Operations - Additional Database Storage (Qualified Offer)'
        'AX_FILE_STORAGE'                                     = 'Dynamics 365 Unified Operations - Additional File Storage (Qualified Offer)'
        'AX_FILESTORAGE'                                      = 'Dynamics 365 Plan - Operations Additional File Storage (Qualified Offer)'
        'AX_SUPPORT_PRODIRECT'                                = 'Pro Direct Support for Dynamics 365 Unified Operations'
        'AXDATABASE'                                          = 'Dynamics 365 Plan - Unified Operations Additional Database Storage (Qualified Offer)'
        'BUSINESS_VOICE_MED'                                  = 'Microsoft 365 Business Voice'
        'CDS_API_CAPACITY'                                    = 'Power Apps and Power Automate capacity add-on'
        'CDS_DB_CAPACITY'                                     = 'Common Data Service Database Capacity'
        'CDS_FILE_CAPACITY'                                   = 'Common Data Service File Capacity'
        'CDS_LOG_CAPACITY'                                    = 'Common Data Service Log Capacity'
        'CDSAICAPACITY'                                       = 'AI Builder Capacity add-on'
        'CRM_AUTO_ROUTING_ADDON'                              = 'Dynamics 365 Field Service - Resource Scheduling Optimization'
        'CRM_ONLINE_PORTAL'                                   = 'Dynamics 365 Enterprise Edition - Additional Portal (Qualified Offer)'
        'CRM_ONLINE_PORTAL_ADDL_PAGE_VIEWS'                   = 'Dynamics 365 Enterprise Edition - Additional Portal Page Views (Qualified Offer)'
        'CRMINSTANCE'                                         = 'Dynamics 365 - Additional Production Instance (Qualified Offer)'
        'CRMSTORAGE'                                          = 'Dynamics 365 - Additional Database Storage (Qualified Offer)'
        'CRMTESTINSTANCE'                                     = 'Dynamics 365 - Additional Non-Production Instance (Qualified Offer)'
        'D365_CE_APPS_TRIAL'                                  = 'Dynamics 365 Customer Engagement Applications Trial'
        'D365_CSI_ADDON'                                      = 'Dynamics 365 Customer Service Insights Addnl Cases'
        'D365_CSI_STANDALONE'                                 = 'Dynamics 365 Customer Service Insights'
        'D365_CUSTOMER_SERVICE_ENT_ATTACH'                    = 'Dynamics 365 Customer Service Enterprise Attach to Qualifying Dynamics 365 Base Offer'
        'D365_CUSTOMER_SERVICE_PRO_ATTACH'                    = 'Dynamics 365 Customer Service Professional Attach to Qualifying Dynamics 365 Base Offer'
        'D365_E-COMMERCE_CLOUDSCALE_BASIC'                    = 'e-Commerce Cloud Scale Unit Basic'
        'D365_E-COMMERCE_CLOUDSCALE_PREMIUM'                  = 'e-Commerce Cloud Scale Unit Premium'
        'D365_E-COMMERCE_CLOUDSCALE_STANDARD'                 = 'e-Commerce Cloud Scale Unit Standard'
        'D365_FIELD_SERVICE_ATTACH'                           = 'Dynamics 365 Field Service Attach to Qualifying Dynamics 365 Base Offer'
        'D365_Operations_Enterprise_Storage_file_v2'          = 'Dynamics 365 Unified Operations - File Capacity'
        'D365_Operations_Enterprise_Storage_v2'               = 'Dynamics 365 Unified Operations - Database Capacity'
        'D365_SALES_ENT_ATTACH'                               = 'Dynamics 365 Sales Enterprise Attach to Qualifying Dynamics 365 Base Offer'
        'D365_SALES_PRO'                                      = 'Dynamics 365 Sales Professional'
        'D365_SALES_PRO_ATTACH'                               = 'Dynamics 365 Sales Professional Attach to Qualifying Dynamics 365 Base Offer'
        'D365_SALES_PRO_SMB'                                  = 'Dynamics 365 Sales Professional (SMB Offer)'
        'D365_VIRTUAL_AGENT_BASE'                             = 'Dynamics 365 Virtual Agent for Customer Service'
        'D365_VIRTUAL_AGENT_USL'                              = 'Dynamics 365 Virtual Agent for Customer Service User License'
        'DESKLESSPACK'                                        = 'Office 365 F3'
        'DYN365_ASSETMANAGEMENT'                              = 'Dynamics 365 Asset Management Addl Assets'
        'DYN365_BUSCENTRAL_DB_CAPACITY'                       = 'Dynamics 365 Business Central Database Capacity'
        'DYN365_BUSCENTRAL_DEVICE'                            = 'Dynamics 365 Business Central Device'
        'DYN365_BUSCENTRAL_ESSENTIAL'                         = 'Dynamics 365 Business Central Essential'
        'DYN365_BUSCENTRAL_PREMIUM'                           = 'Dynamics 365 Business Central Premium'
        'DYN365_BUSCENTRAL_TEAM_MEMBER'                       = 'Dynamics 365 Business Central Team Member'
        'DYN365_CS_CHAT'                                      = 'Dynamics 365 Customer Service Chat'
        'DYN365_CS_CHATBOT'                                   = 'Dynamics 365 Customer Service Chatbot session add-on'
        'DYN365_CS_MESSAGING'                                 = 'Dynamics 365 Customer Service Digital Messaging add-on'
        'DYN365_CUSTOMER_INSIGHTS_ADDON'                      = 'Dynamics 365 Customer Insights Addnl Profiles'
        'DYN365_CUSTOMER_INSIGHTS_ATTACH'                     = 'Dynamics 365 Customer Insights Attach'
        'DYN365_CUSTOMER_INSIGHTS_BASE'                       = 'Dynamics 365 Customer Insights'
        'DYN365_CUSTOMER_SERVICE_PRO'                         = 'Dynamics 365 Customer Service Professional'
        'DYN365_E-COMMERCE_RATINGS_REVIEWS'                   = 'e-Commerce Ratings and Reviews'
        'DYN365_E-COMMERCE_RECOMMENDATIONS'                   = 'e-Commerce Recommendations'
        'DYN365_E-COMMERCE_TIER1'                             = 'e-Commerce Tier 1'
        'DYN365_E-COMMERCE_TIER1_OVERAGE'                     = 'e-Commerce Tier 1 Overage'
        'DYN365_E-COMMERCE_TIER2'                             = 'e-Commerce Tier 2'
        'DYN365_E-COMMERCE_TIER2_OVERAGE'                     = 'e-Commerce Tier 2 Overage'
        'DYN365_E-COMMERCE_TIER3'                             = 'e-Commerce Tier 3'
        'DYN365_E-COMMERCE_TIER3_OVERAGE'                     = 'e-Commerce Tier 3 Overage'
        'DYN365_ENTERPRISE_CUSTOMER_SERVICE'                  = 'Dynamics 365 Customer Service Enterprise'
        'DYN365_ENTERPRISE_FIELD_SERVICE'                     = 'Dynamics 365 Field Service'
        'DYN365_ENTERPRISE_PROJECT_SERVICE_AUTOMATION'        = 'Dynamics 365 Project Service Automation'
        'DYN365_ENTERPRISE_SALES'                             = 'Dynamics 365 Sales Enterprise Edition'
        'DYN365_FINANCE'                                      = 'Dynamics 365 Finance'
        'DYN365_FINANCE_ATTACH'                               = 'Dynamics 365 Finance Attach to Qualifying Dynamics 365 Base Offer'
        'DYN365_FINANCIALS_ACCOUNTANT_SKU'                    = 'Dynamics 365 Business Central External Accountant'
        'DYN365_FRAUD_PROTECTION_ASSESSMENTS'                 = 'Dynamics 365 Fraud Protection Addl Assessments'
        'DYN365_FRAUD_PROTECTION_BASE'                        = 'Dynamics 365 Fraud Protection'
        'DYN365_HUMAN_RESOURCES'                              = 'Dynamics 365 Human Resources'
        'DYN365_HUMAN_RESOURCES_ATTACH'                       = 'Dynamics 365 Human Resources Attach to Qualifying Dynamics 365 Base Offer'
        'DYN365_HUMAN_RESOURCES_SANDBOX'                      = 'Dynamics 365 Human Resources Sandbox'
        'DYN365_HUMAN_RESOURCES_SELF_SERVE'                   = 'Dynamics 365 Human Resources Self Service'
        'DYN365_IOT_INTELLIGENCE_ADDL_MACHINES'               = 'IoT Intelligence Additional Machines'
        'DYN365_IOT_INTELLIGENCE_SCENARIO'                    = 'IoT Intelligence Scenario'
        'DYN365_MARKETING_APP'                                = 'Dynamics 365 Marketing'
        'DYN365_MARKETING_APP_ATTACH'                         = 'Dynamics 365 Marketing Attach'
        'DYN365_MARKETING_APPLICATION_ADDON'                  = 'Dynamics 365 Marketing Additional Application'
        'DYN365_MARKETING_CONTACT_ADDON'                      = 'Dynamics 365 Marketing Addnl Contacts Tier 1'
        'DYN365_MARKETING_CONTACT_ADDON_T2'                   = 'Dynamics 365 Marketing Addnl Contacts Tier 2'
        'DYN365_MARKETING_CONTACT_ADDON_T3'                   = 'Dynamics 365 Marketing Addnl Contacts Tier 3'
        'DYN365_MARKETING_CONTACT_ADDON_T4'                   = 'Dynamics 365 Marketing Addnl Contacts Tier 4'
        'DYN365_MARKETING_CONTACT_ADDON_T5'                   = 'Dynamics 365 Marketing Addnl Contacts Tier 5'
        'DYN365_MARKETING_CONTACT_CE_PLAN_ADDON'              = 'Dynamics 365 Marketing Addnl Contacts Tier 1 for CE Plan'
        'DYN365_MARKETING_SANDBOX_APPLICATION_ADDON'          = 'Dynamics 365 Marketing Additional Non-Prod Application'
        'Dyn365_Operations_Activity'                          = 'Dynamics 365 Unified Operations ? Activity'
        'DYN365_OPS_ORDERLINES'                               = 'Dynamics 365 Unified Operations ? Order Lines'
        'DYN365_RETAIL'                                       = 'Dynamics 365 Commerce'
        'DYN365_RETAIL_ATTACH'                                = 'Dynamics 365 Commerce Attach to Qualifying Dynamics 365 Base Offer'
        'DYN365_SALES_INSIGHTS'                               = 'Dynamics 365 Sales Insights'
        'DYN365_Sales_Insights_CI_AddOn'                      = 'Dynamics 365 Call Intelligence AddOn'
        'DYN365_SCM'                                          = 'Dynamics 365 Supply Chain Management'
        'DYN365_SCM_ATTACH'                                   = 'Dynamics 365 Supply Chain Management Attach to Qualifying Dynamics 365 Base Offer'
        'DYN365_TEAM_MEMBERS'                                 = 'Dynamics 365 Team Members'
        'Dynamics_365_for_Operations_Devices'                 = 'Dynamics 365 Unified Operations ? Device'
        'Dynamics_365_for_Operations_Sandbox_Tier1_SKU'       = 'Dynamics 365 Unified Operations - Sandbox Tier 1:Developer & Test Instance'
        'Dynamics_365_for_Operations_Sandbox_Tier1_SKU_Plan2' = 'Dynamics 365 Plan - Unified Operations Sandbox Tier 1:Developer & Test Instance'
        'Dynamics_365_for_Operations_Sandbox_Tier2_SKU'       = 'Dynamics 365 Unified Operations - Sandbox Tier 2:Standard Acceptance Testing'
        'Dynamics_365_for_Operations_Sandbox_Tier2_SKU_Plan2' = 'Dynamics 365 Plan - Unified Operations Sandbox Tier 2:Standard Acceptance Testing'
        'Dynamics_365_for_Operations_Sandbox_Tier3_SKU'       = 'Dynamics 365 Unified Operations - Sandbox Tier 3:Premier Acceptance Testing'
        'Dynamics_365_for_Operations_Sandbox_Tier3_SKU_Plan2' = 'Dynamics 365 Plan - Unified Operations Sandbox Tier 3:Premier Acceptance Testing'
        'Dynamics_365_for_Operations_Sandbox_Tier4_SKU'       = 'Dynamics 365 Unified Operations - Sandbox Tier 4:Standard Performance Testing'
        'Dynamics_365_for_Operations_Sandbox_Tier4_SKU_Plan2' = 'Dynamics 365 Plan - Unified Operations Sandbox Tier 4:Standard Performance Testing'
        'Dynamics_365_for_Operations_Sandbox_Tier5_SKU'       = 'Dynamics 365 Unified Operations - Sandbox Tier 5:Premier Performance Testing'
        'Dynamics_365_for_Operations_Sandbox_Tier5_SKU_Plan2' = 'Dynamics 365 Plan - Unified Operations Sandbox Tier 5:Premier Performance Testing'
        'E3_VDA_only'                                         = 'Windows 10 Enterprise E3 VDA'
        'EDISCOVERY_STORAGE_500GB'                            = 'Advanced eDiscovery Storage'
        'EDISCOVERY_STORAGE_500GB_FACULTY'                    = 'Advanced eDiscovery Storage for faculty'
        'EMS'                                                 = 'Enterprise Mobility + Security E3'
        'EMS_EDU_FACULTY'                                     = 'Enterprise Mobility + Security A3 for Faculty'
        'EMS_EDU_STUDENT'                                     = 'Enterprise Mobility + Security A3 for Students'
        'EMS_EDU_STUUSEBNFT'                                  = 'Enterprise Mobility + Security A3 for Students use benefit'
        'EMSPREMIUM'                                          = 'Enterprise Mobility + Security E5'
        'EMSPREMIUM_EDU_FACULTY'                              = 'Enterprise Mobility + Security A5 for Faculty'
        'EMSPREMIUM_EDU_STUDENT'                              = 'Enterprise Mobility + Security A5 for Students'
        'EMSPREMIUM_EDU_STUUSEBNFT'                           = 'Enterprise Mobility + Security A5 for Students use benefit'
        'ENTERPRISEPACK'                                      = 'Office 365 E3'
        'ENTERPRISEPACKPLUS_FACULTY'                          = 'Office 365 A3 for faculty'
        'ENTERPRISEPACKPLUS_STUDENT'                          = 'Office 365 A3 for students'
        'ENTERPRISEPACKPLUS_STUUSEBNFT'                       = 'Office 365 A3 for students use benefit'
        'ENTERPRISEPREMIUM'                                   = 'Office 365 E5'
        'ENTERPRISEPREMIUM_FACULTY'                           = 'Office 365 A5 for faculty'
        'ENTERPRISEPREMIUM_NOPSTNCONF'                        = 'Office 365 E5 without Audio Conferencing (Nonprofit Staff Pricing)'
        'ENTERPRISEPREMIUM_NOPSTNCONF_STUUSEBNFT'             = 'Office 365 A5 without Audio Conferencing for students use benefit'
        'ENTERPRISEPREMIUM_STUDENT'                           = 'Office 365 A5 for students'
        'ENTERPRISEPREMIUM_STUUSEBNFT'                        = 'Office 365 A5 for students use benefit'
        'EOP_ENTERPRISE'                                      = 'Exchange Online Protection'
        'EXCHANGE_ANALYTICS_FACULTY'                          = 'Microsoft MyAnalytics for faculty'
        'EXCHANGE_ANALYTICS_STUDENT'                          = 'Microsoft MyAnalytics for students'
        'EXCHANGEARCHIVE'                                     = 'Exchange Online Archiving for Exchange Server'
        'EXCHANGEARCHIVE_ADDON'                               = 'Exchange Online Archiving for Exchange Online'
        'EXCHANGEDESKLESS'                                    = 'Exchange Online Kiosk'
        'EXCHANGEENTERPRISE'                                  = 'Exchange Online (Plan 2)'
        'EXCHANGESTANDARD'                                    = 'Exchange Online (Plan 1)'
        'FLOW_BUSINESS_PROCESS'                               = 'Power Automate per business process plan'
        'FLOW_PER_USER'                                       = 'Power Automate per user plan'
        'FLOW_RUNS'                                           = 'Power Automate Additional Runs per 50,000 (Qualified Offer)'
        'FLOW_RUNS_D365'                                      = 'Flow Additional Runs per 50,000 for Dyn365 (Qualified Offer)'
        'Forms_Pro_AddOn'                                     = 'Forms Pro Addl Responses'
        'GUIDES_USER'                                         = 'Dynamics 365 Guides'
        'GUIDES_USER_FACULTY'                                 = 'Dynamics 365 Guides for Faculty'
        'GUIDES_USER_STUDENT'                                 = 'Dynamics 365 Guides for Students'
        'IDENTITY_THREAT_PROTECTION'                          = 'Microsoft 365 E5 Security'
        'IDENTITY_THREAT_PROTECTION_FACULTY'                  = 'Microsoft 365 A5 Security for faculty'
        'IDENTITY_THREAT_PROTECTION_STUDENT'                  = 'Microsoft 365 A5 Security for students'
        'IDENTITY_THREAT_PROTECTION_STUUSEBNFT'               = 'Microsoft 365 A5 Security for student use benefits'
        'INFORMATION_PROTECTION_COMPLIANCE'                   = 'Microsoft 365 E5 Compliance'
        'INFORMATION_PROTECTION_COMPLIANCE_FACULTY'           = 'Microsoft 365 A5 Compliance for faculty'
        'INFORMATION_PROTECTION_COMPLIANCE_STUDENT'           = 'Microsoft 365 A5 Compliance for students'
        'INTUNE_A'                                            = 'Intune'
        'INTUNE_A_D'                                          = 'Microsoft Intune Device'
        'INTUNE_EDU'                                          = 'Microsoft Intune for Education for Faculty'
        'INTUNE_STORAGE'                                      = 'Intune Extra Storage'
        'KAIZALA_FACULTY'                                     = 'Microsoft Kaizala Pro for faculty'
        'KAIZALA_STANDARD'                                    = 'Microsoft Kaizala Pro'
        'KAIZALA_STUDENT'                                     = 'Microsoft Kaizala Pro for students'
        'M365_DISC0VER_RESPOND'                               = 'Microsoft 365 E5 eDiscovery and Audit'
        'M365_DISC0VER_RESPOND_FACULTY'                       = 'Microsoft 365 A5 eDiscovery and Audit for faculty'
        'M365_DISC0VER_RESPOND_STUDENT'                       = 'Microsoft 365 A5 eDiscovery and Audit for students'
        'M365_F1'                                             = 'Microsoft 365 F1'
        'M365_INFO_PROTECTION_GOVERNANCE'                     = 'Microsoft 365 E5 Information Protection and Governance'
        'M365_INFO_PROTECTION_GOVERNANCE_FACULTY'             = 'Microsoft 365 A5 Information Protection and Governance for faculty'
        'M365_INFO_PROTECTION_GOVERNANCE_STUDENT'             = 'Microsoft 365 A5 Information Protection and Governance for students'
        'M365_INSIDER_RISK_MANAGEMENT'                        = 'Microsoft 365 E5 Insider Risk Management'
        'M365_INSIDER_RISK_MANAGEMENT_FACULTY'                = 'Microsoft 365 A5 Insider Risk Management for faculty'
        'M365_INSIDER_RISK_MANAGEMENT_STUDENT'                = 'Microsoft 365 A5 Insider Risk Management for students'
        'M365EDU_A1'                                          = 'Microsoft 365 A1'
        'M365EDU_A3_FACULTY'                                  = 'Microsoft 365 A3 for faculty'
        'M365EDU_A3_STUDENT'                                  = 'Microsoft 365 A3 for students'
        'M365EDU_A3_STUUSEBNFT'                               = 'Microsoft 365 A3 for students use benefit'
        'M365EDU_A5_FACULTY'                                  = 'Microsoft 365 A5 for faculty'
        'M365EDU_A5_STUDENT'                                  = 'Microsoft 365 A5 for students'
        'M365EDU_A5_STUUSEBNFT'                               = 'Microsoft 365 A5 for students use benefit'
        'MCOCAP'                                              = 'Common Area Phone'
        'MCOCAP_FACULTY'                                      = 'Common Area Phone for faculty'
        'MCOCAP_STUDENT'                                      = 'Common Area Phone for students'
        'MCOEV'                                               = 'Microsoft 365 Phone System'
        'MCOEV_FACULTY'                                       = 'Microsoft 365 Phone System for faculty'
        'MCOEV_STUDENT'                                       = 'Microsoft 365 Phone System for students'
        'MCOMEETADV'                                          = 'Microsoft 365 Audio Conferencing'
        'MCOMEETADV_FACULTY'                                  = 'Microsoft 365 Audio Conferencing for faculty'
        'MCOMEETADV_STUDENT'                                  = 'Microsoft 365 Audio Conferencing for students'
        'MCOPLUSCAL'                                          = 'Skype for Business Plus CAL'
        'MCOPLUSCAL_FACULTY'                                  = 'Skype for Business Plus CAL for faculty'
        'MCOPLUSCAL_STUDENT'                                  = 'Skype for Business Plus CAL for students'
        'MCOPSTN_5'                                           = 'Microsoft 365 Domestic Calling Plan (120 min)'
        'MCOPSTN_5_FACULTY'                                   = 'Microsoft 365 Domestic Calling Plan (120 min) for faculty'
        'MCOPSTN_5_STUDENT'                                   = 'Microsoft 365 Domestic Calling Plan (120 min) for students'
        'MCOPSTN1'                                            = 'Microsoft 365 Domestic Calling Plan'
        'MCOPSTN1_FACULTY'                                    = 'Microsoft 365 Domestic Calling Plan for faculty'
        'MCOPSTN1_STUDENT'                                    = 'Microsoft 365 Domestic Calling Plan for students'
        'MCOPSTN2'                                            = 'Microsoft 365 Domestic and International Calling Plan'
        'MCOPSTN2_FACULTY'                                    = 'Microsoft 365 Domestic and International Calling Plan for faculty'
        'MCOPSTN2_STUDENT'                                    = 'Microsoft 365 Domestic and International Calling Plan for students'
        'MCOPSTN9'                                            = 'Microsoft 365 International Calling Plan for SMB'
        'MDATP_Server'                                        = 'MDATP for Servers'
        'MDATP_XPLAT'                                         = 'Microsoft Defender Advanced Threat Protection'
        'MDATP_XPLAT_EDU'                                     = 'Microsoft Defender Advanced Threat Protection for Education'
        'MEE_FACULTY'                                         = 'Minecraft: Education Edition (per user)'
        'MEETING_ROOM'                                        = 'Meeting Room'
        'MEETING_ROOM_FACULTY'                                = 'Meeting Room for faculty'
        'MEETING_ROOM_STUDENT'                                = 'Meeting Room for students'
        'MICROSOFT_LAYOUT'                                    = 'Dynamics 365 Layout'
        'MICROSOFT_LAYOUT_FACULTY'                            = 'Dynamics 365 Layout for Faculty'
        'MICROSOFT_LAYOUT_STUDENT'                            = 'Dynamics 365 Layout for Students'
        'MICROSOFT_REMOTE_ASSIST'                             = 'Dynamics 365 Remote Assist'
        'MICROSOFT_REMOTE_ASSIST_ATTACH'                      = 'Dynamics 365 Remote Assist Attach'
        'MICROSOFT_REMOTE_ASSIST_ATTACH_FACULTY'              = 'Dynamics 365 Remote Assist Attach for Faculty'
        'MICROSOFT_REMOTE_ASSIST_ATTACH_STUDENT'              = 'Dynamics 365 Remote Assist Attach for Students'
        'MICROSOFT_REMOTE_ASSIST_FACULTY'                     = 'Dynamics 365 Remote Assist for Faculty'
        'MICROSOFT_REMOTE_ASSIST_STUDENT'                     = 'Dynamics 365 Remote Assist for Students'
        'NBPOSTS'                                             = 'Microsoft Social Engagement Additional 10K Posts'
        'O365_BUSINESS_ESSENTIALS'                            = 'Office 365 Business Essentials'
        'O365_BUSINESS_PREMIUM'                               = 'Office 365 Business Premium'
        'O365_DLP'                                            = 'Office 365 Data Loss Prevention'
        'OFFICESUBSCRIPTION'                                  = 'Office 365 ProPlus'
        'OFFICESUBSCRIPTION_FACULTY'                          = 'Office 365 ProPlus for faculty'
        'OFFICESUBSCRIPTION_STUDENT'                          = 'Office 365 ProPlus for students'
        'PBI_PREMIUM_EM3_ADDON'                               = 'Power BI Premium EM3 (Nonprofit Staff Pricing)'
        'PBI_PREMIUM_P1_ADDON'                                = 'Power BI Premium P1'
        'PBI_PREMIUM_P2_ADDON'                                = 'Power BI Premium P2'
        'PBI_PREMIUM_P3_ADDON'                                = 'Power BI Premium P3'
        'PBI_PREMIUM_P4_ADDON'                                = 'Power BI Premium P4'
        'PBI_PREMIUM_P5_ADDON'                                = 'Power BI Premium P5'
        'PHONESYSTEM_VIRTUALUSER'                             = 'Microsoft 365 Phone System - Virtual User'
        'PHONESYSTEM_VIRTUALUSER_FACULTY'                     = 'Microsoft 365 Phone System - Virtual User for faculty'
        'PHONESYSTEM_VIRTUALUSER_STUDENT'                     = 'Microsoft 365 Phone System - Virtual User for students'
        'POWER_BI_PRO'                                        = 'Power BI Pro'
        'POWER_BI_PRO_CE'                                     = 'Power BI Pro (Nonprofit Staff Pricing)'
        'POWER_BI_PRO_FACULTY'                                = 'Power BI Pro for faculty'
        'POWER_BI_PRO_STUDENT'                                = 'Power BI Pro for students'
        'POWERAPPS_PER_APP'                                   = 'Power Apps per app plan'
        'POWERAPPS_PER_USER'                                  = 'Power Apps per user plan'
        'POWERAPPS_PORTALS_LOGIN'                             = 'Power Apps Portals login capacity add-on'
        'POWERAPPS_PORTALS_LOGIN_T2'                          = 'Power Apps Portals login capacity add-on Tier 2 (10 unit min)'
        'POWERAPPS_PORTALS_LOGIN_T3'                          = 'Power Apps Portals login capacity add-on Tier 3 (50 unit min)'
        'POWERAPPS_PORTALS_PAGEVIEW'                          = 'Power Apps Portals page view capacity add-on'
        'POWERAUTOMATE_ATTENDED_RPA'                          = 'Power Automate per user with attended RPA plan'
        'POWERAUTOMATE_UNATTENDED_RPA'                        = 'Power Automate unattended RPA add-on'
        'PROJECT_P1'                                          = 'Project Plan 1'
        'PROJECTESSENTIALS'                                   = 'Project Online Essentials'
        'PROJECTESSENTIALS_FACULTY'                           = 'Project Online Essentials for faculty'
        'PROJECTESSENTIALS_STUDENT'                           = 'Project Online Essentials for students'
        'PROJECTPREMIUM'                                      = 'Project Plan 5'
        'PROJECTPREMIUM_FACULTY'                              = 'Project Plan 5 for faculty'
        'PROJECTPREMIUM_STUDENT'                              = 'Project Plan 5 for students'
        'PROJECTPROFESSIONAL'                                 = 'Project Plan 3'
        'PROJECTPROFESSIONAL_FACULTY'                         = 'Project Plan 3 for faculty'
        'PROJECTPROFESSIONAL_STUDENT'                         = 'Project Plan 3 for students'
        'RIGHTSMANAGEMENT'                                    = 'Azure Information Protection Premium P1'
        'RIGHTSMANAGEMENT_CE'                                 = 'Azure Information Protection Premium P1 (Nonprofit Staff Pricing)'
        'RIGHTSMANAGEMENT_FACULTY'                            = 'Azure Information Protection Premium P1 for Faculty'
        'RIGHTSMANAGEMENT_STUDENT'                            = 'Azure Information Protection Premium P1 for Students'
        'SHAREPOINTENTERPRISE'                                = 'SharePoint Online (Plan 2)'
        'SHAREPOINTSTANDARD'                                  = 'SharePoint Online (Plan 1)'
        'SHAREPOINTSTORAGE'                                   = 'Office 365 Extra File Storage'
        'SHAREPOINTSTORAGE_FACULTY'                           = 'Office 365 Extra File Storage for faculty'
        'SMB_APPS'                                            = 'Business Apps (free)'
        'SPB'                                                 = 'Microsoft 365 Business'
        'SPE_E3'                                              = 'Microsoft 365 E3'
        'SPE_E5'                                              = 'Microsoft 365 E5'
        'SPE_F1'                                              = 'Microsoft 365 F3'
        'STANDARDPACK'                                        = 'Office 365 E1'
        'STANDARDWOFFPACK_FACULTY'                            = 'Office 365 A1 for faculty'
        'STANDARDWOFFPACK_FACULTY_DEVICE'                     = 'Office 365 A1 for faculty (for Device)'
        'STANDARDWOFFPACK_STUDENT'                            = 'Office 365 A1 for students'
        'STANDARDWOFFPACK_STUDENT_DEVICE'                     = 'Office 365 A1 for students (for Device)'
        'STREAM_P2_ADDON'                                     = 'Microsoft Stream Plan 2 for Office 365 Add-On'
        'STREAM_P2_ADDON_FACULTY'                             = 'Microsoft Stream Plan 2 for Office 365 Add-On for faculty'
        'STREAM_P2_ADDON_STUDENT'                             = 'Microsoft Stream Plan 2 for Office 365 Add-On for students'
        'STREAMSTORAGE_500'                                   = 'Microsoft Stream Storage Add-On (500 GB)'
        'STREAMSTORAGE_500_FACULTY'                           = 'Microsoft Stream Storage Add-On (500 GB) for faculty'
        'STREAMSTORAGE_500_STUDENT'                           = 'Microsoft Stream Storage Add-On (500 GB) for students'
        'TEAMS_COMMERCIAL_TRIAL'                              = 'Microsoft Teams Commercial Cloud (User Initiated) Trial'
        'THREAT_INTELLIGENCE'                                 = 'Office 365 Advanced Threat Protection (Plan 2)'
        'THREAT_INTELLIGENCE_FAC'                             = 'Office 365 Advanced Threat Protection (Plan 2) for faculty'
        'THREAT_INTELLIGENCE_STU'                             = 'Office 365 Advanced Threat Protection (Plan 2) for students'
        'VIRTUAL_AGENT_ADDON'                                 = 'Chat session for Virtual Agent'
        'VIRTUAL_AGENT_BASE'                                  = 'Power Virtual Agent'
        'VIRTUAL_AGENT_USL'                                   = 'Power Virtual Agent User License'
        'VISIOCLIENT'                                         = 'Visio Plan 2'
        'VISIOCLIENT_FACULTY'                                 = 'Visio Plan 2 for faculty'
        'VISIOCLIENT_STUDENT'                                 = 'Visio Plan 2 for students'
        'VISIOONLINE_PLAN1'                                   = 'Visio Plan 1'
        'VISIOONLINE_PLAN1_FAC'                               = 'Visio Plan 1 for faculty'
        'VISIOONLINE_PLAN1_STU'                               = 'Visio Plan 1 for students'
        'WACONEDRIVEENTERPRISE'                               = 'OneDrive for Business (Plan 2)'
        'WACONEDRIVESTANDARD'                                 = 'OneDrive for Business (Plan 1)'
        'WIN10_A3_STUB'                                       = 'Windows 10 Enterprise A3 for students use benefit'
        'WIN10_A5_STUB'                                       = 'Windows 10 Enterprise A5 for student use benefits'
        'WIN10_ENT_A3_FAC'                                    = 'Windows 10 Enterprise A3 for faculty'
        'WIN10_ENT_A3_STU'                                    = 'Windows 10 Enterprise A3 for students'
        'WIN10_ENT_A5_FAC'                                    = 'Windows 10 Enterprise A5 for faculty'
        'WIN10_ENT_A5_STU'                                    = 'Windows 10 Enterprise A5 for students'
        'Win10_VDA_E3'                                        = 'Windows 10 Enterprise E3'
        'WIN10_VDA_E5'                                        = 'Windows 10 Enterprise E5'
    
        # Additional items found from Google searching SkuPartNumber's for Offer names:
        AAD_BASIC                                             = 'Azure Active Directory Basic'
        CRMPLAN2                                              = 'Microsoft Dynamics CRM Online Basic'
        CRMSTANDARD                                           = 'Microsoft Dynamics CRM Online'
        DEVELOPERPACK                                         = 'Office 365 Enterprise E3 Developer'
        DYN365_ENTERPRISE_PLAN1                               = 'Dynamics 365 Customer Engagement Plan Enterprise Edition'
        DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE               = 'Dynamics 365 for Sales and Customer Service Enterprise Edition'
        DYN365_ENTERPRISE_TEAM_MEMBERS                        = 'Dynamics 365 for Team Members Enterprise Edition'
        DYN365_FINANCIALS_BUSINESS_SKU                        = 'Dynamics 365 for Financials Business Edition'
        Dynamics_365_for_Operations                           = 'Dynamics 365 UNF OPS Plan ENT Edition'
        ENTERPRISEPACK_USGOV_DOD                              = 'Office 365 E3_USGOV_DOD'
        ENTERPRISEPACK_USGOV_GCCHIGH                          = 'Office 365 E3_USGOV_GCCHIGH'
        ENTERPRISEWITHSCAL                                    = 'Office 365 Enterprise E4'
        EXCHANGE_S_ESSENTIALS                                 = 'Exchange Online Essentials'
        EXCHANGEESSENTIALS                                    = 'Exchange Online Essentials'
        EXCHANGETELCO                                         = 'Exchange Online POP'
        IT_ACADEMY_AD                                         = 'MS IMAGINE ACADEMY'
        LITEPACK                                              = 'Office 365 Small Business'
        LITEPACK_P2                                           = 'Office 365 Small Business Premium'
        MCOIMP                                                = 'Skype for Business Online (Plan 1)'
        MCOPSTN5                                              = 'Skype for Business PSTN Domestic Calling (120 Minutes)'
        MCOSTANDARD                                           = 'Skype for Business Online (Plan 2)'
        MIDSIZEPACK                                           = 'Office 365 Midsize Business'
        POWER_BI_ADDON                                        = 'Power BI for Office 365 Add-On'
        PROJECTCLIENT                                         = 'Project for Office 365'
        PROJECTONLINE_PLAN_1                                  = 'Project Online Premium without Project Client'
        PROJECTONLINE_PLAN_2                                  = 'Project Online with Project for Office 365'
        SMB_BUSINESS                                          = 'Office 365 Business'
        SMB_BUSINESS_ESSENTIALS                               = 'Office 365 Business Essentials'
        SMB_BUSINESS_PREMIUM                                  = 'Office 365 Business Premium'
        SPE_E3_USGOV_DOD                                      = 'Microsoft 365 E3_USGOV_DOD'
        SPE_E3_USGOV_GCCHIGH                                  = 'Microsoft 365 E3_USGOV_GCCHIGH'
        STANDARDWOFFPACK                                      = 'Office 365 Enterprise E2'
        WIN10_PRO_ENT_SUB                                     = 'Windows 10 Enterprise E3'
    
        # Additional items found within various tenants, manually comparing SkuPartNumbers to AAD Portal > Licenses > All Products:
        DYN365_AI_SERVICE_INSIGHTS                            = 'Dynamics 365 Customer Service Insights Trial'
        DYN365_ENTERPRISE_P1_IW                               = 'Dynamics 365 P1 Trial for Information Workers'
        FLOW_FREE                                             = 'Microsoft Flow Free'
        FORMS_PRO                                             = 'Forms Pro Trial'
        MS_TEAMS_IW                                           = 'MS_TEAMS_IW'
        POWER_BI_STANDARD                                     = 'Power BI (free)'
        POWERAPPS_VIRAL                                       = 'Microsoft PowerApps Plan 2 Trial'
        PROJECT_MADEIRA_PREVIEW_IW_SKU                        = 'PROJECT_MADEIRA_PREVIEW_IW_SKU'
        RIGHTSMANAGEMENT_ADHOC                                = 'Rights Management Adhoc'
        SPZA_IW                                               = 'SPZA_IW'
        STREAM                                                = 'STREAM'
        WINDOWS_STORE                                         = 'WINDOWS_STORE'
    }

    # Load hashtable of integer values to RecipientTypeDetails friendly names.
    # Source (link valid on 2019-12-19):
    # -> https://answers.microsoft.com/en-us/msoffice/forum/msoffice_o365admin-mso_exchon-mso_o365b/recipient-type-values/7c2620e5-9870-48ba-b5c2-7772c739c651
    $RecipientTypeDetailsFriendlyNames = @{

        1           = 'UserMailbox'
        2           = 'LinkedMailbox'
        4           = 'SharedMailbox'
        16          = 'RoomMailbox'
        32          = 'EquipmentMailbox'
        128         = 'MailUser'
        2147483648  = 'RemoteUserMailbox'
        8589934592  = 'RemoteRoomMailbox'
        17179869184 = 'RemoteEquipmentMailbox'
        34359738368 = 'RemoteSharedMailbox'
    }

    #EndRegion Reference Material
}

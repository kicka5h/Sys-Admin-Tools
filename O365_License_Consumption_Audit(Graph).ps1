#Get the required params
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [String]$Username,
  
    [Parameter(Mandatory=$true)]
    [String]$Password,

    [Parameter(Mandatory=$true)]
    [String]$tenantId,
  
    [Parameter(Mandatory=$true)]
    [String]$clientId,
  
    [Parameter(Mandatory=$true)]
    [String]$ClientSecret,
    
    [Parameter(Mandatory=$true)]
    [Array]$RequiredLicenses,

    [Parameters(Mandatory=$true)]
    [String]$emailTo
)  

#################################################### CONNECT TO SERVICES ###########################################################################################

# Check connection to O365
function MSOLConnected {
    Get-MsolDomain -ErrorAction SilentlyContinue | out-null
    $result = $?
    return $result
}

if (-not (MSOLConnected)) {
    #  Connect to O365
    Set-ExecutionPolicy RemoteSigned
    [string][ValidateNotNullOrEmpty()]$passwd = $Password
    $secpasswd = ConvertTo-SecureString -String $passwd -AsPlainText -Force
    $cred = New-Object Management.Automation.PSCredential ($Username, $secpasswd)
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic –AllowRedirection
    Import-PSSession $Session
    Connect-MsolService -credential $cred
}

# Set the web request protocol
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Azure AD OAuth Token for Graph API
# Body params
$granttype = 'password'
$scope = 'https://graph.microsoft.com/.default'

# Construct URI
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Construct Body
$body = @{
    client_id     = $clientId
    scope         = $scope
    client_secret = $clientSecret
    grant_type    = $granttype
    username      = $Username
    password      = $Password
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method 'POST' -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

#################################################### AUDIT LICENSE CONSUMPTION ###########################################################################################

# Get all licenses in the tenant
$Filter = "$" + "select=skupartnumber"
$URL = "https://graph.microsoft.com/v1.0/subscribedskus"
$Header = @{
    Authorization = "Bearer $token"
}
$LicenseList = Invoke-WebRequest -Method 'GET' -URI $URL -Headers $Header -UseBasicParsing | ConvertFrom-Json
$LicenseList = $LicenseList.value.id

# Main license consumption check loop 
[array]$Table = $null
ForEach ($License in $LicenseList)
{
    # Get the difference between available licenses and assigned licenses
    $URI = "https://graph.microsoft.com/v1.0/subscribedskus/$License"
    $L = Invoke-WebRequest -Method 'GET' -URI $URI -Headers $Header -UseBasicParsing | ConvertFrom-Json
    $Consumed = $L.consumedUnits
    $Active = $L.prepaidunits.enabled
    $Difference = $Active - $Consumed 

    # For all licenses that we are consuming more that available, or there is less than 1 left
    if ($Difference -le 2)
    {
        # Get the sku part number
        $SkuName = $License.skuPartNumber

        # Hash table containing friendly license name conversions 
        $LicenseName = $(switch($SkuName) {
        'O365_BUSINESS_ESSENTIALS'            { 'Office 365 Business Essentials' }
        'O365_BUSINESS_PREMIUM'               { 'Office 365 Business Premium' }
        'DESKLESSPACK'                        { 'Office 365 (Plan K1)' }
        'DESKLESSWOFFPACK'                    { 'Office 365 (Plan K2)' }
        'LITEPACK'                            { 'Office 365 (Plan P1)' }
        'EXCHANGESTANDARD'                    { 'Office 365 Exchange Online Only' }
        'STANDARDPACK'                        { 'Enterprise Plan E1' }
        'STANDARDWOFFPACK'                    { 'Office 365 (Plan E2)' }
        'ENTERPRISEPACK'                      { 'Enterprise Plan E3' }
        'ENTERPRISEPACKLRG'                   { 'Enterprise Plan E3' }
        'ENTERPRISEWITHSCAL'                  { 'Enterprise Plan E4' }
        'STANDARDPACK_STUDENT'                { 'Office 365 (Plan A1) for Students' }
        'STANDARDWOFFPACKPACK_STUDENT'        { 'Office 365 (Plan A2) for Students' }
        'ENTERPRISEPACK_STUDENT'              { 'Office 365 (Plan A3) for Students' }
        'ENTERPRISEWITHSCAL_STUDENT'          { 'Office 365 (Plan A4) for Students' }
        'STANDARDPACK_FACULTY'                { 'Office 365 (Plan A1) for Faculty' }
        'STANDARDWOFFPACKPACK_FACULTY'        { 'Office 365 (Plan A2) for Faculty' }
        'ENTERPRISEPACK_FACULTY'              { 'Office 365 (Plan A3) for Faculty' }
        'ENTERPRISEWITHSCAL_FACULTY'          { 'Office 365 (Plan A4) for Faculty' }
        'ENTERPRISEPACK_B_PILOT'              { 'Office 365 (Enterprise Preview)' }
        'STANDARD_B_PILOT'                    { 'Office 365 (Small Business Preview)' }
        'VISIOCLIENT'                         { 'Visio Pro Online' }
        'POWER_BI_ADDON'                      { 'Office 365 Power BI Addon' }
        'POWER_BI_INDIVIDUAL_USE'             { 'Power BI Individual User' }
        'POWER_BI_STANDALONE'                 { 'Power BI Stand Alone' }
        'POWER_BI_STANDARD'                   { 'Power-BI Standard' }
        'PROJECTESSENTIALS'                   { 'Project Lite' }
        'PROJECTCLIENT'                       { 'Project Professional' }
        'PROJECTONLINE_PLAN_1'                { 'Project Online' }
        'PROJECTONLINE_PLAN_2'                { 'Project Online and PRO' }
        'ProjectPremium'                      { 'Project Online Premium' }
        'ECAL_SERVICES'                       { 'ECAL' }
        'EMS'                                 { 'Enterprise Mobility Suite' }
        'RIGHTSMANAGEMENT_ADHOC'              { 'Windows Azure Rights Management' }
        'MCOMEETADV'                          { 'PSTN conferencing' }
        'SHAREPOINTSTORAGE'                   { 'SharePoint storage' }
        'PLANNERSTANDALONE'                   { 'Planner Standalone' }
        'CRMIUR'                              { 'CMRIUR' }
        'BI_AZURE_P1'                         { 'Power BI Reporting and Analytics' }
        'INTUNE_A'                            { 'Windows Intune Plan A' }
        'PROJECTWORKMANAGEMENT'               { 'Office 365 Planner Preview' }
        'ATP_ENTERPRISE'                      { 'Exchange Online Advanced Threat Protection' }
        'EQUIVIO_ANALYTICS'                   { 'Office 365 Advanced eDiscovery' }
        'AAD_BASIC'                           { 'Azure Active Directory Basic' }
        'RMS_S_ENTERPRISE'                    { 'Azure Active Directory Rights Management' }
        'AAD_PREMIUM'                         { 'Azure Active Directory Premium' }
        'MFA_PREMIUM'                         { 'Azure Multi-Factor Authentication' }
        'STANDARDPACK_GOV'                    { 'Microsoft Office 365 (Plan G1) for Government' }
        'STANDARDWOFFPACK_GOV'                { 'Microsoft Office 365 (Plan G2) for Government' }
        'ENTERPRISEPACK_GOV'                  { 'Microsoft Office 365 (Plan G3) for Government' }
        'ENTERPRISEWITHSCAL_GOV'              { 'Microsoft Office 365 (Plan G4) for Government' }
        'DESKLESSPACK_GOV'                    { 'Microsoft Office 365 (Plan K1) for Government' }
        'ESKLESSWOFFPACK_GOV'                 { 'Microsoft Office 365 (Plan K2) for Government' }
        'EXCHANGESTANDARD_GOV'                { 'Microsoft Office 365 Exchange Online (Plan 1) only for Government' }
        'EXCHANGEENTERPRISE _GOV'             { 'Microsoft Office 365 Exchange Online (Plan 2) only for Government' }
        'SHAREPOINTDESKLESS _GOV'             { 'SharePoint Online Kiosk' }
        'EXCHANGE_S_DESKLESS_GOV'             { 'Exchange Kiosk' }
        'RMS_S_ENTERPRISE_GOV'                { 'Windows Azure Active Directory Rights Management' }
        'OFFICESUBSCRIPTION_GOV'              { 'Office ProPlus' }
        'MCOSTANDARD_GOV'                     { 'Lync Plan 2G' }
        'SHAREPOINTWAC_GOV'                   { 'Office Online for Government' }
        'SHAREPOINTENTERPRISE_GOV'            { 'SharePoint Plan 2G' }
        'EXCHANGE_S_ENTERPRISE_GOV'           { 'Exchange Plan 2G' }
        'EXCHANGE_S_ARCHIVE_ADDON_GOV'        { 'Exchange Online Archiving' }
        'EXCHANGE_S_DESKLESS_GOV'             { 'Exchange Online Kiosk' }
        'SHAREPOINTDESKLESS'                  { 'SharePoint Online Kiosk' }
        'SHAREPOINTWAC_GOV'                   { 'Office Online' }
        'YAMMER_ENTERPRISE_STANDALONE'        { 'Yammer for the Starship Enterprise' }
        'EXCHANGE_L_STANDARD'                 { 'Exchange Online (Plan 1)' }
        'MCOLITE'                             { 'Lync Online (Plan 1)' }
        'SHAREPOINTLITE'                      { 'SharePoint Online (Plan 1)' }
        'OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ'  { 'Office ProPlus' }
        'EXCHANGE_S_STANDARD_MIDMARKET'       { 'Exchange Online (Plan 1)' }
        'MCOSTANDARD_MIDMARKET'               { 'Lync Online (Plan 1)' }
        'SHAREPOINTENTERPRISE_MIDMARKET'      { 'SharePoint Online (Plan 1)' }
        'OFFICESUBSCRIPTION'                  { 'Office ProPlus' }
        'YAMMER_MIDSIZE'                      { 'Yammer' }
        'DYN365_ENTERPRISE_PLAN1'             { 'Dynamics 365 Customer Engagement Plan Enterprise Edition' }
        'ENTERPRISEPREMIUM_NOPSTNCONF'        { 'Enterprise E5 (without Audio Conferencing)' }
        'ENTERPRISEPREMIUM'                   { 'M365 E5' }
        'MCOSTANDARD'                         { 'Skype for Business Online Standalone Plan 2' }
        'PROJECT_MADEIRA_PREVIEW_IW_SKU'      { 'Dynamics 365 for Financials for IWs' }
        'STANDARDWOFFPACK_IW_STUDENT'         { 'Office 365 Education for Students' }
        'STANDARDWOFFPACK_IW_FACULTY'         { 'Office 365 Education for Faculty' }
        'EOP_ENTERPRISE_FACULTY'              { 'Exchange Online Protection for Faculty' }
        'EXCHANGESTANDARD_STUDENT'            { 'Exchange Online (Plan 1) for Students' }
        'OFFICESUBSCRIPTION_STUDENT'          { 'Office ProPlus Student Benefit' }
        'STANDARDWOFFPACK_FACULTY'            { 'Office 365 Education E1 for Faculty' }
        'STANDARDWOFFPACK_STUDENT'            { 'Microsoft Office 365 (Plan A2) for Students' }
        'DYN365_FINANCIALS_BUSINESS_SKU'      { 'Dynamics 365 for Financials Business Edition' }
        'DYN365_FINANCIALS_TEAM_MEMBERS_SKU'  { 'Dynamics 365 for Team Members Business Edition' }
        'FLOW_FREE'                           { 'Microsoft Flow Free' }
        'POWER_BI_PRO'                        { 'Power BI Pro' }
        'O365_BUSINESS'                       { 'Office 365 Business' }
        'DYN365_ENTERPRISE_SALES'             { 'Dynamics Office 365 Enterprise Sales' }
        'RIGHTSMANAGEMENT'                    { 'Rights Management' }
        'PROJECTPROFESSIONAL'                 { 'Project Professional' }
        'VISIOONLINE_PLAN1'                   { 'Visio Online Plan 1' }
        'EXCHANGEENTERPRISE'                  { 'Exchange Online Plan 2' }
        'DYN365_ENTERPRISE_P1_IW'             { 'Dynamics 365 P1 Trial for Information Workers' }
        'DYN365_ENTERPRISE_TEAM_MEMBERS'      { 'Dynamics 365 For Team Members Enterprise Edition' }
        'CRMSTANDARD'                         { 'Microsoft Dynamics CRM Online Professional' }
        'EXCHANGE_S_ARCHIVE_ADDON_GOV'        { 'Exchange Online Archiving For Exchange Online' }
        'EXCHANGEDESKLESS'                    { 'Exchange Online Kiosk' }
        'SPZA_IW'                             { 'App Connect' }
        'WINDOWS_STORE'                       { 'Windows Store for Business' }
        'MCOEV'                               { 'Microsoft Phone System' }
        'VIDEO_INTEROP'                       { 'Polycom Skype Meeting Video Interop for Skype for Business' }
        'SPE_E5'                              { 'Microsoft 365 E5' }
        'SPE_E3'                              { 'Microsoft 365 E3' }
        'ATA'                                 { 'Advanced Threat Analytics' }
        'MCOPSTN2'                            { 'Domestic and International Calling Plan' }
        'FLOW_P1'                             { 'Microsoft Flow Plan 1' }
        'FLOW_P2'                             { 'Microsoft Flow Plan 2' }
        'EMSPREMIUM'                          { 'Enterprise Mobility + Security E5' }
        'MCOPSTN1'                            { 'Skype for Business PSTN Domestic Calling' }
        'EXCHANGE_S_ENTERPRISE'               { 'Exchange Online (Plan 2)' }
        'CRMPLAN2'                            { 'Dynamics CRM Online Plan 2' }
        'DESKLESS'                            { 'Microsoft StaffHub' }
        'INTUNE_A_VL'                         { 'Intune (Volume License)' }
        'IT_ACADEMY_AD'                       { 'Microsoft Imagine Academy' }
        'OFFICESUBSCRIPTION_FACULTY'          { 'Office 365 ProPlus for Faculty' }
        'POWER_BI_INDIVIDUAL_USER'            { 'Power BI for Office 365 Individual' }
        'POWERAPPS_INDIVIDUAL_USER'           { 'Microsoft PowerApps and Logic flows' }
        'PROJECTONLINE_PLAN_1_FACULTY'        { 'Project Online for Faculty Plan 1' }
        'PROJECTONLINE_PLAN_1_STUDENT'        { 'Project Online for Students Plan 1' }
        'PROJECTONLINE_PLAN_2_FACULTY'        { 'Project Online for Faculty Plan 2' }
        'PROJECTONLINE_PLAN_2_STUDENT'        { 'Project Online for Students Plan 2' }
        'RIGHTSMANAGEMENT_STANDARD_FACULTY'   { 'Information Rights Management for Faculty' }
        'RIGHTSMANAGEMENT_STANDARD_STUDENT'   { 'Information Rights Management for Students' }
        'SHAREPOINTSTANDARD'                  { 'SharePoint Online Plan 1' }
        'WACSHAREPOINTSTD'                    { 'Office Online' }
        'SHAREPOINTENTERPRISE'                { 'SharePoint Online (Plan 2)' }
        'STREAM'                              { 'Microsoft Stream' }
        'SWAY'                                { 'SWAY' }
        'INTUNE_O365'                         { 'INTUNE' }
        'MCOVOICECONF'                        { 'Lync Online (Plan 3)' }
        'Office 365 Midsize Business'         { 'MIDSIZEPACK' }
        'PROJECT_CLIENT_SUBSCRIPTION'         { 'Project Pro for Office 365' }
        'VISIO_CLIENT_SUBSCRIPTION'           { 'Visio Pro for Office 365' }
        'CRMTESTINSTANCE'                     { 'CRM Test Instance' }
        'ONEDRIVESTANDARD'                    { 'OneDrive' }
        'WACONEDRIVESTANDARD'                 { 'OneDrive Pack' }
        'SQL_IS_SSIM'                         { 'Power BI Information Services' }
        'EOP_ENTERPRISE_FACULTY'              { 'Exchange Online Protection' }
        'PROJECT_ESSENTIALS'                  { 'Project Lite' }
        'NBPROFESSIONALFORCRM'                { 'Microsoft Social Listening Professional' }
        'POWERAPPS_VIRAL'                     { 'Microsoft Power Apps & Flow' }
        'BI_AZURE_P2'                         { 'Power BI Pro' }
        'CRMINSTANCE'                         { 'Dynamics CRM Online Additional Production Instance' }
        'CRMPLAN1'                            { 'Dynamics CRM Online Essential' }
        'CRMSTORAGE'                          { 'Dynamics CRM Online Additional Storage' }
        'DESKLESSPACK_YAMME'                  { 'Office 365 (Plan K1) with Yammer' }
        'DESKLESSWOFFPACK_GOV'                { 'Office 365 (Plan K2) for Government' }
        'ENTERPRISEPACKWSCAL'                 { 'Office 365 (Plan E4)' }
        'EXCHANGE_ANALYTICS'                  { 'Delve Analytics' }
        'EXCHANGE_S_STANDARD'                 { 'Exchange Online (Plan 2)' }
        'EXCHANGEARCHIVE'                     { 'Exchange Online Archiving' }
        'EXCHANGETELCO'                       { 'Exchange Online POP' }
        'INTUNE_STORAGE'                      { 'Intune Extra Storage' }
        'LITEPACK_P2'                         { 'Office 365 Small Business Premium' }
        'LOCKBOX'                             { 'Customer Lockbox' }
        'LOCKBOX_ENTERPRISE'                  { 'Customer Lockbox' }
        'MCOIMP'                              { 'Skype for Business Online (Plan 1)' }
        'MCOPLUSCAL'                          { 'Skype for Business Plus CAL' }
        'MCOPSTNC'                            { 'Skype for Business Communication Credits - None?' }
        'MCOPSTNPP'                           { 'Skype for Business Communication Credits - Paid?' }
        'MCOVOICECONF'                        { 'Lync Online (Plan 3)' }
        'MICROSOFT_BUSINESS_CENTER'           { 'Microsoft Business Center' }
        'MIDSIZEPACK'                         { 'Office 365 Midsize Business' }
        'MS-AZR-0145P'                        { 'Azure' }
        'NBPOSTS'                             { 'Social Engagement Additional 10K Posts' }
        'PARATURE_ENTERPRISE'                 { 'Parature Enterprise' }
        'PARATURE_FILESTORAGE_ADDON'          { 'Parature File Storage Addon' }
        'PARATURE_SUPPORT_ENHANCED'           { 'Parature Support Enhanced' }
        'PROJECTONLINE_PLAN1_FACULTY'         { 'Project Online for Faculty' }
        'PROJECTONLINE_PLAN1_STUDENT'         { 'Project Online for Students' }
        'SHAREPOINT_PROJECT_EDU'              { 'Project Online for Education' }
        'SHAREPOINTENTERPRISE_EDU'            { 'SharePoint (Plan 2) for EDU' }
        'SHAREPOINTSTANDARD_YAMMER'           { 'Sharepoint Standard with Yammer' }
        'SHAREPOINTWAC_EDU'                   { 'Office Online for Education' }
        'WACONEDRIVEENTERPRISE'               { 'OneDrive for Business (Plan 2)' }
        'WACSHAREPOINTENT'                    { 'Office Web Apps with SharePoint (Plan 2)' }
        'YAMMER_ENTERPRISE_STANDALONE'        { 'Yammer Enterprise' }
        'TEAMS_COMMERCIAL_TRIAL'              { 'Teams Legacy Trial License' }
        'Ax7_USER_TRIAL'                      { 'Microsoft Dynamics AX7 User Trial' }
        'POWERFLOW_P2'                        { 'Microsoft PowerApps Plan 2' }
        'WIN_DEF_ATP'                         { 'Windows Defender Advanced Threat Protection' }
        'CRM_ONLINE_PORTAL'                   { 'CRM Online' }
        'MEETING_ROOM'                        { 'Surface Hub Meeting License' }
        'ADALLOM_STANDALONE'                  { 'Adallom Standalone' }

        default {'License Name Unknown'}
})

        # Create the table for the email
        [array]$Table += [PSCustomObject]@{
            
            # Account for any accidental duplicates by converting collections into readable lists
            License = (@($LicenseName) -join ',')
            "Remaining" = $Difference
             }
    }

        # If the friendly license name is missing from the hash table, replace it with the sku ID
        ForEach ($T in $Table.License)
        {
            if ($T -like "License Name Unknown")
            {
                $T = $License.AccountSkuId
            }
        }
}

#################################################### AUDIT LICENSE ASSIGNMENT ###########################################################################################

# Gets all employee accounts so that we can limit the list to just employee accounts 
$Filter = "$" + "Select=mail"
$URL = "https://graph.microsoft.com/v1.0/users?$Filter"
$Header = @{
    Authorization = "Bearer $token"
    'Content-type' = 'application/json'
}
$UserList = Invoke-WebRequest -Method 'GET' -Uri $URL -Headers $Header -UseBasicParsing | ConvertFrom-Json
$Users = $UserList.value | Select-Object -ExpandProperty Mail

# Main license assignment check loop
[array]$Table2 = $null
[array]$RequiredLicenses = $null
ForEach ($User in $Users)
{
    $O365Users = Get-MsolUser -UserPrincipalName $User | Select-Object DisplayName, Licenses 
    
    ForEach ($O365User in $O365Users)
    {
        $UserLicense = $O365User.Licenses
        $ULicense = $UserLicense.AccountSkuID
        $UserName = $O365User.DisplayName

        If ($ULicense -notcontains $RequiredLicenses)
        {             
            #Compare the list of asigned licenses to the above to find out which is missing 
            $MissingLicenses = $RequiredLicenses | ?{$ULicense -notcontains $_}

            $Missing = ForEach ($Missing in $MissingLicenses)
            {
                # Split the tenant name from the license sku
                $L, $MLicenses = $Missing.Split(':')
                $MLicenses
            }

            # Hash table containing friendly license name conversions 
            $MissingLicenseList = $(switch($Missing) {
            'O365_BUSINESS_ESSENTIALS'            { 'Office 365 Business Essentials' }
            'O365_BUSINESS_PREMIUM'               { 'Office 365 Business Premium' }
            'DESKLESSPACK'                        { 'Office 365 (Plan K1)' }
            'DESKLESSWOFFPACK'                    { 'Office 365 (Plan K2)' }
            'LITEPACK'                            { 'Office 365 (Plan P1)' }
            'EXCHANGESTANDARD'                    { 'Office 365 Exchange Online Only' }
            'STANDARDPACK'                        { 'Enterprise Plan E1' }
            'STANDARDWOFFPACK'                    { 'Office 365 (Plan E2)' }
            'ENTERPRISEPACK'                      { 'Enterprise Plan E3' }
            'ENTERPRISEPACKLRG'                   { 'Enterprise Plan E3' }
            'ENTERPRISEWITHSCAL'                  { 'Enterprise Plan E4' }
            'STANDARDPACK_STUDENT'                { 'Office 365 (Plan A1) for Students' }
            'STANDARDWOFFPACKPACK_STUDENT'        { 'Office 365 (Plan A2) for Students' }
            'ENTERPRISEPACK_STUDENT'              { 'Office 365 (Plan A3) for Students' }
            'ENTERPRISEWITHSCAL_STUDENT'          { 'Office 365 (Plan A4) for Students' }
            'STANDARDPACK_FACULTY'                { 'Office 365 (Plan A1) for Faculty' }
            'STANDARDWOFFPACKPACK_FACULTY'        { 'Office 365 (Plan A2) for Faculty' }
            'ENTERPRISEPACK_FACULTY'              { 'Office 365 (Plan A3) for Faculty' }
            'ENTERPRISEWITHSCAL_FACULTY'          { 'Office 365 (Plan A4) for Faculty' }
            'ENTERPRISEPACK_B_PILOT'              { 'Office 365 (Enterprise Preview)' }
            'STANDARD_B_PILOT'                    { 'Office 365 (Small Business Preview)' }
            'VISIOCLIENT'                         { 'Visio Pro Online' }
            'POWER_BI_ADDON'                      { 'Office 365 Power BI Addon' }
            'POWER_BI_INDIVIDUAL_USE'             { 'Power BI Individual User' }
            'POWER_BI_STANDALONE'                 { 'Power BI Stand Alone' }
            'POWER_BI_STANDARD'                   { 'Power-BI Standard' }
            'PROJECTESSENTIALS'                   { 'Project Lite' }
            'PROJECTCLIENT'                       { 'Project Professional' }
            'PROJECTONLINE_PLAN_1'                { 'Project Online' }
            'PROJECTONLINE_PLAN_2'                { 'Project Online and PRO' }
            'ProjectPremium'                      { 'Project Online Premium' }
            'ECAL_SERVICES'                       { 'ECAL' }
            'EMS'                                 { 'Enterprise Mobility Suite' }
            'RIGHTSMANAGEMENT_ADHOC'              { 'Windows Azure Rights Management' }
            'MCOMEETADV'                          { 'PSTN conferencing' }
            'SHAREPOINTSTORAGE'                   { 'SharePoint storage' }
            'PLANNERSTANDALONE'                   { 'Planner Standalone' }
            'CRMIUR'                              { 'CMRIUR' }
            'BI_AZURE_P1'                         { 'Power BI Reporting and Analytics' }
            'INTUNE_A'                            { 'Windows Intune Plan A' }
            'PROJECTWORKMANAGEMENT'               { 'Office 365 Planner Preview' }
            'ATP_ENTERPRISE'                      { 'Exchange Online Advanced Threat Protection' }
            'EQUIVIO_ANALYTICS'                   { 'Office 365 Advanced eDiscovery' }
            'AAD_BASIC'                           { 'Azure Active Directory Basic' }
            'RMS_S_ENTERPRISE'                    { 'Azure Active Directory Rights Management' }
            'AAD_PREMIUM'                         { 'Azure Active Directory Premium' }
            'MFA_PREMIUM'                         { 'Azure Multi-Factor Authentication' }
            'STANDARDPACK_GOV'                    { 'Microsoft Office 365 (Plan G1) for Government' }
            'STANDARDWOFFPACK_GOV'                { 'Microsoft Office 365 (Plan G2) for Government' }
            'ENTERPRISEPACK_GOV'                  { 'Microsoft Office 365 (Plan G3) for Government' }
            'ENTERPRISEWITHSCAL_GOV'              { 'Microsoft Office 365 (Plan G4) for Government' }
            'DESKLESSPACK_GOV'                    { 'Microsoft Office 365 (Plan K1) for Government' }
            'ESKLESSWOFFPACK_GOV'                 { 'Microsoft Office 365 (Plan K2) for Government' }
            'EXCHANGESTANDARD_GOV'                { 'Microsoft Office 365 Exchange Online (Plan 1) only for Government' }
            'EXCHANGEENTERPRISE _GOV'             { 'Microsoft Office 365 Exchange Online (Plan 2) only for Government' }
            'SHAREPOINTDESKLESS _GOV'             { 'SharePoint Online Kiosk' }
            'EXCHANGE_S_DESKLESS_GOV'             { 'Exchange Kiosk' }
            'RMS_S_ENTERPRISE_GOV'                { 'Windows Azure Active Directory Rights Management' }
            'OFFICESUBSCRIPTION_GOV'              { 'Office ProPlus' }
            'MCOSTANDARD_GOV'                     { 'Lync Plan 2G' }
            'SHAREPOINTWAC_GOV'                   { 'Office Online for Government' }
            'SHAREPOINTENTERPRISE_GOV'            { 'SharePoint Plan 2G' }
            'EXCHANGE_S_ENTERPRISE_GOV'           { 'Exchange Plan 2G' }
            'EXCHANGE_S_ARCHIVE_ADDON_GOV'        { 'Exchange Online Archiving' }
            'EXCHANGE_S_DESKLESS_GOV'             { 'Exchange Online Kiosk' }
            'SHAREPOINTDESKLESS'                  { 'SharePoint Online Kiosk' }
            'SHAREPOINTWAC_GOV'                   { 'Office Online' }
            'YAMMER_ENTERPRISE_STANDALONE'        { 'Yammer for the Starship Enterprise' }
            'EXCHANGE_L_STANDARD'                 { 'Exchange Online (Plan 1)' }
            'MCOLITE'                             { 'Lync Online (Plan 1)' }
            'SHAREPOINTLITE'                      { 'SharePoint Online (Plan 1)' }
            'OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ'  { 'Office ProPlus' }
            'EXCHANGE_S_STANDARD_MIDMARKET'       { 'Exchange Online (Plan 1)' }
            'MCOSTANDARD_MIDMARKET'               { 'Lync Online (Plan 1)' }
            'SHAREPOINTENTERPRISE_MIDMARKET'      { 'SharePoint Online (Plan 1)' }
            'OFFICESUBSCRIPTION'                  { 'Office ProPlus' }
            'YAMMER_MIDSIZE'                      { 'Yammer' }
            'DYN365_ENTERPRISE_PLAN1'             { 'Dynamics 365 Customer Engagement Plan Enterprise Edition' }
            'ENTERPRISEPREMIUM_NOPSTNCONF'        { 'Enterprise E5 (without Audio Conferencing)' }
            'ENTERPRISEPREMIUM'                   { 'Enterprise E5 (with Audio Conferencing)' }
            'MCOSTANDARD'                         { 'Skype for Business Online Standalone Plan 2' }
            'PROJECT_MADEIRA_PREVIEW_IW_SKU'      { 'Dynamics 365 for Financials for IWs' }
            'STANDARDWOFFPACK_IW_STUDENT'         { 'Office 365 Education for Students' }
            'STANDARDWOFFPACK_IW_FACULTY'         { 'Office 365 Education for Faculty' }
            'EOP_ENTERPRISE_FACULTY'              { 'Exchange Online Protection for Faculty' }
            'EXCHANGESTANDARD_STUDENT'            { 'Exchange Online (Plan 1) for Students' }
            'OFFICESUBSCRIPTION_STUDENT'          { 'Office ProPlus Student Benefit' }
            'STANDARDWOFFPACK_FACULTY'            { 'Office 365 Education E1 for Faculty' }
            'STANDARDWOFFPACK_STUDENT'            { 'Microsoft Office 365 (Plan A2) for Students' }
            'DYN365_FINANCIALS_BUSINESS_SKU'      { 'Dynamics 365 for Financials Business Edition' }
            'DYN365_FINANCIALS_TEAM_MEMBERS_SKU'  { 'Dynamics 365 for Team Members Business Edition' }
            'FLOW_FREE'                           { 'Microsoft Flow Free' }
            'POWER_BI_PRO'                        { 'Power BI Pro' }
            'O365_BUSINESS'                       { 'Office 365 Business' }
            'DYN365_ENTERPRISE_SALES'             { 'Dynamics Office 365 Enterprise Sales' }
            'RIGHTSMANAGEMENT'                    { 'Rights Management' }
            'PROJECTPROFESSIONAL'                 { 'Project Professional' }
            'VISIOONLINE_PLAN1'                   { 'Visio Online Plan 1' }
            'EXCHANGEENTERPRISE'                  { 'Exchange Online Plan 2' }
            'DYN365_ENTERPRISE_P1_IW'             { 'Dynamics 365 P1 Trial for Information Workers' }
            'DYN365_ENTERPRISE_TEAM_MEMBERS'      { 'Dynamics 365 For Team Members Enterprise Edition' }
            'CRMSTANDARD'                         { 'Microsoft Dynamics CRM Online Professional' }
            'EXCHANGE_S_ARCHIVE_ADDON_GOV'        { 'Exchange Online Archiving For Exchange Online' }
            'EXCHANGEDESKLESS'                    { 'Exchange Online Kiosk' }
            'SPZA_IW'                             { 'App Connect' }
            'WINDOWS_STORE'                       { 'Windows Store for Business' }
            'MCOEV'                               { 'Microsoft Phone System' }
            'VIDEO_INTEROP'                       { 'Polycom Skype Meeting Video Interop for Skype for Business' }
            'SPE_E5'                              { 'Microsoft 365 E5' }
            'SPE_E3'                              { 'Microsoft 365 E3' }
            'ATA'                                 { 'Advanced Threat Analytics' }
            'MCOPSTN2'                            { 'Domestic and International Calling Plan' }
            'FLOW_P1'                             { 'Microsoft Flow Plan 1' }
            'FLOW_P2'                             { 'Microsoft Flow Plan 2' }
            'EMSPREMIUM'                          { 'Enterprise Mobility + Security E5' }
            'MCOPSTN1'                            { 'Skype for Business PSTN Domestic Calling' }
            'EXCHANGE_S_ENTERPRISE'               { 'Exchange Online (Plan 2)' }
            'CRMPLAN2'                            { 'Dynamics CRM Online Plan 2' }
            'DESKLESS'                            { 'Microsoft StaffHub' }
            'INTUNE_A_VL'                         { 'Intune (Volume License)' }
            'IT_ACADEMY_AD'                       { 'Microsoft Imagine Academy' }
            'OFFICESUBSCRIPTION_FACULTY'          { 'Office 365 ProPlus for Faculty' }
            'POWER_BI_INDIVIDUAL_USER'            { 'Power BI for Office 365 Individual' }
            'POWERAPPS_INDIVIDUAL_USER'           { 'Microsoft PowerApps and Logic flows' }
            'PROJECTONLINE_PLAN_1_FACULTY'        { 'Project Online for Faculty Plan 1' }
            'PROJECTONLINE_PLAN_1_STUDENT'        { 'Project Online for Students Plan 1' }
            'PROJECTONLINE_PLAN_2_FACULTY'        { 'Project Online for Faculty Plan 2' }
            'PROJECTONLINE_PLAN_2_STUDENT'        { 'Project Online for Students Plan 2' }
            'RIGHTSMANAGEMENT_STANDARD_FACULTY'   { 'Information Rights Management for Faculty' }
            'RIGHTSMANAGEMENT_STANDARD_STUDENT'   { 'Information Rights Management for Students' }
            'SHAREPOINTSTANDARD'                  { 'SharePoint Online Plan 1' }
            'WACSHAREPOINTSTD'                    { 'Office Online' }
            'SHAREPOINTENTERPRISE'                { 'SharePoint Online (Plan 2)' }
            'STREAM'                              { 'Microsoft Stream' }
            'SWAY'                                { 'SWAY' }
            'INTUNE_O365'                         { 'INTUNE' }
            'MCOVOICECONF'                        { 'Lync Online (Plan 3)' }
            'Office 365 Midsize Business'         { 'MIDSIZEPACK' }
            'PROJECT_CLIENT_SUBSCRIPTION'         { 'Project Pro for Office 365' }
            'VISIO_CLIENT_SUBSCRIPTION'           { 'Visio Pro for Office 365' }
            'CRMTESTINSTANCE'                     { 'CRM Test Instance' }
            'ONEDRIVESTANDARD'                    { 'OneDrive' }
            'WACONEDRIVESTANDARD'                 { 'OneDrive Pack' }
            'SQL_IS_SSIM'                         { 'Power BI Information Services' }
            'EOP_ENTERPRISE_FACULTY'              { 'Exchange Online Protection' }
            'PROJECT_ESSENTIALS'                  { 'Project Lite' }
            'NBPROFESSIONALFORCRM'                { 'Microsoft Social Listening Professional' }
            'POWERAPPS_VIRAL'                     { 'Microsoft Power Apps & Flow' }
            'BI_AZURE_P2'                         { 'Power BI Pro' }
            'CRMINSTANCE'                         { 'Dynamics CRM Online Additional Production Instance' }
            'CRMPLAN1'                            { 'Dynamics CRM Online Essential' }
            'CRMSTORAGE'                          { 'Dynamics CRM Online Additional Storage' }
            'DESKLESSPACK_YAMME'                  { 'Office 365 (Plan K1) with Yammer' }
            'DESKLESSWOFFPACK_GOV'                { 'Office 365 (Plan K2) for Government' }
            'ENTERPRISEPACKWSCAL'                 { 'Office 365 (Plan E4)' }
            'EXCHANGE_ANALYTICS'                  { 'Delve Analytics' }
            'EXCHANGE_S_STANDARD'                 { 'Exchange Online (Plan 2)' }
            'EXCHANGEARCHIVE'                     { 'Exchange Online Archiving' }
            'EXCHANGETELCO'                       { 'Exchange Online POP' }
            'INTUNE_STORAGE'                      { 'Intune Extra Storage' }
            'LITEPACK_P2'                         { 'Office 365 Small Business Premium' }
            'LOCKBOX'                             { 'Customer Lockbox' }
            'LOCKBOX_ENTERPRISE'                  { 'Customer Lockbox' }
            'MCOIMP'                              { 'Skype for Business Online (Plan 1)' }
            'MCOPLUSCAL'                          { 'Skype for Business Plus CAL' }
            'MCOPSTNC'                            { 'Skype for Business Communication Credits - None?' }
            'MCOPSTNPP'                           { 'Skype for Business Communication Credits - Paid?' }
            'MCOVOICECONF'                        { 'Lync Online (Plan 3)' }
            'MICROSOFT_BUSINESS_CENTER'           { 'Microsoft Business Center' }
            'MIDSIZEPACK'                         { 'Office 365 Midsize Business' }
            'MS-AZR-0145P'                        { 'Azure' }
            'NBPOSTS'                             { 'Social Engagement Additional 10K Posts' }
            'PARATURE_ENTERPRISE'                 { 'Parature Enterprise' }
            'PARATURE_FILESTORAGE_ADDON'          { 'Parature File Storage Addon' }
            'PARATURE_SUPPORT_ENHANCED'           { 'Parature Support Enhanced' }
            'PROJECTONLINE_PLAN1_FACULTY'         { 'Project Online for Faculty' }
            'PROJECTONLINE_PLAN1_STUDENT'         { 'Project Online for Students' }
            'SHAREPOINT_PROJECT_EDU'              { 'Project Online for Education' }
            'SHAREPOINTENTERPRISE_EDU'            { 'SharePoint (Plan 2) for EDU' }
            'SHAREPOINTSTANDARD_YAMMER'           { 'Sharepoint Standard with Yammer' }
            'SHAREPOINTWAC_EDU'                   { 'Office Online for Education' }
            'WACONEDRIVEENTERPRISE'               { 'OneDrive for Business (Plan 2)' }
            'WACSHAREPOINTENT'                    { 'Office Web Apps with SharePoint (Plan 2)' }
            'YAMMER_ENTERPRISE_STANDALONE'        { 'Yammer Enterprise' }
            'TEAMS_COMMERCIAL_TRIAL'              { 'Teams Legacy Trial License' }
            'Ax7_USER_TRIAL'                      { 'Microsoft Dynamics AX7 User Trial' }
            'POWERFLOW_P2'                        { 'Microsoft PowerApps Plan 2' }
            'WIN_DEF_ATP'                         { 'Windows Defender Advanced Threat Protection' }
            'CRM_ONLINE_PORTAL'                   { 'CRM Online' }
            'MEETING_ROOM'                        { 'Surface Hub Meeting License' }
            'ADALLOM_STANDALONE'                  { 'Adallom Standalone' }

            default {'License Name Unknown'}
        })

            # Create the table for the email
            [array]$Table2 += [PSCustomObject]@{
                Name = $UserName
                "Missing Licenses" = (@($MissingLicenseList) -join ', ')
             }
        }
    }
}

#################################################### SEND THE EMAIL AUDIT ###############################################################################################

# Set up the CSS for the email 
$CSS = @"
    table.rosyBrownTable {
        border: 4px solid #BA8697;
        background-color: #555555;
        width: 400px;
        text-align: center;
        border-collapse: collapse;
      }
      table.rosyBrownTable td, table.rosyBrownTable th {
        border: 1px solid #555555;
        padding: 4px 4px;
      }
      table.rosyBrownTable tbody td {
        font-size: 13px;
        font-weight: bold;
        color: #FFFFFF;
      }
      table.rosyBrownTable tr:nth-child(even) {
        background: #BA8697;
      }
      table.rosyBrownTable td:nth-child(even) {
        background: #BA8697;
      }
      table.rosyBrownTable tfoot {
        font-weight: bold;
        background: #BA8697;
        border-top: 1px solid #444444;
      }
      table.rosyBrownTable tfoot .links {
        text-align: right;
      }
      table.rosyBrownTable tfoot .links a{
        display: inline-block;
        background: #FFFFFF;
        color: #BA8697;
        padding: 2px 8px;
        border-radius: 5px;
      }
"@

# Prepare the body of the email. Use Fragment to join the two tables.
$Report = ($Table | Sort-Object License, "Remaining" | Select License, "Remaining" | ConvertTo-Html -Fragment )
$Report2 = ($Table2 | Sort-Object Name, "Missing Licenses" | Select Name, "Missing Licenses" | ConvertTo-Html -Fragment )

$Body = "

$CSS

$Report

<br>

$Report2

<br>
<br>

"

# Get credentials to send the email
[string][ValidateNotNullOrEmpty()]$passwd = $Password
$secpasswd = ConvertTo-SecureString -String $passwd -AsPlainText -Force
$cred = New-Object Management.Automation.PSCredential ($Username, $secpasswd)

# Set up the email params
$Date = Get-Date -Format 'MMMM dd yyyy'
$From = $Username
$Subject = "O365 License Audit " + $Date
$SMTPPort = "587"
$SMTPServer = "smtp.office365.com"

# Send the email
Send-MailMessage -From $From -to $emailTo -Subject $Subject -Body $Body -BodyAsHTML -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $cred

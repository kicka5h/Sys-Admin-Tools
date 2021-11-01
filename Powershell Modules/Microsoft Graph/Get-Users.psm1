Function Get-Users {
    [CmdletBinding()]
    param(
        [switch]$All
    )

    $Token = Get-GraphToken -Application

    Function Get-Pages ($Body, $Headers, $URL) {
        do {
            $Results = Invoke-RestMethod -Headers $Headers -Uri $URL -UseBasicParsing -Method "GET" -ContentType "application/json"
            if ($Results.value) {
                $QueryResults += $Results.value
            }
            else {
                $QueryResults += $Results
            }
            $URL = $Results.'@odata.nextlink'
        } until (!($URL))

        $QueryResults
    }

    #region initialize custom classes
    class User {
        [string]$ObjectID
        [string]$Name
        [string]$PreferredName
        [string]$FirstName
        [string]$LastName
        [string]$Description
        [string]$StartDate
        [string]$Company
        [string]$Class
        [string]$Office
        [string]$Department
        [string]$Title
        [string]$EmployeeType
        [string]$DistinguishedName
        [string]$UserPrincipalName
        [string]$samAccountName
        [string]$Visible
        [string]$Active
        [bool]$IsAdmin
        [string]$AccessLevel
        [string]$PasswordLastChanged
        [System.Collections.ArrayList]$Licenses = @()

        # user constructor
        User($R){
            $this.Name = $R.displayName
            $this.FirstName = $R.givenname
            $this.LastName = $R.surName
            $this.PreferredName = ($R.displayName).Replace($this.LastName, "")
            $this.Description = $R.extension_b797a1cc26244252ac880ab6957387ce_description
            $this.StartDate = $R.extension_b797a1cc26244252ac880ab6957387ce_extensionAttribute2
            $this.Company = $R.companyName
            $this.Office = $R.OfficeLocation
            $this.Class = $R.extension_b797a1cc26244252ac880ab6957387ce_extensionAttribute3
            $this.Department = $R.Department
            $this.Title = $R.jobTitle
            $this.EmployeeType = $R.extension_b797a1cc26244252ac880ab6957387ce_employeeType
            $this.DistinguishedName = $R.onPremisesDistinguishedName
            $this.UserPrincipalName = $R.Userprincipalname
            $this.samAccountName = $R.mailNickname
            $this.Visible = $R.showInAddressList
            $this.ObjectID = $R.id
            $this.PasswordLastChanged = $R.lastPasswordChangeDateTime

            if ([string]::IsNullOrEmpty($R.extension_b797a1cc26244252ac880ab6957387ce_extensionAttribute10) -and $this.Title -notmatch "vendor") {
                $this.AccessLevel = "FTE"
            }
            else { $this.AccessLevel = $R.extension_b797a1cc26244252ac880ab6957387ce_extensionAttribute10 }

            if ($R.DistinguishedName -match "Terminated") { $this.Active = $False }
            else { $this.Active = $True }

            if ($R.displayName -match "Admin") { $this.IsAdmin = $true }
            else { $this.IsAdmin = $false }

            ForEach ($L in $R.assignedLicenses.skuID){
                $Lic = $null
                switch ($L)
                {
                    "c5928f49-12ba-48f7-ada3-0d743a3601d5" {$Lic = "VISIO"}
                    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" {$Lic = "STREAM"}
                    "b05e124f-c7cc-45a0-a6aa-8cf78c946968" {$Lic = "EMSPREMIUM"}
                    "c7df2760-2c81-4ef7-b578-5b5392b571df" {$Lic = "ENTERPRISEPREMIUM"}
                    "87bbbc60-4754-4998-8c88-227dca264858" {$Lic = "POWERAPPS_INDIVIDUAL_USER"}
                    "f8a1db68-be16-40ed-86d5-cb42ce701560" {$Lic = "POWER_BI_PRO"}
                    "6470687e-a428-4b7a-bef2-8a291ad947c9" {$Lic = "WINDOWS_STORE"}
                    "6fd2c87f-b296-42f0-b197-1e91e994b900" {$Lic = "ENTERPRISEPACK"}
                    "99cc8282-2f74-4954-83b7-c6a9a1999067" {$Lic = "M365_E5_SUITE_COMPONENTS"}
                    "f30db892-07e9-47e9-837c-80727f46fd3d" {$Lic = "FLOW_FREE"}
                    "e43b5b99-8dfb-405f-9987-dc307f34bcbd" {$Lic = "MCOEV"}
                    "0dab259f-bf13-4952-b7f8-7db8f131b28d" {$Lic = "MCOPSTN1"}
                    "606b54a9-78d8-4298-ad8b-df6ef4481c80" {$Lic = "CCIBOTS_PRIVPREV_VIRAL"}
                    "dcb1a3ae-b33f-4487-846a-a640262fadf4" {$Lic = "POWERAPPS_VIRAL"}
                    "338148b6-1b11-4102-afb9-f92b6cdc0f8d" {$Lic = "DYN365_ENTERPRISE_P1_IW"}
                    "610b16c2-bc9b-4b6b-b59f-0168123049ad" {$Lic = "VIDEO_INTEROP"}
                    "6070a4c8-34c6-4937-8dfb-39bbc6397a60" {$Lic = "MEETING_ROOM"}
                    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" {$Lic = "POWER_BI_STANDARD"}
                    "47794cd0-f0e5-45c5-9033-2eb6b5fc84e0" {$Lic = "MCOPSNTC"}
                    "26d45bd9-adf1-46cd-a9e1-51e9a5524128" {$Lic = "ENTERPRISEPREMIUM_NOPSTNCONF"}
                    "111046dd-295b-4d6d-9724-d52ac90bd1f2" {$Lic = "WIN_DEF_ATP"}
                    "ddfae3e3-fcb2-4174-8ebd-3023cb213c8b" {$Lic = "POWERFLOW_P2"}
                    "710779e8-3d4a-4c88-adb9-386c958d1fdf" {$Lic = "TEAMS_EXPLORATORY"}
                    "90d8b3f8-712e-4f7b-aa1e-62e7ae6cbe96" {$Lic = "SMB_APPS"}
                    "fcecd1f9-a91e-488d-a918-a96cdb6ce2b0" {$Lic = "AX7_USER_TRIAL"}
                    "093e8d14-a334-43d9-93e3-30589a8b47d0" {$Lic = "RMS_BASIC"}
                    "05e9a617-0261-4cee-bb44-138d3ef5d965" {$Lic = "SPE_E3"}
                    "53818b1b-4a27-454b-8896-0dba576410e6" {$Lic = "PROJECTPROFESSIONAL"}
                    "d3b4fe1f-9992-4930-8acb-ca6ec609365e" {$Lic = "MCOPSTN2"}
                    "8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b" {$Lic = "RIGHTSMANAGEMENT_ADHOC"}
                    "18181a46-0d4e-45cd-891e-60aabd171b4e" {$Lic = "STANDARDPACK"}
                }

                $this.Licenses.Add($Lic)
            }
        }
    }
    #endregion

    Write-Verbose "Getting results from Graph"

    if ($All){
        $URL = "https://graph.microsoft.com/beta/users?`$Select=assignedLicenses,accountEnabled,displayName,onPremisesDistinguishedName,onPremisesExtensionAttributes,jobTitle,provisionedPlans,givenName,surname,lastPasswordChangeDateTime,extension_b797a1cc26244252ac880ab6957387ce_extensionAttribute2,companyName,officeLocation,extension_b797a1cc26244252ac880ab6957387ce_extensionAttribute3,department,extension_b797a1cc26244252ac880ab6957387ce_employeeType,userPrincipalName,mailNickName,showInAddressList,id,extension_b797a1cc26244252ac880ab6957387ce_extensionAttribute10,extension_b797a1cc26244252ac880ab6957387ce_description&top=999"
    }

    $Request = Get-Pages -Headers @{Authorization = "Bearer $Token"} -URL $URL

    #region Get results for passed parameters
    $Result = $Request | Where-Object {!([string]::IsNullOrEmpty($_.jobTitle)) -and $_.jobtitle -notmatch "vendor" -and $_.accountEnabled -eq 'True'}
    #endregion

    #region Check for null results
    if ([string]::IsNullOrEmpty($Result)){
        Write-Warning "Options were selected for Get-Users, but no results were returned." -WarningAction Continue
        break
        # Warning preference set to continue, but then break the script. this just looks cleaner.
    }
    #endregion

    #region Get users and put them in lists based on passed parameters
    [System.Collections.ArrayList]$DefaultList = @()
    Foreach ($R in $Result) {
            $User = [User]::new($R)

        <#
        $Licenses = $User.GetLicenses($R.assignedLicenses, $Token)
        ForEach ($L in $Licenses) {
            $User.AddLicense($L)
        }
        #>
        if ($Default) {
            Write-Verbose "Getting a list of default users"
            $DefaultList.Add($User) | Out-Null
        }
    }
    #endregion

    #region Return lists based on passed parameters
    if ($Default) {
        Write-Verbose "Returning Default user results"
        $DefaultList
    }
    #endregion
}

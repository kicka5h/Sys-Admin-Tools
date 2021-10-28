function Add-360User {
    param(
    [Parameter(Mandatory=$false)]
    [string]$Token,
    [Parameter(Mandatory=$true)]
    [string]$UPN
    )

    #build basic request info
    $url = "https://bittitan.360learning.com/api/v1/users?company=5df910d1c44dee73669ddbcd&apiKey=$Token"

    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
    }

    $body = @{
        mail = $UPN
        password = New-RandomPassword
        firstName = $FirstName
        lastname = $LastName
        job = $Title
        "groups[0]" = $GroupID1
        "groups[1]" = $GroupID2
        sendCredentials = "false"}

    Try{
        #send the request
        $response = Invoke-RestMethod -Uri $url -Method 'POST' -Headers $headers -Body $body

        #handle reactivated account
        if ($response.warning -match "reactivated"){
            Write-Warning "$UPN previously had an account in 360Learning and it has been reactivated."
        }     
    } Catch [System.Net.WebException] {
            Write-Warning "An account already exists in 360Learning for $UPN. A new account will not be created."
    }

    #region Working with managers and non-managers
    #add the user's manager 
    $URLFrag = "?company=5df910d1c44dee73669ddbcd&apiKey="
    $ManagerURL = "https://bittitan.360learning.com/api/v1/users/$UPN/managers/$Mgr$URLFrag$Token"

    try{
        $addmanager = Invoke-RestMethod -Uri $ManagerURL -Method 'PUT'
    }
    catch{
        Write-Warning "Something went wrong while adding Manager."
    }

    #add the user's direct reports
    if ($DirectReports) {
        foreach ($DR in $DirectReports) {
            try{
                $URLFrag = "?company=5df910d1c44dee73669ddbcd&apiKey="
                $ManagerURL = "https://bittitan.360learning.com/api/v1/users/$DR/managers/$UPN$URLFrag$Token"
                $addmanager = Invoke-RestMethod -Uri $ManagerURL -Method 'PUT'
            }
            catch{
                Write-Warning "Something went wrong while adding direct reports."
            }
        }
    }
}
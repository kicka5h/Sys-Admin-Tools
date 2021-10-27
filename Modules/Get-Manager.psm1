Function Get-Manager {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Userprincipalname
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

    class User {
        [string]$ObjectID
        [string]$Name
        [string]$FirstName
        [string]$LastName
        [string]$Title
        [string]$UserPrincipalName
        [string]$DistinguishedName

        # User constructor
        User ($R) {
            $this.ObjectID = $R.id
            $this.UserPrincipalName = $R.UserPrincipalName
            $this.Name = $R.displayName
            $this.FirstName = $R.GivenName
            $this.LastName = $R.Surname
            $this.Title = $R.jobTitle
            $this.DistinguishedName = $R.onPremisesDistinguishedName
        }
    }

    [System.Collections.ArrayList]$ManagerList = @()
    $URL = "https://graph.microsoft.com/beta/users/$UserPrincipalName/Manager"

    $ManRequest = @()
    Try{
        $ManRequest = Get-Pages -Headers @{Authorization = "Bearer $Token"} -URL $URL
    }
    Catch{
        Write-Verbose "This user does not have a manager."
    }

    ForEach ($M in $ManRequest) {
        $Man = [User]::new($ManRequest)

        $ManagerList.Add($Man) | Out-Null
    }

    return $ManagerList
}
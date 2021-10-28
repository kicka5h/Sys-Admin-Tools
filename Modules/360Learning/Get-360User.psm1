Function Get-360User {
    param(
        [Parameter(Mandatory=$True)]
        [string]$Token,
        [Parameter(Mandatory=$True)]
        [string]$UPN
    )#endparam

    class User {
        [string]$Name
        [string]$UPN
        [string]$ID
        [pscustomobject]$Certifications
        [pscustomobject]$Groups
        [pscustomobject]$Labels

        User($U) {
            $this.Name = $U.firstName + " " + $U.lastName
            $this.UPN = $U.Mail
            $this.ID = $U._id
            $this.Certifications = $U.certifications
            $this.Groups = $U.groups
            $this.Labels = $U.labels
        }
    }

    try {
        $URL = "https://bittitan.360learning.com/api/v1/users/$UPN`?company=5df910d1c44dee73669ddbcd&apiKey=$Token"
        $Response = Invoke-RestMethod -Uri $URL -Method GET

        $User = [user]::New($Response)
        $User
    }
    Catch [System.Net.WebException] {
        $User = $null
        Write-Warning "$UPN was not found in 360Learning."
    }
}
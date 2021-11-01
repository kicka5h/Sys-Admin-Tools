Function Get-WriteGraph {
    [CmdletBinding(DefaultParameterSetName='Application')]
    Param (
        $ClientId,
        $ClientSecret, 
        $TenantID,
        [Parameter(Mandatory=$false, ParameterSetName="Application")]
        [switch]$Application,
        [Parameter(Mandatory=$false, ParameterSetName="Delegated")]
        [switch]$Delegated,
        [Parameter(Mandatory=$false, ParameterSetName="Delegated")]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty 
    )

    if (!$Application -and !$Delegated) {
        $Application = $true
    }

    # Get an authentication token using application OAuth
    # Body params
    $scope = 'https://graph.microsoft.com/.default'

    $ErrorActionPreference = "Continue"
    if ($Application) {
        $granttype = 'client_credentials'
        $ContentType = "application/x-www-form-urlencoded"

        # Construct URI
        $uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"

        # Construct Body
        $body = @{
            client_id      = $clientId
            scope          = $scope
            client_secret  = $clientSecret
            grant_type     = $granttype
            "content-type" = $ContentType
        }

        # Get OAuth 2.0 Token
        $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -Body $body -UseBasicParsing

        # Access Token
        $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
        $token
    }

    if ($Delegated){
        $granttype = 'password'
        $ContentType = "application/x-www-form-urlencoded"

        # Construct URI
        $uri = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"

        $Password = $Credential.GetNetworkCredential().Password

        # Construct Body
        $body = @{
            client_id      = $clientId
            scope          = $scope
            client_secret  = $clientSecret
            grant_type     = $granttype
            "content-type" = $ContentType
            Username       = $Credential.UserName
            Password       = $Password
        }

        # Get OAuth 2.0 Token
        $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -Body $body -UseBasicParsing

        # Access Token
        $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
        $token
    }
}
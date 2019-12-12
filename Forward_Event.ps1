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
  [String]$ForwardSubject,

  [Parameter(Mandatory=$true)]
  [String]$ForwardUser,

  [Parameter(Mandatory=$true)]
  [String]$Note
)

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
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Get some information on the user who will receive the forwarded event
$ForwardDisplayURL = "https://graph.microsoft.com/beta/users/$ForwardUser"
$DisplayResponse = Invoke-WebRequest -Method "GET" -Uri $ForwardDisplayURL -Headers $Headers -UseBasicParsing | ConvertFrom-Json
$ForwardDisplay = $DisplayResponse.userPrincipalName

# Get the most updated event ID
$ContentType = "application/json"

$Headers = @{
    Authorization  = "Bearer $token"
    "Content-Type" = $ContentType
}

$Filter = "?$" + "Select=Subject" | Out-String

$EventURL = "https://graph.microsoft.com/beta/users/$UserName/calendar/events$Filter"

$Request = Invoke-WebRequest -Method "GET" -Uri $EventURL -Headers $Headers -UseBasicParsing | ConvertFrom-Json
$EventIDs = $Request.value.id

# Check the Event IDs for matches to the desired event subject
$CalHeader = @{
    Authorization = "Bearer $token"
}

ForEach ($Event in $EventIDs)
{
    $CalURL = "https://graph.microsoft.com/beta/users/$CalUserName/events/$Event$Filter"

    $EventRequest = Invoke-WebRequest -Method GET -Uri $CalURL -Headers $CalHeader -UseBasicParsing
    $Result = $EventRequest.content | ConvertFrom-Json
    $Subject = $Result.subject

    # Do something with the event IDs that match the subject
    If ($Subject -match $ForwardSubject)
    {
$Body = @"
        {
          "ToRecipients":[
              {
                "emailAddress": {
                  "address": "$ForwardUser",
                  "name":"$ForwardDisplay"
                }
              }
             ],
          "Comment": "$Note" 
        }
"@

        # Forward the event
        $ForwardURL = "https://graph.microsoft.com/beta/me/events/$Event/forward"

        $SendResult = Invoke-WebRequest -Method POST -Uri $ForwardURL -Headers $Headers -Body $Body -UseBasicParsing
    }
}



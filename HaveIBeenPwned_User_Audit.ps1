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
  [String]$HaveIBeenPwndAPIKey,

  [Parameter(Mandatory=$true)]
  [String]$emailTo
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
$tokenRequest = Invoke-WebRequest -Method 'POST' -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

## Null the arrays ##
[array]$breachResult = $Null
[array]$List2 = $Null

# Get the organization name for the HaveIBeenPwned headers
$Select = "$" + "Select=displayName"
$URI = "https://graph.microsoft.com/v1.0/organization?$Select"
$Head = @{
    Authorization = "Bearer $token"
}
$Org = Invoke-WebRequest -Method 'GET' -Uri $URI -Headers $Head -UseBasicParsing | ConvertFrom-Json
$OrgName = $Org.value | Select-Object -ExpandProperty DisplayName

# Build the header
$headers = @{
    "User-Agent"  = $OrgName + " Breached User Account Check"
    "hibp-api-key" = $HaveIBeenPwndAPIKey
}

# Get a list of user email addresses in the tenant
$Filter = "$" + "Select=mail"
$URL = "https://graph.microsoft.com/v1.0/users?$Filter"
$Header = @{
    Authorization = "Bearer $token"
    'Content-type' = 'application/json'
}
$User = Invoke-WebRequest -Method 'GET' -Uri $URL -Headers $Header -UseBasicParsing | ConvertFrom-Json
$UPN = $User.value | Select-Object -ExpandProperty Mail

# Find breached user accounts and build the email tables #
foreach ($U in $UPN) 
{
    $uriEncodeEmail = [uri]::EscapeDataString($U)
    $uri = "https://haveibeenpwned.com/api/v3/breachedaccount/$uriEncodeEmail"
    $breachResult = $null
        try {
            [array]$breachResult = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction SilentlyContinue
        }
        catch {
            if($error[0].Exception.response.StatusCode -match "NotFound"){
                Write-Host "No Breach detected for $U" -ForegroundColor Green
            }else{
                Write-Host "Cannot retrieve results due to rate limiting or suspect IP. You may need to try a different computer or wait 24 hours."
            }
        }
        if ($breachResult) 
        {
            $breachList = foreach ($breach in $breachResult) {
            $breach | Add-Member -MemberType NoteProperty -Name Email -Value "$U"
            $breach | Add-Member -MemberType NoteProperty -Name BreachedName -Value "$($breach.Title)"
            $breach | Add-Member -MemberType NoteProperty -Name BreachedDate -Value "$($breach.BreachDate)"

            $breach | Select-Object Email, BreachedName, BreachedDate, LastPasswordChange 
        }

            foreach ($B in $breachList) 
            {
                # Get 30 days from the Breach Date
                $BD = $breach.BreachDate
                $Today = Get-Date -format yyyy-MM-dd
                $30DA = (Get-Date).AddDays(-30).ToString('yyyy-MM-dd')

                $PassDate = Get-Date $user.LastPasswordChangeTimestamp -format yyyy-MM-dd

                # If breach date was within the last 30 days, format the email #
                if ($BD -le $Today -and $BD -gt $30DA)
                {
                    $List1 = New-Object PSObject

                    $List1 | Add-Member -MemberType NoteProperty -Name Pwned_Email -Value "$email"
                    $List1 | Add-Member -MemberType NoteProperty -Name Breach_Name -Value "$($breach.BreachedName)"
                    $List1 | Add-Member -MemberType NoteProperty -Name Breach_Date -Value "$($BD)"
                    $List1 | Add-Member -MemberType NoteProperty -Name Password_Change_Date -Value "$($PassDate)"

                    [array]$List2 += $List1
                }
            }

        }
        # API limiting: one request per IP every 1500 milliseconds. Limit violations will be blocked for 24 hours. Set scheduled task to prevent multiple instances #
        Start-sleep -Milliseconds 2000
}

## Send the email for compromised accounts ##
if($BD -le $Today -and $BD -gt $30DA)
{
    ## Set the Date ##
    $Date = Get-Date -Format 'MMMM dd yyyy'

    ## Set the CSS for the email ##
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

    ## Add originating information. (HaveIBeenPwned ULA requires credit and link to the website) ##
    $HIBP = "<a href='https://haveibeenpwned.com'>HaveIBeenPwned</a>"
    $Footer = "This report uses API and data provided by $HIBP <br>" 

    $emailList = ( $List2 | Sort-Object -unique Pwned_Email, Breach_Name, Breach_Date, Password_Change_Date | Select Pwned_Email, Breach_Name, Breach_Date, Password_Change_Date | ConvertTo-Html -Head $CSS)
 
    $Date = Get-Date -Format 'MMMM dd yyyy'

    [string][ValidateNotNullOrEmpty()]$passwd = $Password
    $secpasswd = ConvertTo-SecureString -String $passwd -AsPlainText -Force
    $cred = New-Object Management.Automation.PSCredential ($Username, $secpasswd)
 
    $From = $Username
    $Subject = "HaveIBeenPwned User Audit " + $Date
    $emailBody = "$emailList $Footer"
    $SMTPPort = "587"
    $SMTPServer = "smtp.office365.com"
 
    Send-MailMessage -From $From -to $emailTo -Subject $Subject -Body $emailBody -BodyAsHTML -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $cred
}

## If no compromised accounts are found ##
else 
{
    $Date = Get-Date -Format 'MMMM dd yyyy'

    [string][ValidateNotNullOrEmpty()]$passwd = $Password
    $secpasswd = ConvertTo-SecureString -String $passwd -AsPlainText -Force
    $cred = New-Object Management.Automation.PSCredential ($Username, $secpasswd)
 
    $From = $Username
    $Subject = "HaveIBeenPwned User Audit " + $Date
    $emailBody = "Good news! No user accounts have been breached in the last 30 days." + $Footer
    $SMTPPort = "587"
    $SMTPServer = "smtp.office365.com"
 
    Send-MailMessage -From $From -to $emailTo -Subject $Subject -Body $emailBody -BodyAsHTML -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $cred
}

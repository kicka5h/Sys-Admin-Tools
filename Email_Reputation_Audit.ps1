#Get the required params
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [String]$Username,
  
    [Parameter(Mandatory=$true)]
    [String]$Password,

    [Parameter(Mandatory=$true)]
    [String]$To
)

# Set up the CSS for the email 
#$CSS = @"
#<Style>
#    table { margin: auto; font-family: Brandon Grotesque; border-collapse: collapse; }
#    table, th, td { border: 1px solid #E3E3E3; }
#    th { background: #AFC2C4; color: #fff; max-width: 400px; padding: 5px 10px; }
#    td { font-size: 11px; padding: 5px 20px; color: #000; }
#    tr { background: #fff; }
#<⁄style>
#"@

#[string][ValidateNotNullOrEmpty()]$passwd = $Password
#$secpasswd = ConvertTo-SecureString -String $passwd -AsPlainText -Force
#$cred = New-Object Management.Automation.PSCredential ($Username, $secpasswd)
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic –AllowRedirection
#Import-PSSession $Session
#Connect-MsolService -credential $cred

# Optional code for getting the email list from Quarantine
#$SenderList = Get-QuarantineMessage | Select-Object -ExpandProperty SenderAdrress
#$List = $SenderList | Sort-Object -Property @{Expression={$_.Trim()}} -Unique

$List = "geschient@orgtracks.com", "nooreply@c14.lpwwtdyrclvtqcq.com", "support@thehappily.co", "nytdirect@nytimes.com"

$Header = @{
    Accept = "application/json"
    "Content-Type" = "application/x-www-form-urlencoded"
    Key = "qb5ek8ngex0i0zqwwuiumoueeis7rgalc4azvxp50l8l5ykn"
}

[array]$table = $null
foreach ($Email in $List){
    $Query = "?summary=true"
    $URI = "https://emailrep.io/$Email$Query"

    $Response = Invoke-WebRequest -Method GET -Uri $URI -Headers $Header -UseBasicParsing | ConvertFrom-Json
    
    $Reputation = $Response.reputation
    $Suspicious = $Response.suspicious
    $Malicious  = $Response.details.malicious_activity_recent
    $Spam       = $Response.details.spam
    $Summary    = $Response.summary

    [array]$table += [PSCustomObject]@{
        Email      = $Email
        Reputation = $Reputation
        Suspicious = $Suspicious
        Malicious  = $Malicious
        Spam       = $Spam 
        Summary    = $Summary
    }
}

# Prepare the body of the email. Use Fragment to keep it simple.
$Report = ($Table | Sort-Object Email, Reputation, Suspicious, Malicious, Spam, Summary | Select-Object Email, Reputation, Suspicious, Malicious, Spam, Summary | ConvertTo-Html -Fragment )

$Body = @"
$CSS
$Report
"@

# Get credentials to send the email
[string][ValidateNotNullOrEmpty()]$passwd = $Password
$secpasswd = ConvertTo-SecureString -String $passwd -AsPlainText -Force
$cred = New-Object Management.Automation.PSCredential ($Username, $secpasswd)

# Set up the email params
$Date       = Get-Date -Format 'MMMM dd yyyy'
$From       = $Username
$Subject    = "Email Reputation Audit " + $Date
$SMTPPort   = "587"
$SMTPServer = "smtp.office365.com"

# Send the email
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $cred
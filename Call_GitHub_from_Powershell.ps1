#Get the required params
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [String]$Token,

  [Parameter(Mandatory=$true)]
  [String]$Username,

  [Parameter(Mandatory=$true)]
  [String]$Repo,

  [Parameter(Mandatory=$true)]
  [String]$Filename
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$URI = "https://api.github.com/repos/$Username/$Repo/contents/$FileName"

$Headers = @{
    accept = "application/vnd.github.v3.raw"
    authorization = "Token " + $Token
}

Try 
{
    Write-Host "Starting powershell script $Filename from $Repo..." -ForegroundColor Yellow
    $Script = Invoke-RestMethod -Uri $URI -Headers $Headers
    Invoke-Expression $Script
} 
catch [System.Net.WebException] 
{
    Write-Host "Error connecting to $Filename. Please check your file name or repo name and try again." -ForegroundColor Red
}

Write-Host "Success. Now executing $Filename from $Repo." -ForegroundColor Green
Param(
    $servicePrincipalConnection
)

$connectionName = $servicePrincipalConnection
$servicePrincipalConnection = Get-AutomationConnection -Name $connectionName

Write-Output "Logging in to Exchange..."
Connect-ExchangeOnline -AppId $servicePrincipalConnection.ApplicationID `
                       -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint `
                       -Organization "migrationwiz.onmicrosoft.com"

class Booking {
    [string]$Name
    [string]$Address
    [string]$Owner
    [string]$Site
    [string]$SiteMail
    [bool]$Hidden
    [bool]$Active
}

$Booking = get-mailbox -RecipientTypeDetails Scheduling

[array]$BookingList = $null
ForEach ($B in $Booking) {
    $Error[0] = $null
    $BSite  = [Booking]::New()

    $Address = $B.UserPrincipalName
    $Name = $B.DisplayName
    $Owner = ($B.ForwardingSmtpAddress).Split(':')
    $Site = "https://outlook.office365.com/owa/calendar/$Address/bookings/"
    $SiteMail = $B.WindowsEmailAddress
    $Hidden = $B.HiddenFromAddressListsEnabled

    $TestSite = Invoke-WebRequest $Site

    if ($Error[0].Exception -match (400)) { $BSite.Active = $False }
    else { $BSite.Active = $True }

    $BSite.Name = $Name
    $BSite.Address = $Address
    $BSite.Site = $Site
    $BSite.SiteMail = $SiteMail
    $BSite.Owner = $Owner[1]
    $BSite.Hidden = $Hidden

    [array]$BookingList += $BSite
}
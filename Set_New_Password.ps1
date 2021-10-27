param(
    $UserPrincipalName,
    $Credential
)

function Get-RandomO365Pass() {

    function Get-RandomCharacters($length, $characters) { 
        $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
        $private:ofs="" 
        return [String]$characters[$random]
    }

    function Scramble-String([string]$inputString){     
        $characterArray = $inputString.ToCharArray()   
        $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
        $outputString = -join $scrambledStringArray
        return $outputString 
    }

    # List of characters accepted by O365, minus ones that may be misread
    $pwChars = @('A','B','C','D','E','F','G','H','J','K','L','M','N','P','Q','R','S','T','U','V','W','X','Y','Z',`
                'a','b','c','d','e','f','g','h','i','j','k','m','n','o','p','q','r','s','t','u','v','w','x','y','z',`
                '1','2','3','4','5','6','7','8','9','~','!','@','#','$','%','^','*','(',')','+','=','-','_','{','}','[',']','\',':',';','?','/')

    # Instantiating list for the password
    [System.Collections.Generic.List[int]]$toAsciiValue = @()

    foreach ($Char in $pwChars) {
        $ascii = [int][char]$Char
        $toAsciiValue.Add($ascii)
    }

    $password = ($toAsciiValue | Get-Random -count 6 | ForEach-Object { [char]$_ }) -Join ""
    $password += Get-RandomCharacters -length 1 -characters 'abcdefghikmnopqrstuvwxyz'
    $password += Get-RandomCharacters -length 1 -characters 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $password += Get-RandomCharacters -length 1 -characters '123456789'
    $password += Get-RandomCharacters -length 1 -characters '~!@#$%^*()+=_-{}[]\:;?/)'

    $password = Scramble-String $password
  
    return $password
}

function Set-NewUserPass ($Token, $UPN, $Password) {
    $FindIDURL = "https://graph.microsoft.com/beta/users/$UPN/authentication/passwordMethods"
    $ID = (Invoke-RestMethod $FindIDURL -Method GET -Headers @{Authorization = "Bearer $Token"}).value.id

    $NewPassBody = "{
    `n  `"newPassword`": `"$Password`"
    `n}"

    $NewPassHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $NewPassHeaders.Add("Authorization", "Bearer $Token")
    $NewPassHeaders.Add("Content-Type", "application/json")
    
    $ChangePassURL = "https://graph.microsoft.com/beta/users/$UPN/authentication/passwordMethods/$ID/resetPassword"
    $ResetPassURL = Invoke-WebRequest $ChangePassURL -Method POST -Headers $NewPassHeaders -Body $NewPassBody

    $ProgressHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $ProgressHeaders.Add("Authorization", "Bearer $Token")
    $ProgressHeaders.Add("Content-Type", "application/json")

    $ProgressURL = $ResetPassURL.Headers.Location

    Do {
        $ProgressCheckURL = Invoke-RestMethod $ProgressURL -Method GET -Headers $ProgressHeaders
        $CheckStatus = $ProgressCheckURL.status
        $StatusDetail = $ProgressCheckURL.statusDetail
    }
    While ($checkStatus -eq "running")

    $NewUserPass = [pscustomobject]@{
        "Username" = $UPN
        "Status" = $CheckStatus
        "Details" = $StatusDetail
        "New Password" = $NewPass
    }

    $NewUserPass
}

$graphToken = Get-GraphToken -Delegated -Credential $Credential

$findUser = Get-Users -All | Where-Object {$_.userprincipalName -match $UserPrincipalName}

# Check to see if the user exists, generate O365 compliant password, create secret, set O365 password
if ($findUser) {
    $password = Get-RandomO365Pass
    Set-NewUserPass -Token $graphToken -UPN $findUser.userprincipalName -Password $password
}
else {
    Write-Ouput "$UserPrincipalName does not exist. Please check your spelling and try again."
}
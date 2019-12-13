<#
.SYNOPSIS
  <Intended to aid in the population of an O365 test environment by creating users and Teams. Recommended to run in a tenant with no users or Teams. Not recommened for prod tenants.>
.DESCRIPTION
  <This script creates 2 users, 10 Teams, 5 Channels in those Teams, posts 5 messages to each channel, and uploads 5 txt files to each channel site>
.INPUTS
  <Global admin user credentials, user license assignment, Tenant ID/ClientID/SecretID for Azure Native App for Graph>
.OUTPUTS
  <None>
.NOTES
  Version:        1.0
  Author:         <Ash Karczag>
  Co-Author:      <Antonio Vargas>
  Creation Date:  <Dec 2019>
  Purpose/Change: Initial script development
#>

#Get the required params
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [String]$AdminUser,

  [Parameter(Mandatory=$true)]
  [String]$AdminPass,

  [Parameter(Mandatory=$true)]
  [String]$License,

  [Parameter(Mandatory=$true)]
  [String]$tenantId,

  [Parameter(Mandatory=$true)]
  [String]$clientId,

  [Parameter(Mandatory=$true)]
  [String]$ClientSecret
)

# Set the web request protocol
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Connect to services
function MSOLConnected {
    Get-MsolDomain -ErrorAction SilentlyContinue | out-null
    $result = $?
    return $result
}

if (-not (MSOLConnected)) {
    #  Connect to O365
    Write-Host "Connecting to O365 and Microsoft Teams..." -ForegroundColor Gray
    Try{
        Set-ExecutionPolicy RemoteSigned -ErrorAction Stop
    }
    catch{
        write-Host "Failed to set the execution policy to run this script. Please open the PowerShell as Admin and try again." -ForegroundColor Red
        Exit
    }
    Set-ExecutionPolicy RemoteSigned
    $AdminPass1 = ConvertTo-SecureString $AdminPass -AsPlainText -Force;
    $cred = new-object -typename system.management.automation.pscredential -argumentlist $AdminUser, $AdminPass1
    Connect-MsolService -credential $cred
    Connect-MicrosoftTeams -credential $cred
}

# Set the domain value
$OnmicrosoftDomain = Get-MsolDomain | ? {$_.Name -like "*onmicrosoft.com"}
$YourDomain = "@" + ($OnmicrosoftDomain | Select-object -ExpandProperty Name)

# CREATE AND LICENSE USER ACCOUNTS
# Grab random username values
Write-Host "Creating new users in your tenant..." -ForegroundColor Gray
$OptVal = "sandbox gregory","register cricked","module temporary","slumbering little","unarmored bradford","granny descent","drove now","title health","suspect diving","stunt skylight","wind rope","aidless heaviness","barkisin radius","exponent hybrid","bello scope","supermom bomb","imaging sophronia","woolwich wilfer","milk urban","wood empanada","groom dragley","disorder cod","belling request","shrine nextdoor"
$OptValFinal = $OptVal | Get-Random -Count 2
[array]$UserName = $null
[array]$UserName = $OptValFinal

# Create Names and Usernames for each of the values entered
ForEach ($User in $UserName){
    $UPN1 = $UserName -replace '"',''
}

$Userlist = $UPN1 | ForEach-Object { $U = "" | Select-Object First, Last, Full, UPN
    $U.first = $_ -split " " | Select-Object -first 1
    $U.last = $_ -split " " | Select-Object -last 1
    $U.full = $_
    $U.upn = $U.First[0] + $U.Last + $YourDomain
    $U}

# Create the user accounts
#Array for users
$UsersCreated = @()
ForEach ($NewUser in $Userlist){
    #Check if user exists
    $UserExists = get-msoluser -UserPrincipalName $NewUser.upn -ErrorAction SilentlyContinue
    if($UserExists){
        Write-Host "User $($NewUser.upn) already exists. Bypassing creation and user won't be added to channels during this script execution" -ForegroundColor Yellow
        break
    }
    $UPN = $NewUser.UPN
    Write-Host "Creating user $UPN..." -ForegroundColor Green
    $UsersCreated += New-MSolUser -DisplayName $newUser.full -FirstName $newuser.first -LastName $newuser.last -UserPrincipalName $newuser.upn -UsageLocation US -LicenseAssignment $License
}
                                                                     
# CREATE MICROSOFT TEAMS
# Grab Random Team Names
$OptVal2 = "stool attic", "hobo political", "dish fast", "indignant secluded", "swivel thaging","yawler chad","snow duration","found wager","gumps doctor","shadow commotion","drier dumplings","overfill intrepid","starch somewhere","groffee pipe", "halls each","baby viewer","candles cabby","benny duke","ladle breed","grus symbols","dartmoor broke","missy siren","thirteen clever","legun lovable","dindow shinking","pucker sheep","front mustang","globe rightful","dizziness scratchy","orogeny bumped","ecotone sizably","destiny forget"
#Count how many teams you need
$NumTeams = (get-team).Count

if ($NumTeams -lt 10){
    $NumTeamsCreate = 10 - $NumTeams
}
else{
    $NumTeamsCreate = 0
    write-host "Bypassing the creation of Teams, since you already have 10 or more" -ForegroundColor Yellow
}

# Grab Random Team Names
Write-Host "Creating new Teams in your tenant..." -ForegroundColor Gray
$OptVal2 = "stool attic", "hobo political", "dish fast", "indignant secluded", "swivel thaging","yawler chad","snow duration","found wager","gumps doctor","shadow commotion","drier dumplings","overfill intrepid","starch somewhere","groffee pipe", "halls each","baby viewer","candles cabby","benny duke","ladle breed","grus symbols","dartmoor broke","missy siren","thirteen clever","legun lovable","dindow shinking","pucker sheep","front mustang","globe rightful","dizziness scratchy","orogeny bumped","ecotone sizably","destiny forget"

if ($NumTeamsCreate -ne 0){
    $TeamsCreated = 0
    #loop to create the teams we need
    DO{
        $TeamsCreated++
        $NameAttempts = 0
        #loop to get a proper team name
        DO{
            $NameAttempts++
            #get the teams Name
            $OptFinal2 = $OptVal2 | Get-Random -Count 1
            # Create the Teams DisplayNames and UPNs
            $Group = $Optfinal2 -replace '"',''
            $G = $Group
            $FirstName = $G -split " " | Select-Object -first 1
            $LastName = $G -split " " | Select-Object -last 1
            $DisplayName = $G
            $UPN = $FirstName[0] + $LastName
            #Check if team exists
            $TeamExists = get-team -MailNickname $UPN
            If (!($TeamExists)){
                break
            }
        } Until ($NameAttempts -eq 10)
        #Create the team
        Write-Host "Creating a Team called $DisplayName..." -ForegroundColor Green
        New-Team -DisplayName $DisplayName -MailNickname $UPN | out-null
    } Until ($TeamsCreated -eq $NumTeamsCreate)
}

#Grab 10 Teams

$TeamList = Get-Team | Select -first 10

#AddUsers to Teams
Write-Host "Adding users to the Teams in your tenant..." -ForegroundColor Gray
Foreach ($Team in $TeamList){
    if ($UsersCreated){
        Write-Host "Adding Users to Team $($Team.Displayname)..." -ForegroundColor DarkGreen
        foreach ($UserCreated in $UsersCreated){
            #Check if the user exists in the team
            $TeamMembers = Get-TeamUser -GroupID $Team.GroupID
            $Exists = $TeamMembers.User -Contains $UserCreated.UserPrincipalName
            If ($Exists -eq $False){
                Write-Host "Adding $($UserCreated.UserPrincipalName) to Team $($Team.DisplayName).." -ForegroundColor Green
                Add-TeamUser -groupid $Team.GroupID -user $UserCreated.UserPrincipalName | out-null 
            }
            Else{
                Write-Host "User $($UserCreated.UserPrincipalName) is already a member of Team $($Team.DisplayName).." -ForegroundColor yellow
            }
        }
    }
    Else{
        Write-Host "No users were created and need to be added to Teams" -ForegroundColor Yellow
    }
}

# Create Channels in each of the Teams
$OptVal3 = "upload tin", "jet dond", "miss charitable", "annoyed each", "okay malarkey", "saddlebag chortle", "primary shopping", "emit driftwood", "beer protozoan", "path hombre",     "sulphate farrow", "merlin flandy", "creepy unlovely", "lewis establishment", "tuxedo debit", "yoda rafter", "fox pasta", "holder sing", "red maroon", "nuclear unguarded", "antidote consumer", "tiptop prada", "scheme racehorse", "herb peanut", "greater engineer", "stinking subaltern", "programmes spinal", "turnover fuel", "rebalance stifling", "blue peterbook", "almanac opium", "casting incisive", "advise sardine", "necan compromise", "sneaky micker"

# Start new channels loop (unfinished)
ForEach ($B in $Teamlist){
    $NumChannels = (Get-TeamChannel -GroupId $B.GroupID).count

    If($NumChannels -lt 5){$NumChannelsCreate = 5 - $NumChannels}
    Else{
        $NumChannelsCreate = 0
        write-host "Bypassing the creation of channels for Team $($B.DisplayName), since you already have 5 or more" -ForegroundColor Yellow
    }

    Write-host "Creating $($NumChannelsCreate) channels in Team " $B.DisplayName -ForegroundColor Green
       
    if($NumChannelsCreate -ne 0){
        $ChannelsCreated = 0
        #Loop to create the channels we need
        Do{
            $ChannelsCreated++
            $NamesAttempted = 0 
            #loop to get a proper channel name
            Do{
                $NamesAttempted++
                $OptFinal3 = $OptVal3 | Get-Random -Count 1
                #Check if channel exists
                $ChannelExists = Get-TeamChannel -GroupID $B.GroupID | Where-Object {$_.DisplayName -eq $OptFinal3}
                If (!($ChannelExists)){break}
            } Until ($NamesAttempted -eq 5)
            #Create the Channels
            Write-Host "Creating channel $OptFinal3..." -ForegroundColor DarkGreen
            New-TeamChannel -GroupID $B.GroupID -DisplayName $OptFinal3 | Out-Null
        } Until ($ChannelsCreated -eq $NumChannelsCreate) 
    }
}
# End new channels block (unfinished)

# POST MESSAGES TO THE TEAMS CHANNELS
Write-Host "Creating usable data in your Teams..." -ForegroundColor Gray
$OptVal4 = "The lake is a long way from here.", "If I don't like something, I'll stay away from it.", "Writing a list of random sentences is harder than I initially thought it would be.", "I want more detailed information.", "Abstraction is often one floor above you.", "We need to rent a room for our party.", "I am counting my calories, yet I really want dessert.", "Is the stitch trade better than the plastic?", "What if the dizzy employment ate the chest?", "Is the scold serve better than the practice?", "The lumpy attempt can't clip the pool.", "Did the burly opposite really record the cost?", "It was then the gaudy development met the mediocre secret.", "The dimwitted definition gazes into the key church.", "Did the awkward zone really record the travel?", "Did the burly opposite really record the cost?", "It was then the gaudy development met the mediocre secret.", "The dimwitted definition gazes into the key church.", "Did the awkward zone really record the travel?"

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
    username      = $AdminUser
    password      = $AdminPass
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# Message loop
ForEach ($L in $Teamlist)
{
    $TID = $L.GroupID
    $TeamName = $L.DisplayName

    try{
        $CID = Get-TeamChannel -GroupID $TID -ErrorAction Stop| Select-Object -ExpandProperty ID
        Write-Host "Posting messages to channels in Team $TeamName..." -ForegroundColor Green
        ForEach ($C in $CID)
        {
            # Get a random message to post
            $ChannelChats = $OptVal4 | Get-Random -Count 5

            ForEach ($Chat in $ChannelChats){
            # Load the JSON payload
                $Body = @{
                    'body' =
                        @{
                            'content' = $Chat
                        }
                    } | convertto-json

                # Build the URL
                $URL = "https://graph.microsoft.com/beta/teams/$TID/channels/$C/messages"

                # Send the channel chat post request
                $Request = Invoke-RestMethod -Uri $URL -Method "Post" -Header @{Authorization = "Bearer $token"} -Body $Body -ContentType "application/json"
            }
        }

    }
    Catch{
        Write-Host "Skipping messages to channels in Team $TeamName due to error: $Error[0].ErrorDetails.Message" -ForegroundColor Yellow
    } 
}

# UPLOAD FILES TO TEAMS SITE
# Get some random file names
$OptValue5 = "spring", "sassy", "correspond", "stitch", "dad", "tear", "yell", "learn", "seem", "mellow", "inaugurate", "read", "cable", "select", "mass"
$fileType = ".txt"

# Build the file list
[array]$FileNames = $null
[array]$fileNames = foreach ($O in $OptValue5){
    $O + $fileType
}

$FileHead = @{
    Authorization = "Bearer $token"
    "Content-Type" = "Text/Plain"
    Size = 11370
}

# A bunch of digits from Pi
$Body = "314159265358979323846264338327950288419716939937510582097494459230781640628620899862803482534211706798214808651328230664709384460955058223172535940812848111745028410270193852110555964462294895493038196442881097566593344612847564823378678316527120190914564856692346034861045432664821339360726024914127372458700660631558817488152092096282925409171536436789259036001133053054882046652138414695194151160943305727036575959195309218611738193261179310511854807446237996274956735188575272489122793818301194912983367336244065664308602139494639522473719070217986094370277053921717629317675238467481846766940513200056812714526356082778577134275778960917363717872146844090122495343014654958537105079227968925892354201995611212902196086403441815981362977477130996051870721134999999837297804995105973173281609631859502445945534690830264252230825334468503526193118817101000313783875288658753320838142061717766914730359825349042875546873115956286388235378759375195778185778053217122680661300192787661119590921642019893809525720106548586327886593615338182796823030195203530185296899577362259941389124972177528347913151557485724245415069595082953311686172785588907509838175463746493931925506040092770167113900984882401285836160356370766010471018194295559619894676783744944825537977472684710404753464620804668425906949129331367702898915210475216205696602405803815019351125338243003558764024749647326391419927260426992279678235478163600934172164121992458631503028618297455570674983850549458858692699569092721079750930295532116534498720275596023648066549911988183479775356636980742654252786255181841757467289097777279380008164706001614524919217321721477235014144197356854816136115735255213347574184946843852332390739414333454776241686251898356948556209921922218427255025425688767179049460165346680498862723279178608578438382796797668145410095388378636095068006422512520511739298489608412848862694560424196528502221066118630674427862203919494504712371378696095636437191728746776465757396241389086583264599581339047802759009946576407895126946839835259570982582262052248940772671947826848260147699090264013639443745530506820349625245174939965143142980919065925093722169646151570985838741059788595977297549893016175392846813826868386894277415599185592524595395943104997252468084598727364469584865383673622262609912460805124388439045124413654976278079771569143599770012961608944169486855584840635342207222582848864815845602850601684273945226746767889525213852254995466672782398645659611635488116170313586767106436587660551655133113317022718232156877362195848216856465284606970661905439540140651063097333651381196333165949030392164270853542280497980267149118956364251748913441214263615"

# File loop
ForEach ($M in $Teamlist)
{
    $TeamID = $M.GroupID
    $TNameSite = $M.DisplayName
    try{
        $Sitename = Get-TeamChannel -GroupID $TeamID -ErrorAction Stop | Select-Object -ExpandProperty DisplayName
        Write-Host "Uploading files to channel sites in Team $TNameSite..." -ForegroundColor Green
        ForEach ($S in $Sitename)
        {
            # Get a random files to upload
            $ChannelFile = $FileNames | Get-Random -Count 5
    
            ForEach ($CF in $ChannelFile){
                $CF = $CF + ':'
    
                # Build the URL
                $Endpoint = "https://graph.microsoft.com/beta/groups/$TeamID/drive/items/root:/$S/$CF/Content"
    
                # Send the channel chat post request
                $Request = Invoke-RestMethod -Uri $Endpoint -Method "PUT" -Header $FileHead -Body $Body -ContentType "application/json"
            }
        }
    }
    Catch{
        Write-Host "Skipping file upload to channel sites in Team $TNameSite due to error: $Error[0].ErrorDetails.Message" -ForegroundColor Yellow

    }
}

Write-Host "Congratulations! Tenant configuration is complete."

Exit

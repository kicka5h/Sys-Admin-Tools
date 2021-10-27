param(
    $CalendarOwner,
    $UserList
    #list of user objects from Get-Users that will get the events forwarded
)

Write-Output "Getting a token for graph"
$graphtoken = Get-GraphToken -Application

class Event {
    [string]$id
    [string]$subject
    [string]$body
    $attendees

    Event ($E) {
        $this.id = $E.id
        $this.subject = $E.subject
        $this.body = $E.body
        $this.attendees = $E.attendees
    }
}

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


$url = "https://graph.microsoft.com/beta/users/$CalendarOwner/events"
$headers = @{
    Authorization = "Bearer $GraphToken"
    "Content-Type" = "application/json"
}

Write-Output "Getting calendar details from calendar owner"
$EventList = Get-Pages -Headers $headers -url $url

ForEach ($User in $UserList) {
    $Startdate = Get-Date ($User.startdate)
    $UPN = $user.userprincipalname
    $Name = $User.name
    $Type = $User.EmployeeType

    Foreach ($E in $EventList) {
            $eventname = $E.subject
            $id = $E.id
            $body = 
@"
    { 
        "ToRecipients":[
              {
                "emailAddress": {
                  "address":"$UPN",
                  "name":"$Name"
                }
              }
             ],
         "Comment": "This meeting was forwarded automatically." 
    }
"@
            $frag = "/$id/forward"

            $ForwardURL = $URL + $Frag

            try {
                $SendResult = Invoke-WebRequest -Method POST -Uri $ForwardURL -Headers $Headers -Body $Body -UseBasicParsing
                Write-Output "Forwarding $ForwardURL to $UPN"
            }
            catch {
                Write-Warning "$Eventname was not forwarded to $Name"
            }
    }
}
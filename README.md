# Sys-Admin-Tools
Sys-Admin-Tools is a collection of powershell tools that make a sysadmin's life just a little easier. 
<br>

##  Populate_Teams_Data.ps1
Using a combination of O365 Powershell and Graph API. Intended to aid in the population of an O365 test environment by creating users and Teams. Recommended to run in a tenant with no users or Teams. Can be run multiple times. Not recommened for prod tenants. 
<br>

#### Prerequisites
<ul>
    <li> A sandbox or test O365 tenant to run the script
    <li> An Azure native app to access Graph API
</ul>

## HaveIBeenPwned_User_Audit.ps1
Uses a combination of Graph API and HaveIBeenPwned API. Audits O365 user accounts for breaches present in the HaveIBeenPwned database, and sends an email if there was a breach in the last 30 days.
<br>

#### Prerequisites
<ul>
    <li> A subscription from HaveIBeenPwned for $3.50/ month to access the API
</ul>

## O365_License_Consumption_Audit(Graph).ps1
Leveraging both Graph API and O365 Powershell, this script will alert you when your license availability is below 2 for purchased licenses, and when users are not licensed with a defined set of required licenses. 
<br>

## Forward_Event.ps1
This script leverages Graph API to grab an event object from a source calendar and forward it to a target calendar based on a defined subject string.
<br>

#### Prerequisites
<ul>
    <li> Azure Native App with the following delegated API permissions for Graph: User.Read, User.ReadWrite, User.ReadBasic.All, User.Read.All, User.ReadWrite.All, Directory.Read.All, Directory.ReadWrite.All, Directory.AccessAsUser.All 
    <li> The executing user account needs full access delegation to the target user calendar
    <li> Both accounts need to be a member of an Exchange admin role that allows Application Impersonation
</ul>

## Send_Email.ps1
Use HTML/CSS and Powershell to style and send attractive emails.
<br>

## Email_Reputation_Audit.ps1
Use the free emailrep.io tool to check the online reputation of a list of email accounts pulled from O365 quarantine.
<br>
#### Prerequisites
<ul>
    <li> A free API key from emailrep.io
</ul>

## Call_GitHub_from_Powershell.ps1
This script uses a GitHub personal access token passed in an HTTP call to GitHub API to locally run a script from a private GitHub repo.
<br>

#### Prerequisites
<ul>
    <li> Create a user access token on your GitHub account
</ul>

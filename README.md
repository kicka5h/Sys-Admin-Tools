# Sys-Admin-Tools
Sys Admin Tools is a collection of powershell tools that make a sys admin's life just a little easier. 
<br>

### Populate_Teams_Data.ps1
Using a combination of O365 Powershell and Graph API. Intended to aid in the population of an O365 test environment by creating users and Teams. Recommended to run in a tenant with no users or Teams. Not recommened for prod tenants.
<br>

### HaveIBeenPwned_User_Audit.ps1
Uses a combination of Graph API and HaveIBeenPwned API. Requires a month or monthly subscription from HaveIBeenPwned. Audits O365 user accounts for breaches present in the HaveIBeenPwned database.
<br>

### Forward_Event.ps1
This script leverages Graph API to grab an event object from a source calendar and forward it to a target calendar.
<br>

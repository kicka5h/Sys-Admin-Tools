$Chrome = (Get-Item (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe').'(Default)').VersionInfo

if($Chrome){
    [System.Diagnostics.Process]::Start("chrome.exe", "https://www.github.com/kicka5h/sys-admin-tools" ) 
    [System.Diagnostics.Process]::Start("chrome.exe", "https://www.kicka5h.io" ) 
}
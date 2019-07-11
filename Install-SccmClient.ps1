$SCCMServerName = 
$SiteCode = 
$MPServer = 
If (Test-Path c:\Software\SCCM-Client) {
Echo 'Path already exists.'
}
Else { New-Item -ItemType directory -Path c:\Software\SCCM-Client}

Copy-Item -path \\$SCCMServerName\SMS_CHI\Client\* -Destination c:\Software\SCCM-Client

cd c:\Software\SCCM-Client

.\ccmsetup.exe /mp:$MPServer SMSSITECODE=$SiteCode /skipprereq:windowsupdateagent30-x64.exe
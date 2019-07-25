#Checks connection to SCCM Server (). If unavilable able an entry is added to the host file.

<# Creates host entry for 
IF (Test-Connection -ComputerName ){
$HostAccessible = "True"
Write-Host SCCM Server:  is accessible. 
} Else{
$HostAccessible = "False"
Write-Host SCCM Server:  is NOT accessible. Updating host file....
If ((Get-Content "$($env:windir)\system32\Drivers\etc\hosts" ) -notcontains "*IP/SERVERNAME*")  
 {ac -Encoding UTF8  "$($env:windir)\system32\Drivers\etc\hosts" "*IP/SERVERNAME*" }
 Start-sleep -Seconds 5
}
#>


If (Test-Path c:\Software\SCCM-Client) {
Echo 'Path already exists.'
Copy-Item -path \\Sharelocation\public\Microsoft\SCCM\Clients\Current\Client\* -Destination c:\Software\SCCM-Client -Force -Recurse
}
Else { New-Item -ItemType directory -Path c:\Software\SCCM-Client
Copy-Item -path \\sharelocation\public\Microsoft\SCCM\Clients\Current\Client\* -Destination c:\Software\SCCM-Client -Force -Recurse
}


#Command to install SCCM Client
IF(Test-Path C:\Software\SCCM-Client\ccmsetup.exe){
C:\Software\SCCM-Client\ccmsetup.exe /mp: SMSSITECODE= SMSMP= /skipprereq:windowsupdateagent30-x64.exe /skipprereq:silverlight.exe
}Else{
Echo "C:\Software\SCCM-Client\ccmsetup.exe does not exist"
start-sleep -Seconds 10
exit
}


#Pause to allow log to be created
Start-Sleep -Seconds 20

#Vaiable for log location
$SuccessfullInstallString = Select-String -Path C:\Windows\ccmsetup\Logs\ccmsetup.log -Pattern "CcmSetup is exiting with return code 0"
$Counter =  0
$Limit = 10
Do {
Start-Sleep -Seconds 30
$Counter++
 }
 Until (($SuccessfullInstallString -ne $null) -or ($Counter -gt $Limit))

#Machine Policy Retrieval Cycle
Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}"

#Wait 30 Seconds
Start-Sleep -Seconds 30

#Machine Policy Evaluation Cycle
Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}"


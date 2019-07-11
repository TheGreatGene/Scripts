###################################################################################
# This script users monitorinfoview to gather monitor EDID data, the data is stored in a csv file, 
# powershell extracts the serial and model # from the csv, and writes the data to the registry
#
#
# Created by: Gene Shelby 
#LAST UPDATED: 11/17/2016
###################################################################################
#Stores computer name in a vaiable
$ComputerName = [Environment]::MachineName

#Stores computer model in a vairable
$ComputerModelNumber = Get-WmiObject -Class Win32_ComputerSystem | fl Model

#Stores computer Serial # in a vairable
$ComputerSerialNumber = gwmi win32_bios | fl SerialNumber


#Store last last user to login in variable
$LastUserToLogon = (Get-WmiObject -Class Win32_ComputerSystem).UserName

#Store date/time in variable
$TimeStamp = get-date -format f

#Checks to see if active.monitors.csv file exists
IF (Test-Path C:\Software\Tools\monitorinfoview\attached.monitors.csv) { 
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - attachedmonitors.monitors.csv file exists"
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Proceeding with script"
} Else {
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - attachedmonitors.monitors.csv file does not exist"
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Exiting script"
Exit
}


#Imports the .CSV file and stores data in $ActiveMonitors Variables
IF (Test-Path C:\Software\Tools\monitorinfoview\attached.monitors.csv) { 
$ActiveMonitors = Import-csv "C:\Software\Tools\monitorinfoview\attached.monitors.csv" -Header MonitorName, Active, SerialNumber, ManufactureWeek, ManufacturerID, ProductID, MaximumResolution, ImageSize, MaximumImageSize, HorizontalFrequency, VerticalFrequency, Digital, StandbyMode, SuspendMode, Low-PowerMode, DefaultGTF, DisplayGamma, SerialNumberNumeric, EDIDVersion, RegistryKey, ComputerName, LastUpdateTime
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - CSV Imported"
}


#Monitor1 Variables
$Monitor1Model = $ActiveMonitors[0].MonitorName
$Monitor1Serial = $ActiveMonitors[0].SerialNumber
$Monitor1NumericSerial = $ActiveMonitors[0].SerialNumberNumeric

#Monitor 2 Variables
$Monitor2Model = $ActiveMonitors[1].MonitorName
$Monitor2Serial = $ActiveMonitors[1].SerialNumber
$Monitor2NumericSerial = $ActiveMonitors[1].SerialNumberNumeric

#Monitor 3 Variables
$Monitor3Model = $ActiveMonitors[2].MonitorName
$Monitor3Serial = $ActiveMonitors[2].SerialNumber
$Monitor3NumericSerial = $ActiveMonitors[2].SerialNumberNumeric

#Determines Serial# prefix based on Acer Model #
IF ($Monitor1Model -eq "ACER V243H" -or "ACER V240HL" -or "ACER S240HL") {
$AcerModelPrefix = "ET"
}
IF ($Monitor1Model -eq "ACER V246HL") {
$AcerModelPrefix = "MM"
}

#Get Final Serial # for Monitor 1
IF([string]::IsNullOrEmpty($ActiveMonitors[0].MonitorName)) { 
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Monitor 1 not connected"
} Else {
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Monitor 1 connected"
$Monitor1TrueSerialNumberBegining =  $Monitor1Serial.Substring(0,8)
[int]$Monitor1TrueDecimalSerialNumberMiddle = $Monitor1NumericSerial.Substring(0,10)
$Monitor1TrueHexSerialNumberMiddle = "{0:X}" -F $Monitor1TrueDecimalSerialNumberMiddle
$Monitor1TrueSerialNumberEnding =  $Monitor1Serial.Substring(8)
$Monitor1TrueSerialNumberFinal = $AcerModelPrefix + $Monitor1TrueSerialNumberBegining + $Monitor1TrueHexSerialNumberMiddle + $Monitor1TrueSerialNumberEnding
}


#Get Final Serial # for Monitor 2
IF([string]::IsNullOrEmpty($ActiveMonitors[1].MonitorName)) { 
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Monitor 2 not connected"
} Else {
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Monitor 2 connected"
$Monitor2TrueSerialNumberBegining =  $Monitor2Serial.Substring(0,8)
[int]$Monitor2TrueDecimalSerialNumberMiddle = $Monitor2NumericSerial.Substring(0,10)
$Monitor2TrueHexSerialNumberMiddle = "{0:X}" -F $Monitor2TrueDecimalSerialNumberMiddle
$Monitor2TrueSerialNumberEnding =  $Monitor2Serial.Substring(8)
$Monitor2TrueSerialNumberFinal = $AcerModelPrefix + $Monitor2TrueSerialNumberBegining + $Monitor2TrueHexSerialNumberMiddle + $Monitor2TrueSerialNumberEnding
}

#Get Final Serial # for Monitor 3
IF([string]::IsNullOrEmpty($ActiveMonitors[2].MonitorName)) { 
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Monitor 3 not connected"
} Else {
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Monitor 3 connected"
$Monitor3TrueSerialNumberBegining =  $Monitor3Serial.Substring(0,8)
[int]$Monitor3TrueDecimalSerialNumberMiddle = $Monitor3NumericSerial.Substring(0,10)
$Monitor3TrueHexSerialNumberMiddle = "{0:X}" -F $Monitor3TrueDecimalSerialNumberMiddle
$Monitor3TrueSerialNumberEnding =  $Monitor3Serial.Substring(8)
$Monitor3TrueSerialNumberFinal = $AcerModelPrefix + $Monitor3TrueSerialNumberBegining + $Monitor3TrueHexSerialNumberMiddle + $Monitor3TrueSerialNumberEnding
}

# Write Monitor 1 data to registry
IF([string]::IsNullOrEmpty($ActiveMonitors[0].MonitorName)) { 

} Else {
new-item -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor1' -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor1' -Name "MonitorModel" -PropertyType "String" -Value "$Monitor1Model" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor1' -Name "SerialNumber" -PropertyType "String" -Value "$Monitor1TrueSerialNumberFinal" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor1' -Name "Connected To" -PropertyType "String" -Value "$ComputerName" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor1' -Name "Updated On" -PropertyType "String" -Value "$TimeStamp" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor1' -Name "Last Logged on by" -PropertyType "String" -Value "$LastUserToLogon" -Force
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Created/Updated Monitor 1 registry keys"
}

# Write Monitor 2 data to registry
IF([string]::IsNullOrEmpty($ActiveMonitors[1].MonitorName)) { 
} Else {
new-item -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor2' -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor2' -Name "MonitorModel" -PropertyType "String" -Value "$Monitor2Model" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor2' -Name "SerialNumber" -PropertyType "String" -Value "$Monitor2TrueSerialNumberFinal" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor2' -Name "Connected To" -PropertyType "String" -Value "$ComputerName" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor2' -Name "Updated On" -PropertyType "String" -Value "$TimeStamp" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor2' -Name "Last Logged on by" -PropertyType "String" -Value "$LastUserToLogon" -Force
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Created/Updated Monitor 2 registry keys"
}

# Write Monitor 3 data to registry
IF([string]::IsNullOrEmpty($ActiveMonitors[2].MonitorName)) { 
} Else {
new-item -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor3' -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor3' -Name "MonitorModel" -PropertyType "String" -Value "$Monitor3Model" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor3' -Name "SerialNumber" -PropertyType "String" -Value "$Monitor3TrueSerialNumberFinal" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor3' -Name "Connected To" -PropertyType "String" -Value "$ComputerName" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor3' -Name "Updated On" -PropertyType "String" -Value "$TimeStamp" -Force
new-itemproperty -path 'HKLM:\HARDWARE\DEVICEMAP\VIDEO\MonitorInventory\Monitor3' -Name "Last Logged on by" -PropertyType "String" -Value "$LastUserToLogon" -Force
Add-Content C:\Software\Tools\monitorinfoview\attachedmonitors.log "$TimeStamp - Created/Updated Monitor 3 registry keys"
}

#SCCM Commands https://www.systemcenterdudes.com/configuration-manager-2012-client-command-list/

#Machine Policy Retrieval Cycle #1
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Starting first Machine Policy Retrival Cycle"
Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}"
Start-Sleep -s 30
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - First Machine Policy Retrival Cycle Completed"

#Machine Policy Evaluation Cycle#1
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Starting first Machine Policy Evaluation Cycle"
Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}"
Start-Sleep -s 30
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - First Machine Policy Evaluation Cycle Completed"

#Machine Policy Retrieval Cycle #2
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Starting Second Machine Policy Retrival Cycle"
Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}"
Start-Sleep -s 30
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Second Machine Policy Retrival Cycle Completed"

#Machine Policy Evaluation Cycle#2
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Starting Second Machine Policy Evaluation Cycle"
Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}"
Start-Sleep -s 30
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Second Machine Policy Evaluation Cycle Completed"

#Hardware Inventory Cycle
Start-Sleep -s 90
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Starting Hardware Inventory Cycle"
Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000001}"
Start-sleep -s 30
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Hardware Inventory Cycle Completed"


#Create Success entry in log file
Add-Content C:\Software\Tools\monitorinfoview\attached.monitors.log "$TimeStamp - Script completed."


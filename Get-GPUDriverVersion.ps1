$CurrentDriver = Get-WmiObject Win32_PnPSignedDriver| select devicename, driverversion | where {$_.devicename -like "Intel(R) HD Graphics 4600"}
if($CurrentDriver.driverversion -eq "21.20.16.4627"){
Write-Host "Installed"
}
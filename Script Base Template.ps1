TRY{
#### COMMON VARIABLES ####


#Stores computer name in a vaiable
$ComputerName = [Environment]::MachineName
#$ComputerNameLowercase = $ComputerName.ToLower()

#Stores computer model in a vairable
#$ComputerModelNumber = Get-WmiObject -Class Win32_ComputerSystem | Format-List Model

#Stores computer Serial # in a vairable
#$ComputerSerialNumber = Get-WmiObject win32_bios | Format-List SerialNumber


#Store last last user to login in variable
#$LastUserToLogon = (Get-WmiObject -Class Win32_ComputerSystem).UserName

#Store date/time in variable
$TimeStamp = get-date -format f

#### LOAD ADDITIONAL PSSNAPINs ####

#### LOAD ADDITIONAL PSSNAPINs ####

#Active Directory
#Import-Module ActiveDirectory -ErrorAction SilentlyContinue

#Exchange2010
#add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010

#Exchange2013
#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

#### SCRIPT SPECIFIC VARIABLES ####



#### FUNCTIONS ####
Function Start-PcmSock23{
Set-Location '\Program Files (x86)\ALK Technologies\PMW230\tcpip'
Start-Process pcmsock23 -ArgumentList "PC_MILER 8230"}

Function Start-PcmSock25{
Set-Location '\ALK Technologies\PMW250\tcpip'
Start-Process pcmsock25 -ArgumentList "PC_MILER 8250"}

Function Start-PcmSock30{
Set-Location '\ALK Technologies\PCMILER30\MVS'
Start-Process pcmsock30 -ArgumentList "PC_MILER 8300"}
Function Start-PcmSock31{
Set-Location '\ALK Technologies\PCMILER31\MVS'
Start-Process pcmsock31 -ArgumentList "PC_MILER 8310"}


#### INSTALL APPLICATION/ RUN COMMANDS####


}

CATCH {
########### Only used if an error occurs the the above code ###########
#######################################################################

#Stores error messages in a variables
 $ErrorMessage = $_.Exception.Message
 $FailedItem = $_.Exception.ItemName
 IF ($null -ne $FailedItem){
 #Store date/time in variable
$TimeStamp = get-date -format f
 Add-Content C:\Software\$ComputerName.StriveLogistics.log "$TimeStamp -Attempt to remove unapproved local admin unsuccessful $ErrorMessage"
 }
  IF ($null -ne $FailedItem){
 #Store date/time in variable
$TimeStamp = get-date -format f
 Add-Content C:\Software\$ComputerName.StriveLogistics.log "$TimeStamp - Attempt to remove unapproved local admin unsuccessful $FailedItem"
 }
##################################### Send FAILURE email notification #####################################
###########################################################################################################
}



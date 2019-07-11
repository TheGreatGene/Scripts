# This script will remove the currently activated Office 365 license
# Useful when somebody leaves the company and you disable his/her account or remove their assinged Office 365 license while the license is actively used on a workstation.
# https://jaapwesselius.com/2014/11/12/weve-run-into-a-problem-with-your-office-365-subscription/
# Created on 10/18/2016 by Gene Shelby

#Directions
# 1. Run Script
# 2. Reboot
# 3. Enter new users Office 365 Credentials

cd "C:\Program Files\Microsoft Office\Office16"
Cscript ospp.vbs /dstatus | Select-String -Pattern "Last 5 characters of installed product key:*"
Read-Host -Prompt “Copy the product key and ENTER to continue”
$InputProductKey = Read-Host -Prompt “Enter product key”
Cscript ospp.vbs /unpkey:$InputProductKey
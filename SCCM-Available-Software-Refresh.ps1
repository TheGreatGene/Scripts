###################################################################################
# This script runs SCCM client actions to update the list of available software.
#
#
# Created by: Gene Shelby
#LAST UPDATED: 11/30/2016
###################################################################################

#SCCM Commands https://www.systemcenterdudes.com/configuration-manager-2012-client-command-list/

#Machine Policy Retrieval Cycle	
Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}"
Start-sleep -s 15
#Machine Policy Evaluation Cycle
Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}"


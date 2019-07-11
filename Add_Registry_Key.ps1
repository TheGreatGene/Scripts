# Script to add a Registry Key

New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook -Name "ForcePSTPath" -PropertyType "ExpandString" -Value '%LOCALAPPDATA%\Microsoft\Outlook'

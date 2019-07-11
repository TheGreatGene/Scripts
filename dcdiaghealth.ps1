<# DCDiag Health Report

Description: Daily / on demand script that checks AD health using DCDiag and emails out report.

Source: The internet.

Version Control:

1 - Tim Sutton
 - Initial Implmentation
 - Minor tweaks from source to suit our environment.

#>



Function sendEmail ([String] $body)
{
	$MailMessage = New-Object System.Net.Mail.MailMessage
	$MailMessage.From = "adhealth@strivelogistics.com"
	$MailMessage.To.Add("chamill@strivelogistics.com")
	$MailMessage.Subject = "DCDiag Health Summary v2"
	$MailMessage.Body = $body
	#$MailMessage.Priority = "High"
	$MailMessage.IsBodyHtml = $True

	$SMTPClient = New-Object System.Net.Mail.SMTPClient
	$SMTPClient.Host = "mail.strivelogistics.com"
	$SMTPClient.Send($MailMessage)
}





Function convertToVertical ([String] $testname)
{

$stringlength = $testname.Length

for ($i=0; $i -lt $stringlength; $i++)
{
	$newname = $newname + "<BR>" + $testname.Substring($i,1)
}

$newname 
}






import-module ActiveDirectory
$ADInfo=Get-ADDomain
$allDCs=$ADInfo.ReplicaDirectoryServers

$testnamecount=0

$a = "<style>"
$a = $a + "body{color:#717D7D;background-color:#F5F5F5;font-size:8pt;font-family:'trebuchet ms', helvetica, sans-serif;font-weight:normal;padding-:0px;margin:0px;overflow:auto;}"
#$a = $a + "a{font-family:Tahoma;color:#717D7D;Font-Size:10pt display: block;}"
$a = $a + "table,td,th {font-family:Tahoma;color:Black;Font-Size:8pt}"
$a = $a + "th{font-weight:bold;background-color:#ADDFFF;}"
#$a = $a + "td {background-color:#E3E4FA;text-align: center}"
$a = $a + "</style>"

##############################
foreach ($item in $allDCs)
{
	$logfile = "C:\_Scripts\DCDiagHealth\dcdiag_$item.txt"
	#Dcdiag.exe /v /s:$item >> $logfile
	Dcdiag.exe /s:$item >> $logfile

	#New-Variable "AllResults$item" -force
	#$c = $AllResults + $item
	 
	$AllResults = New-Object Object
	$AllResults | Add-Member -Type NoteProperty -Name "ServerName" -Value $item
	#$TestCat = $Null
	
	$table+="<tr><td>$item</td>"
	Get-Content $logfile | %{
		Switch -RegEx ($_)
		{
			 #"Running"       { $TestCat    = ($_ -Replace ".*tests on : ").Trim() }
			 "Starting"      { $TestName   = ($_ -Replace ".*Starting test: ").Trim() }
			 "passed|failed" { If ($_ -Match "passed") { 
			 $TestStatus = "Psd" 
			  } Else { 
			 $TestStatus = "Fld" 
			  } }
		}

		If ($TestName -ne $Null -And $TestStatus -ne $Null)
		{
			$TestNameVertical = convertToVertical($TestName)


			$AllResults | Add-Member -Name $("$TestNameVertical".Trim()) -Value $TestStatus -Type NoteProperty -force
		 
			if($TestStatus -eq "Fld"){
				$table+="<td style=""background-color:red;"">$TestStatus</td>"
			}else{
				$table+="<td style=""background-color:green;"">$TestStatus</td>"
			}
			
			if($testnamecount -lt 29){
                $allTestNames = $allTestNames + "<BR>" + $TestName
				$testnames+="<td style=""background-color:#CCE3FF;"">$TestNameVertical</td>"
				#$testnames+="<td class=""titlestyle"">t<BR>e<BR>s<BR>t</td>"
				$testnamecount++
			}
				   
			$TestName = $Null; $TestStatus = $Null
		}
		New-Variable "last$item" -force -Value $AllResults
	}
	$table+="</tr>"
	Remove-Item $logfile
	
}



	
	
	
	
	

$html="<html><head>$a</head><table><tr><td>S<BR>e<BR>r<BR>v<BR>e<BR>r<BR>N<BR>a<BR>m<BR>e</td>" +$testnames + "</tr>" + $table + "</table><BR><BR>Tests ran: $allTestNames</html>"
#$html | out-file "C:\_Scripts\DCDiagHealth\final.html"
$body = $html | out-string
sendEmail $body
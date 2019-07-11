Get-aduser *UserName*| Where {$_.whencreated -like 5/10/2018}



$AllADUsersCreationDate = Get-aduser *UserName* -properties whencreated


$AllADUsersCreationDate $AllADUsersCreationDate.whencreated
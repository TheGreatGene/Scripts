###
#https://devblogs.microsoft.com/scripting/weekend-scripter-manipulating-word-and-excel-with-powershell/
Function OpenWordDoc($Filename){

$Word=NEW-Object –comobject Word.Application

Return $Word.documents.open($Filename)

}

Function SearchAWord($Document,$findtext,$replacewithtext){ 

  $FindReplace=$Document.ActiveWindow.Selection.Find

  $matchCase = $false;

  $matchWholeWord = $true;

  $matchWildCards = $false;

  $matchSoundsLike = $false;

  $matchAllWordForms = $false;

  $forward = $true;

  $format = $false;

  $matchKashida = $false;

  $matchDiacritics = $false;

  $matchAlefHamza = $false;

  $matchControl = $false;

  $read_only = $false;

  $visible = $true;

  $replace = 2;

  $wrap = 1;

  $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl)

}

Function SaveAsWordDoc($Document,$FileName){

$Document.Saveas([REF]$Filename)

$Document.close()

}

Function OpenExcelBook($FileName){

$Excel=new-object -ComObject Excel.Application

Return $Excel.workbooks.open($Filename)

}

Function SaveExcelBook($Workbook){

$Workbook.save()

$Workbook.close()

}


$Workbook=OpenExcelBook –Filename '\\servername\users-Template-Final.xlsx'
Function ReadCellData($Workbook,$Cell){

$Worksheet=$Workbook.Activesheet

Return $Worksheet.Range($Cell).text

}

$Row=2

Do{ 
$DisplaynameData=ReadCellData -Workbook $Workbook -Cell "C$Row"

If ($Data.length –ne 0) {
$DOC = OpenWordDoc -Filename '\\servername\Welcome Doc_AA_PoweShell_TEMPLATE.docx'


SearchAWord -Document $Doc -findtext '***DisplayName***' -replacewithtext $DisplayNameData

#####  START DATE   #####
$StartDateData=ReadCellData -Workbook $Workbook -Cell "Q$Row"
SearchAWord -Document $Doc -findtext '***StartDate***' -replacewithtext $StartDateData

#####  POSITION    #####
$PositionData=ReadCellData -Workbook $Workbook -Cell "H$Row"
SearchAWord -Document $Doc -findtext '***Position***' -replacewithtext $PositionData

#####   SUPERVISOR   #####
$SupervisorData=ReadCellData -Workbook $Workbook -Cell "I$Row"
SearchAWord -Document $Doc -findtext '***Supervisor***' -replacewithtext $SupervisorData

##### DEPARTMENT ###
$DepartmentData=ReadCellData -Workbook $Workbook -Cell "R$Row"
SearchAWord -Document $Doc -findtext '***Department***' -replacewithtext $DepartmentData

##### HOME OFFICE ######
$OfficeLocationData=ReadCellData -Workbook $Workbook -Cell "E$Row"
SearchAWord -Document $Doc -findtext '***OfficeLocation***' -replacewithtext $OfficeLocationData


##### USER NAME ####
$UserNameData=ReadCellData -Workbook $Workbook -Cell "D$Row"
SearchAWord -Document $Doc -findtext '***UserName***' -replacewithtext $UserNameData


##### EMAIL #####
$EmailData=ReadCellData -Workbook $Workbook -Cell "F$Row"
SearchAWord -Document $Doc -findtext '***Email***' -replacewithtext $EmailData

##### ADPassword #####
$ADPasswordData=ReadCellData -Workbook $Workbook -Cell "X$Row"
SearchAWord -Document $Doc -findtext '***Password***' -replacewithtext $ADPasswordData

##### PhoneNumber #####
$PhoneNumberData=ReadCellData -Workbook $Workbook -Cell "S$Row"
SearchAWord -Document $Doc -findtext '***PhoneNumber***' -replacewithtext $PhoneNumberData

##### EXT #####
$ExtData=ReadCellData -Workbook $Workbook -Cell "V$Row"
SearchAWord -Document $Doc -findtext '***Ext***' -replacewithtext $ExtData

##### FaxNumber #####
$FaxNumberData=ReadCellData -Workbook $Workbook -Cell "W$Row"
SearchAWord -Document $Doc -findtext '***FaxNumber***' -replacewithtext $FaxNumberData

##### ComputerName #####
$ComputerNameData=ReadCellData -Workbook $Workbook -Cell "Y$Row"
SearchAWord -Document $Doc -findtext '***ComputerName***' -replacewithtext $ComputerNameData

##### TV-ID #####
$TVIDData=ReadCellData -Workbook $Workbook -Cell "Z$Row"
SearchAWord -Document $Doc -findtext '***TVID***' -replacewithtext $TVIDData


##### PhoneMAC #####
$PhoneMACData=ReadCellData -Workbook $Workbook -Cell "U$Row"
SearchAWord -Document $Doc -findtext '***PhoneMAC***' -replacewithtext $PhoneMACData


##### PhoneID #####
$PhoneIDData=ReadCellData -Workbook $Workbook -Cell "T$Row"
SearchAWord -Document $Doc -findtext '***PhoneID***' -replacewithtext $PhoneIDData


$SaveName="\\servername\Welcome_Doc_$UserNameData.docx"

SaveAsWordDoc –document $Doc –Filename $Savename

$Row++

}

$Data=ReadCellData -Workbook $Workbook -Cell "A$Row" 
}Until ($Data -eq "")

SaveExcelBook –workbook $Workbook




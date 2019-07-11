### Part I: This PS script shows you what packages are in WMI and not in the Content Library AND vice versa. 

$WMIPkgList = Get-WmiObject -Namespace Root\SCCMDP -Class SMS_PackagesInContLib | Select -ExpandProperty PackageID | Sort-Object

$ContentLib = (Get-ItemProperty -path HKLM:SOFTWARE\Microsoft\SMS\DP -Name ContentLibraryPath)

$PkgLibPath = ($ContentLib.ContentLibraryPath) + "\PkgLib"

$PkgLibList = (Get-ChildItem $PkgLibPath | Select -ExpandProperty Name | Sort-Object)

$PkgLibList = ($PKgLibList | ForEach-Object {$_.replace(".INI","")})

$PksinWMIButNotContentLib = Compare-Object -ReferenceObject $WMIPkgList -DifferenceObject $PKgLibList -PassThru | Where-Object { $_.SideIndicator -eq "<=" } 

$PksinContentLibButNotWMI = Compare-Object -ReferenceObject $WMIPkgList -DifferenceObject $PKgLibList -PassThru | Where-Object { $_.SideIndicator -eq "=>" } 

Write-Host Items in WMI but not the Content Library

Write-Host ========================================

$PksinWMIButNotContentLib

Write-Host Items in Content Library but not WMI

Write-Host ====================================

$PksinContentLibButNotWMI
$PksinContentLibButNotWMI.Count







### Part II: This PS script removes the package from WMI (using the list from Part I):

Foreach ($Pkg in $PksinWMIButNotContentLib){ Get-WmiObject -Namespace Root\SCCMDP -Class SMS_PackagesInContLib -Filter "PackageID = '$Pkg'" | Remove-WmiObject -Confirm }





### Part III: This PS script removes the INI file (using the list from Part I):

Foreach ($Pkg in $PksinContentLibButNotWMI){ Remove-Item -Path "$PkgLibPath\$Pkg.INI" -Confirm }
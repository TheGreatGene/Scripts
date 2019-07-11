Set-ExecutionPolicy Unrestricted
Import-PackageProvider -Name Nuget

IF (!(Get-PackageProvider -Name Nuget)){

Install-PackageProvider -Name NuGet -force -Confirm:$false

} Else {

# NuGet Already Installed
}

Import-Module -Name Carbon

IF (!(Get-Module -Name Carbon)){

Install-Module -Name Carbon -Force -Confirm:$false -AllowClobber

} Else {
#Carbon already installed
}

Set-ExecutionPolicy RemoteSigned
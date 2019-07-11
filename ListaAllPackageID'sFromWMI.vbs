strComputer  = "."
strNamespace = "rootsccmdp"
 
Set objCol = GetObject("winmgmts:\" & strComputer & "" & strNamespace).InstancesOf("SMS_PackagesInContLib")
For Each Package In objCol
       WScript.Echo Package.PackageID
Next
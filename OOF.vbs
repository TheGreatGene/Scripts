'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------

Call Main

' ################################################
' The starting point of execution for this script.
' ################################################
Sub Main()
	If CheckScriptEngineName("cscript.exe") = True Then
		
		Dim wsArgNum
		
		' Get the number of parameters from command-line.
		wsArgNum = WScript.Arguments.Count
		
		Select Case wsArgNum
			Case 1
				Dim strFirstArg
				strFirstArg = UCase(WScript.Arguments(0))
				
				If strFirstArg = "TRUE" Then
					ToggleOOFState CBool(strFirstArg)
				ElseIf strFirstArg = "FALSE" Then
					ToggleOOFState CBool(strFirstArg)
				Else
					WScript.Echo "Invalid argument."
				End If
				
			Case Else
				WScript.Quit
		End Select
	Else
		WScript.Echo "This VBScript file can be only run using the <cscript.exe> engine."
	End If
End Sub

' #########################################################################################
' Check if the current engine name for executing script is the same as the given name.
' The Name can be "WScript.exe (default)" or "CScript.exe", and it does not case sensitive.
' #########################################################################################
Function CheckScriptEngineName(Byval Name)
	' Set the default return value to False.
	CheckScriptEngineName = False
	
	Dim strTempName
	
	strTempName = UCase(Name)
	
	Select Case strTempName
		Case "WSCRIPT.EXE", "CSCRIPT.EXE"
			If InStrRev(UCase(WScript.FullName), strTempName, -1, 1) <> 0 Then
				' Change the return value to True.
				CheckScriptEngineName = True
			End If
	End Select
End Function

' ###############################
' Toggle the Out of Office state.
' ###############################
Public Sub ToggleOOFState(Byval OOFState)
    Dim olApp
    Dim blnIsCreated
    Dim blnOOFState
    Dim stoItem
    Dim stoItems
    Dim olPropAccessor
    
    On Error Resume Next
    
    If SingleInstanceModeForAutomationObject(olApp, blnIsCreated, "Outlook.Application") = True Then
        
        ' The name of the property whose value is to be returned.
        Const PR_OOF_STATE = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
        
        ' Set reference to the Stores collection.
        Set stoItems = olApp.GetNamespace("MAPI").Stores
        
        ' /* Step through each store object in the stores collection. */
        For Each stoItem In stoItems
            
            ' /* a primary Exchange mailbox store. */
            If stoItem.ExchangeStoreType = olPrimaryExchangeMailbox Then
                'Obtain an instance of PropertyAccessor class.
                Set olPropAccessor = stoItem.PropertyAccessor
                
                ' Get the Out of Office state.
                blnOOFState = olPropAccessor.GetProperty(PR_OOF_STATE)
                
                ' /* Toggle the Out of Office state if the old state is not the same as the given state. */
                If blnOOFState <> OOFState Then Call olPropAccessor.SetProperty(PR_OOF_STATE, OOFState)
            End If
            
        Next
        
    End If
    
    If blnIsCreated = True Then olApp.Quit
    
    ' /* Release memory. */
    Set olApp = Nothing
    Set stoItem = Nothing
    Set stoItems = Nothing
    Set olPropAccessor = Nothing
End Sub

' ################################################################
' Use the single instance mode for a predefined automation object.
' ################################################################
Public Function SingleInstanceModeForAutomationObject(AutomationObject, IsCreated, Byval ClassName)
    Dim blnResult
    
    ' The default return value is False.
    blnResult = False
    ' An instance of the automation object is already running.
    IsCreated = False
    
    ' Ignore the exception.
    On Error Resume Next
    
    ' Try to get an instance of the automation object.
    Set AutomationObject = GetObject(, ClassName)
    
    ' /* Couldn't get the instance. */
    If Err.Number <> 0 Then
        ' Clear all property settings of the Err object.
        Err.Clear
        
        ' Try to create an instance of the automation object.
        Set AutomationObject = CreateObject(ClassName)
        
        ' /* The instance was created successfully. */
        If Err.Number = 0 Then blnResult = True : IsCreated = True
    Else
        blnResult = True
    End If
    
    SingleInstanceModeForAutomationObject = blnResult
End Function
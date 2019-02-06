' Get Working Directory
Dim Ws, strPath, objFSO, objFile
Set Ws = CreateObject("WScript.Shell")

' Get Script Location
strPath = Wscript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strPath)
strCurDir = objFSO.GetParentFolderName(objFile) 

' Place Holder
Dim oaccess

' If Access is Closed, run script
Do While Process_Terminate("MSACCESS.exe") = True
Loop

' create access object
set oaccess = createobject("access.application")
' Open Database2 (assuming in the current directory)
oaccess.opencurrentdatabase strCurDir & "\Database2.accdb", FALSE
' Activate Macro
oaccess.docmd.runMacro "HTMLOUT"
' Close the File
oaccess.closecurrentdatabase
' Close Access
oaccess.quit
set oaccess=nothing
MsgBox "Script Exit - DataBase was Updated", 0,"Script Update"
WScript.Quit(0)


Public Function Process_Terminate (ByVal sProcName)	
' This function terminates a process by its name, and returns "True" when succeeds, and  "False" - otherwise
'	Input parameters:
'		sProcName	-	Process name
'	Output parameters:
'		none
'	Usage:
'		rc = Process_Terminate ("Aces.exe")
'		 
	Dim oProcess, colProcessList, objWMIService, WshNetwork
	Set WshNetwork = CreateObject("WScript.Network")
	Set objWMIService = GetObject("winmgmts:\\" & WshNetwork.ComputerName)
	Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")
	Process_Terminate = FALSE
	For Each oProcess in colProcessList
		if StrComp(UCase(oProcess.Name), UCase(sProcName)) = 0 Then
			'oProcess.Terminate
			'MsgBox "Access is Open"
			'Wscript.Sleep(500)
			Process_Terminate = TRUE
		End if
	Next
	'if Process_Terminate = FALSE Then MsgBox "Access is Close"
	Set WshNetwork = Nothing
	Set objWMIService = Nothing
	Set colProcessList = Nothing
End Function
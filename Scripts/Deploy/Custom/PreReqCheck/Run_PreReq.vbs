' // ***************************************************************************
' //
' // File:      Run_PreReq.vbs
' // 
' // Version:   1
' // 
' // Purpose:   Validates Model
' // 
' // Usage:     cscript Run_PreReq.vbs
' //
' // Return:	   0	Success
' //			4201	Error - LAN Check
' //			4202	Error - Model Check
' //			4203	Error - SecureDoc Check
' //
' // Created:   1.0 Brandon Hilgeman (2/4/2015)
' //
' // Modified:  
' // 
' // ***************************************************************************

Option Explicit

'//----------------------------------------------------------------------------
'//  Open Library
'//----------------------------------------------------------------------------

Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim sScriptPath : sScriptPath = Replace(WScript.scriptfullname,WScript.scriptname,"")
Dim sLibraryPath : sLibraryPath = sScriptPath & "ScriptLibrary_2.08.vbs"
If oFSO.FileExists(sLibraryPath) Then
    Dim sLibrary : sLibrary = oFSO.OpenTextFile(sLibraryPath).ReadAll
    ExecuteGlobal sLibrary
Else 
    WScript.echo "  ERROR: Unable to open Library File"
    WScript.Quit(10)
End If

'//----------------------------------------------------------------------------
'//  Configure
'//----------------------------------------------------------------------------

'//----------------------------------------------------------------------------
'//  Main
'//----------------------------------------------------------------------------

LogStart

'CheckLAN
CheckModel

LogEnd

'//----------------------------------------------------------------------------
'//  Functions and SubRoutines
'//----------------------------------------------------------------------------

Sub CheckLAN
	WriteToLog "CheckLAN"
	
	Dim sMSG, iRet

	'Display LanPrompt
	If IsLanConnected = False Then
		sMSG = "ERROR: LAN is not Connected" & vbCrLf & vbCrLf & "The ETHERNET cable must be plugged in to perform deployment" & vbCrLf & "Please connect the cable and try again"
		WriteToLog sMSG
		HideTS
		iRet = MsgBox(sMSG, 4144, "LAN Check Error")
		LogEnd
		WScript.Quit(4201)
	End If	
	
End Sub

Sub CheckModel
	WriteToLog "CheckModel"
	
	Dim bIsSupported
	Dim sMSG, iRet
	
	bIsSupported = False

	'Find Model
	Dim sModel, oItem, cItems
	Set cItems = oWMI.ExecQuery("Select * from Win32_ComputerSystem")
	For Each oItem In cItems
		sModel = oItem.Model
	Next
	
	'Detect Model in List
	Dim sText, sLine
	sText = oFSO.OpenTextFile(sScriptPath & "PreReq_Models.txt",1).ReadAll
	For Each sLine In Split(sText,vbCrLf)
		If InStr(1,sModel,sLine,1) And Trim(sLine) <> "" Then
			WriteToLog "** Detected Supported Model: " & sModel & " (" & sLine & ")"
			bIsSupported = True
			Exit For
		End If
	Next	
	
	'Display Prompt
	If bIsSupported = False Then
		sMSG = "ERROR: Unsupported Model" & vbCrLf & vbCrLf & "Model: " & sModel & vbCrLf & "Supported: " & bIsSupported & vbCrLf & vbCrLf & "Deployment of this model is not currently supported." & vbCrLf & "The Task Sequence will now terminate"
		WriteToLog sMSG
		HideTS
		iRet = MsgBox(sMSG, 4144, "Unsupported Model Error")
		LogEnd
		WScript.Quit(4202)
	End If
End Sub

Sub HideTS
	On Error Resume Next
	Dim oTS
	Set oTS = CreateObject("Microsoft.SMS.TsProgressUI")
	oTS.CloseProgressDialog
	Err.Clear
End Sub
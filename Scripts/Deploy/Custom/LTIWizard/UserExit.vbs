Function UserExit(sType, sWhen, sDetail, bSkip)
	UserExit = Success
End Function

'---------------------------- Initialize --------------------------------
Dim INIFile, FSO, LogFile, DeployRoot, sOS, GetOSName, OSVersion
Set FSO = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject("WScript.Shell")
tempFolder = fso.GetSpecialFolder(2)
INIFile = "X:\MININT\SMSOSD\OSDLOGS\LTIAnswer.INI"
LogFile = "X:\MININT\SMSOSD\OSDLOGS\UserExit.log"



' Set LTI INI answer file and log path in MININT\SMSOSD\OSDLogs
If FSO.FolderExists("C:\MININT") Then
	INIFile = "C:\MININT\SMSOSD\OSDLOGS\LTIAnswer.INI"
	LogFile = "C:\MININT\SMSOSD\OSDLOGS\UserExit.log"
End If

' Create log file if not already started
If Not FSO.FileExists(LogFile) Then
	Set Logtxt = FSO.CreateTextFile(LogFile, True)
	Logtxt.Close
	Log("Start " & WScript.ScriptName)
End If

' Run LTI and create answer file
Function RunLTI
	If Not FSO.FileExists(INIFile) Then
		Dim sCMD : sCMD = "Powershell.exe -WindowStyle Hidden -executionpolicy bypass -file ""%DEPLOYROOT%\Scripts\Custom\LTIWizard\LTIWizard.ps1"" %COSDCOMPUTERNAME%" 
		Log sCMD
		Call oShell.Run(sCMD, 1, True)
	End If
End Function


' Functions called from CustomSettings.ini
Function OSDComputerName
	OSDComputerName = iniRead (INIFILE,UCase("MDTLTI"),UCase("OSDComputerName"))
	Log("OSDComputerName: " & OSDComputerName)
End Function

Function cTS
	cTS = iniRead (INIFILE,UCase("MDTLTI"),UCase("cTS"))
	Log("cTS value: " & cTS)
End Function

Function cTSList
	cTSList = iniRead (INIFILE,UCase("MDTLTI"),UCase("cTSList"))
	Log("cTSList: " & cTSList)
End Function

Function cAppProfile
	cAppProfile = iniRead (INIFILE,UCase("MDTLTI"),UCase("cAppProfile"))
	Log("cAppProfile: " & cAppProfile)
End Function

Function cOperatingSystem
	cOperatingSystem = INIRead (INIFILE,UCase("MDTLTI"),UCase("cOS"))
	Log("cOS: " & cOperatingSystem)
End Function

Function cDomain
	cDomain = INIRead (INIFILE,UCase("MDTLTI"),UCase("cDomain"))
	Log("cDomain: " & cDomain)
End Function

Function cImagingOU
	cImagingOU = INIRead (INIFILE,UCase("MDTLTI"),UCase("cImagingOU"))
	Log("cImagingOU: " & cImagingOU)
End Function

Function cTSList
	cTSList = INIRead (INIFILE,UCase("MDTLTI"),UCase("cTSList"))
	Log("cTSList: " & cTSList)
End Function

Function cImagingOU
	cImagingOU = INIRead (INIFILE,UCase("MDTLTI"),UCase("cImagingOU"))
	Log("cImagingOU: " & cImagingOU)
End Function

Function cTargetOU
	cTargetOU = INIRead (INIFILE,UCase("MDTLTI"),UCase("cTargetOU"))
	Log("cTargetOU: " & cTargetOU)
End Function



'---------------------------- Log Function --------------------------------
Function Log(text)
	Set LogWrite= FSO.OpenTextFile(LogFile, 8, True)
	LogWrite.WriteLine(Date & " " & Time & ": " & text)
	LogWrite.Close
End Function
'--------------------------------------------------------------------------

'------------------------- INI Read Function ------------------------------
Function INIRead(strFile,strSection,strKey)
	Dim objReadFile,strLine,blnFoundSect,blnReadAll,intKeySize
	If Not FSO.FileExists(strFile) Then
		INIRead = "*FILEERR*" : Exit Function
	End If
	Set objReadFile = FSO.OpenTextFile(strFile,1)
	blnFoundSect = False : blnReadAll=False : intKeySize = Len(strKey)+1
	Do While Not objReadFile.AtEndOfStream
		strLine= objReadFile.ReadLine
		If blnFoundSect Then
			If Left(strLine,1)="[" Then
				If blnReadAll=False Then INIRead = "*KEYERR*"
				objReadFile.close : Exit Function
			End If
			If blnReadAll Then
				INIRead = IniRead & strLine & vbCrLf
			Else
				If Left(strLine,intKeySize) = strKey & "=" Then
					INIRead=Right(strLine,Len(strLine)-intKeySize)
					objReadFile.close : Exit Function
				End If
			End If
		Else
			If Left(strLine,1)="[" Then
				If strLine = "[" & strSection & "]" Then
					blnFoundSect=True : If strKey="" Then blnReadAll=True
				End If
			End If
		End If
	Loop
	If blnReadAll=False Then
		If blnFoundSect Then
			INIRead = "*KEYERR*"
		Else
			INIRead = "*SECTERR*"
		End If
	End If
	objReadFile.close
End Function
'--------------------------------------------------------------------------
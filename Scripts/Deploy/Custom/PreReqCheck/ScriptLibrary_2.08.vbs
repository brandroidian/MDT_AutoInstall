'-------------------------------------------------------------------------------
'----
'----   Script Library
'----   
'----   File:       Script_Library_2.03.vbs
'----   
'----   Version:    2013.10.30.2.03
'----   
'----   Purpose:    Contains common functions and routines useful use in other scripts
'----   
'----   Modified:   2013.10.09.10.00 (2.01)
'----               Created Function: CopyFolder
'----               Modified Function: CopyFile - Includes additional error handling
'----   
'-------------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------------
'---    Global Objects
'-------------------------------------------------------------------------------

Dim oShell : Set oShell = WScript.CreateObject("WScript.Shell")
Dim oNetwork : Set oNetwork = WScript.CreateObject("WScript.Network")
Dim oWMI : Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
Dim oReg : Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv")

'-------------------------------------------------------------------------------
'---    Global Variables
'-------------------------------------------------------------------------------

Dim sComputerName: sComputerName = oNetwork.ComputerName
Dim sWindowsPath: sWindowsPath = oShell.expandenvironmentstrings("%WINDIR%") & "\"
Dim sProgramPath: sProgramPath = oShell.ExpandEnvironmentStrings("%PROGRAMFILES%") & "\"
Dim sTempPath : sTempPath = oShell.ExpandEnvironmentStrings("%TEMP%") & "\"

Dim sScriptStart, sScriptEnd

'-------------------------------------------------------------------------------
'---    Default Global Variables
'-------------------------------------------------------------------------------

Dim sLog : sLog = "c:\Windows\GHS\" & WScript.ScriptName & ".log"
Dim sLogMSI : sLogMSI = "c:\Windows\GHS\" & WScript.ScriptName & ".MSI.log"
Dim sMode: sMode = "INSTALL"
Dim bLog : bLog = False

'-------------------------------------------------------------------------------
'---    Global Constants
'-------------------------------------------------------------------------------

'---    File Open Operations
Const cForRead = 1
Const cForWrite = 2
Const cForAppend = 8

'---    Run Operations
Const cHidden = 0
Const cDisplay = 1
Const cWait = True
Const cDoNotWait = False

'---    Registry
Const cHKLM = &H80000002

'-------------------------------------------------------------------------------
'---    Common Functions
'-------------------------------------------------------------------------------

'---    WriteToLog
'---    Writes string to log
'---    Parameter1: String - Text to be written to log
Sub WriteToLog(sText)
	If bLog = True Then
		On Error Resume Next
	    Dim oFile
	    Set oFile = oFSO.OpenTextFile(sLog, 8, True)
	        oFile.WriteLine "<![LOG[" & sText & "]LOG]!><time=""" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now) & ".000+000"" date=""" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & """ component=""" & Wscript.ScriptName & """ context="""" type=""1"" thread="""" file=""" & wscript.ScriptName & """>"
	    oFile.Close
	    Err.Clear
    End If
End Sub

'---    LogStart
Sub LogStart
	sScriptStart = CDate(Now)
	CreateLogFolder
	bLog = True
    WriteToLog "------------ Start Script: " & WScript.ScriptName & " " & Now & " ------------"
End Sub

'---    LogEnd
Sub LogEnd
	sScriptEnd = CDate(Now)
    WriteToLog "------------ End Script: " & Now & " Runtime: " & DateDiff("n",sScriptStart,sScriptEnd) & " minutes" & " ------------"
    bLog = False
End Sub

'---    SetMode
'---    Checks arguments to detect uninstall
Sub SetMode
	If WScript.Arguments.Count > 0 Then
		Dim cArgs, i
		Set cArgs = WScript.Arguments
		For i = 0 To cArgs.Count - 1
			If UCase(cArgs(i)) = "/UNINSTALL" Then
				sMode = "UNINSTALL"
			ElseIf UCase(cArgs(i)) = "/UPGRADE" Or UCase(cArgs(i)) = "/UPDATE" Then
				sMode = "UPGRADE"
			End If
		Next
	End If
	WriteToLog "** SetMode: " & sMode
	
End Sub

'---    CreateLogFolder
'---    Creates Log Folder If it doesn't exist
Sub CreateLogFolder
	If oFSO.FolderExists(sLog) = False Then
		WriteToLog "** CreateLogFolder"
		CreateFolderStructure(sLog)
	End If
End Sub

'---    CreateFolderStructure
'---    Recursively creates folder.
'---    Parameter1: String - Folder to Create
Sub CreateFolderStructure(sFolder)
	WriteToLog "** CreateFolderStructure"
	WriteToLog "  """ & sFolder & """"
	Dim sNewFolder
	Dim iCheck : iCheck = 1
	Do Until iCheck = 0
		iCheck = InStr(iCheck + 1,sFolder,"\",1)
		If iCheck <> 0 Then
			sNewFolder = Left(sFolder,iCheck - 1)
			If oFSO.FolderExists(sNewFolder) = False Then
				oFSO.CreateFolder sNewFolder
				WriteToLog "  Folder Created: """ & sNewFolder & """"
			End If
		End If
	Loop
End Sub

'---    CopyFile
'---    Copy's file with logging
'---    Parameter1: String - Folder to Create
Function Copyfile(sSource, sDestination)
	Dim sPath, iCheck
	
	WriteToLog "** Copyfile"
	WriteToLog "  Source: """ & sSource & """"
	WriteToLog "  Destination: """ & sDestination & """"
	
	'Ensure destination folder exists
	If oFSO.FileExists(sDestination) = False Then
		sPath = sDestination
		iCheck = 1
		Do Until iCheck = 0
			sPath = Left(sDestination, iCheck)
			iCheck = InStr(iCheck+1,sDestination,"\")
		Loop
		
		If oFSO.FolderExists(sPath) = False Then
			CreateFolderStructure sPath
		End If	
	End If

	If oFSO.FileExists(sSource) Then
		Dim iret
		iret = oFSO.copyfile(sSource,sDestination,True)
		If iret=0 Then
			WriteToLog "Copy Completed: Success"
		Else
			WriteToLog "Copy Completed: Unsucessful"
		End If
		Copyfile = iret
	Else
		WriteToLog "Copy Completed: Unsucessful (Invalid Source Path)"
	End If
	
End Function

'---    CopyFolder
'---    Copy's Folder with logging
'---    Parameter1: String - Folder to copy
'---    Parameter2: String - Folder Destination
Function CopyFolder(sSource, sDestination)
	Dim sPath, iCheck
	
	WriteToLog "** CopyFolder"
	WriteToLog "  Source: """ & sSource & """"
	WriteToLog "  Destination: """ & sDestination & """"
	
	'Ensure destination folder exists
	If oFSO.FolderExists(sDestination) = False Then
		CreateFolderStructure sDestination
	End If
	
	'Remove Trailing slash
	If Right(sSource,1) = "\" Then
		sSource = Left(sSource,Len(sSource)-1)
	End If
	If Right(sDestination,1) = "\" Then
		sDestination = Left(sDestination,Len(sDestination)-1)
	End If

	If oFSO.FolderExists(sSource) Then
		Dim iret
		iret = oFSO.copyFolder(sSource,sDestination,True)
		If iret=0 Then
			WriteToLog "Copy Completed: Success"
		Else
			WriteToLog "Copy Completed: Unsucessful"
		End If
		CopyFolder = iret
	Else
		WriteToLog "Copy Completed: Unsuccessful (Invalid Source Path)"
	End If
	
End Function

Function FileExists(sFile)
	FileExists = True
	WriteToLog "**Check if """ & sFile & """ exists on system"
	If oFSO.FileExists(sFile) Then
		WriteToLog "**Success: """ & sFile & """ Exists on this system"
		FileExists = True
	Else
		WriteToLog "**File: """ & sFile & """ Does not exist on this system"
		FileExists = False
	End If	
		
End Function

Function FolderExists(sFolder)
	FolderExists = True
	WriteToLog "**Check if """ & sFolder & """ exists on system"
	If oFSO.FolderExists(sFolder) Then
		WriteToLog "**Success: """ & sFolder & """ Exists on this system"
		FolderExists = True
	Else
		WriteToLog "**Folder: """ & sFolder & """ Does not exist on this system"
		FolderExists = False
	End If
End Function

Function DeleteFile(sFile)
	If FileExists(sFile) Then
		oFSO.DeleteFile sFile,True
		If FileExists(sFile) = False Then
			WriteToLog "**Deleted: " & sFile & """ Successfully"""
		Else
			WriteToLog "**Could Not Delete: " & sFile & ""
		End If
	End If
End Function

Function DeleteFolder(sFolder)
	If FolderExists(sFolder) Then
		oFSO.DeleteFolder sFolder,True
		If FolderExists(sFolder) = False Then
			WriteToLog "**Deleted: " & sFolder & """ Successfully"""
		Else
			WriteToLog "**Could Not Delete: " & sFolder & ""
		End If
	End If
End Function



'-------------------------------------------------------------------------------
'---    OS Information
'-------------------------------------------------------------------------------

'---    GetOSName
'---    Returns OS Name
'---    Example: "Microsoft Windows 7 Enterprise"
Function GetOSName
	Dim cItems, oItem
	Set cItems = oWMI.ExecQuery("Select * from Win32_OperatingSystem")
	For Each oItem In cItems
		GetOSName = oItem.Caption
	Next
	WriteToLog "** GetOSName Returned: " & GetOSName
End Function

'---    GetShortOS
'---    Returns shortened version of the OS name
'---    Parameter1: String - Text containing full or longer OS name	
Function GetShortOS (sOS)
    If instr(sOS, "XP") <> 0 Then
        GetShortOS = "XP"
    ElseIf InStr(sOS, "VISTA") <> 0 Then
        GetShortOS = "VISTA"
    ElseIf InStr(sOS, "7") <> 0 Then
            GetShortOS = "7"            
    End If
    WriteToLog "** GetShortOS Returned: " & GetShortOS
End Function
	
'---    GetOSVersion
'---    Returns OS Version
'---    Example: "6.1.7600"
Function GetOSVersion
	Dim cItems, oItem
	Set cItems = oWMI.ExecQuery("Select * from Win32_OperatingSystem")
	For Each oItem In cItems
		GetOSVersion = oItem.Version
	Next
	WriteToLog "** GetOSVersion Returned: " & GetOSVersion
End Function

'---    GetOSArchitecture
'---    Returns OS Architecture
'---    Example: "64"
Function GetOSArchitecture
    On Error Resume Next
	Dim cItems, oItem
	Set cItems = oWMI.ExecQuery("Select * from Win32_OperatingSystem")
	For Each oItem In cItems
		If UCase(oItem.OSArchitecture) = "32-BIT" Then
			GetOSArchitecture = "32"
		ElseIf UCase(oItem.OSArchitecture) = "64-BIT" Then
			GetOSArchitecture = "64"
		End If
	Next
	'---    If unable to acquire architecture from WMI
	Dim sOS
	If GetOSArchitecture <> "64-bit" and GetOSArchitecture <> "32-bit" Then
        sOS = oShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
        If sOS = "x86" Then
            GetOSArchitecture = "32"
        ElseIf sOS = "AMD64" Then
            GetOSArchitecture = "64"
        End IF
	End If
	WriteToLog "** GetOSArchitecture Returned: " & GetOSArchitecture
	Err.Clear
End Function

'---    LogSoftware
'---    Logs Installed software
'---    Example: "64"
Sub LogSoftware
	WriteToLog "** LogSoftware"
	Dim cItems, oItem
	On Error Resume Next
    Set cItems = oWMI.ExecQuery("Select * from Win32_Product")
    If Err.Number = 0 Then
		For Each oItem In cItems
			WriteToLog "  """ & oItem.Name & """ Version """ & oItem.Version & """"
		Next
    End If
    Err.Clear      
End Sub

'-------------------------------------------------------------------------------
'---    Hardware
'-------------------------------------------------------------------------------

'--- IsLANConnected
'--- Is Local Area Network Cable plugged in
'--- Returns boolean

Function IsLANConnected
Dim aStatusStr, strComputer, oWMI, cItems, oItem
aStatusStr = split("0,1,Connected,3,4,5,Disconnected,?", ",")

strComputer = "."
Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set cItems = oWMI.ExecQuery("Select * from Win32_NetworkAdapter where AdapterTypeID = 0")

For Each oItem in cItems
  If Instr(oItem.NetConnectionID, "Local Area Connection") > 0 Or Instr(oItem.NetConnectionID, "Ethernet") > 0 Then
    If aStatusStr(oItem.NetConnectionStatus) = "Connected" Then
    	IsLANConnected = True
    	WriteToLog "** Local Area Connection status: " & aStatusStr(oItem.NetConnectionStatus) & ""
    Else
    	IsLANConnected = False
    	WriteToLog "** Local Area Connection status: " & aStatusStr(oItem.NetConnectionStatus) & ""
    End If		
  End if
Next
End Function


'-------------------------------------------------------------------------------
'---    Drivers
'-------------------------------------------------------------------------------

Sub ImportDrivers(sPath)
	'Detect if a folder or individual file was passed
	If Right(UCase(sPath),3) = "INF" Then
		ImportDriverINF sPath
	Else
		ImportDriversRecursion sPath
	End If
End Sub

Function ImportDriverINF(sPath)
	Dim sCMD, iRet
	sCMD = "pnputil.exe -i -a """ & sPath & """"
	WriteToLog "  ** Importing Driver: """ & sPath & """"
	WriteToLog "    About to Run Command: " & sCMD
	ImportDriverINF = oShell.Run(sCMD,0,True)
	If iRet = 0 Then
		WriteToLog "    SUCCESS: Driver Imported successfully"
	Else
		WriteToLog "    ERROR: Driver Import failed"
	End If
End Function

Sub ImportDriversRecursion(sPath)
	'WScript.Echo sPath
	Dim oFolder,oFile,oSubFolder
	
	Set oFolder = oFSO.GetFolder(sPath)
	
	For Each oFile In oFolder.Files
		If Right(UCase(oFile.Path),3) = "INF" Then
			ImportDriverINF oFile.Path
		End If
	Next
	
	
	For Each oSubFolder In oFolder.SubFolders
		ImportDriversRecursion oSubFolder.Path
	Next
End Sub


'-------------------------------------------------------------------------------
'---    Processes
'-------------------------------------------------------------------------------

'---    IsProcessRunning
'---    Parameter1: String - Process to search for
'---    Returns boolean
'---    Example: "True"
Function IsProcessRunning(sSearch)
	Dim cItems, oItem
	Set cItems = oWMI.ExecQuery("Select * from Win32_Process where Name like '" & sSearch & "'")
	If cItems.Count = 0 Then
		IsProcessRunning = False
	Else 
		IsProcessRunning = True
	End If
	WriteToLog "** IsProcessRunning '" & sSearch & "' Returned: " & IsProcessRunning
End Function

'---    CheckProcesses
'---    Parameter1: String - Processes to search for
Sub CheckProcesses(sList)
	WriteToLog "** CheckProcesses (" &  sList & ")"
	Dim aList, i, bCheck
	bCheck = False
	aList = Split(sList,",")
	If UBound(aList) <> 0 Then
		For i = 0 To UBound(aList)
			Dim sProcess: sProcess = aList(i)
			If IsProcessRunning(sProcess) Then
				sProcess = Replace(UCase(sProcess),"%","")
				WriteToLog "  PLEASE CLOSE """ & sProcess & """ AND RERUN SCRIPT"
				bCheck = True
			End If
		Next
	End If
	
	If bCheck = True Then
		WriteToLog "  ERROR: Close open applications and rerun installation"
		LogEnd
		WScript.Quit(10)
	End If
End Sub

'---    IsAppRunning
'---    Parameter1: String - Application Name to search for
'---    Returns boolean
'---    Example: "True"
Function IsAppRunning(sSearch)
	Dim oWord, cTasks, oTask
	Set oWord = CreateObject("Word.Application")
	Set cTasks = oWord.Tasks
	IsAppRunning = False
	For Each oTask in cTasks
		If oTask.Visible And InStr(1,oTask.Name,sSearch,1) Then
			IsAppRunning = True
		End If
	Next
	oWord.Quit
	WriteToLog "** IsAppRunning '" & sSearch & "' Returned: " & IsAppRunning
End Function

'---    KillProcess
'---    Kills processes running with defined name
'---    Parameter1: String - Process to search for
Sub KillProcess(sSearch)
	WriteToLog "** KillProcess: '" & sSearch & "'"
	Dim cItems, oItem
	Set cItems = oWMI.ExecQuery("Select * from Win32_Process where Name like '" & sSearch & "'")
	For Each oItem In cItems
		oItem.Terminate
		WriteToLog "  " & oItem.Name & " (" & oItem.Handle & ")"
	Next
End Sub	

'-------------------------------------------------------------------------------
'---    Software
'-------------------------------------------------------------------------------

'---    RunInstall
'---    Runs installation command
'---    Parameter1: String - Command to run
'---    Returns Integer - error number from uninstall 

Function RunInstall(sCMD,bQuit)
	WriteToLog "** RunInstall"
        NoZoneChecks
	Dim iRet
	WriteToLog "  Run command: " & sCMD
	iRet = oShell.Run(sCMD,0,True)
	If iRet = 0 Or iRet = 3010 Then
		WriteToLog "  SUCCESS: Command Completed Successfully With Return Code (" & iRet & ")"
	Else
		WriteToLog "  ERROR: Command Completed with Return Code (" & iRet & ")"
		If bQuit = True Then
			RunBeforeQuit(iRet)
			LogEnd
			WScript.Quit(iRet)
		End If
	End If
	RunInstall = iRet
	
End Function

'---    RunUninstall
'---    Searches for software to uninstall
'---    Parameter1: String - Product to search for
'---    Parameter2: String - Parameters needed for silent install
'---    Returns Integer - error number from uninstall 
Function RunUninstall(sProduct, sVersion)
	WriteToLog "** RunUninstall '" & sProduct & "'"
	If sVersion = "" Then
	   sVersion = "%"
	End If
	Dim cItems, oItem
	Dim iRet, sName, sVendor, sIdentifyingNumber
	RunUninstall = 0
    On Error Resume Next
    Set cItems = oWMI.ExecQuery("Select * from Win32_Product Where Name like '" & sProduct & "' And Version like '" & sVersion & "'")
    If Err.Number = 0 Then
    	For Each oItem In cItems
    		On Error Resume Next
    		WriteToLog "  """ & oItem.Name & """ Version """ & oItem.Version & """"
    		sName = oItem.Name
    		sVendor = oItem.Vendor
    		sIdentifyingNumber = oItem.IdentifyingNumber
    		If sIdentifyingNumber <> "" Then
    			                WriteToLog "  Uninstalling: " & sVendor & " " & sName
                sStart = CDate(Now)
                iRet = oItem.Uninstall
                sStop = CDate(Now)

                WriteToLog "  RunTime: " & DateDiff("n",sStart,sStop)
                WriteToLog "  Return: " & iRet
                If iRet = 0 Or iRet = 3010 Then
                	'Success
                Else
                	RunUninstall = iRet
                End If
    		End If
    		Err.Clear
		Next
    Else
    	WriteToLog "  ERROR WHEN DETECTING SOFTWARE INSTALLATION"
    End If
    Err.Clear
End Function

Function RunInstallDontLog(sCMD)
	WriteToLog "  ** RunInstall"
        NoZoneChecks
	Dim iRet
	WriteToLog "  About to run command: ** REDACTED **"
	iRet = oShell.Run(sCMD,0,True)
	If iRet = 0 Then
		WriteToLog "  SUCCESS: Command Completed Successfully"
	Else
		WriteToLog "  ERROR: Command Completed with Error (" & iRet & ")"
	End If
	RunInstallDontLog = iRet
End Function

'---    IsSoftwareInstalled
'---    Searches WMI for record of software install
'---    Parameter1: String - Product to search for
'---    Parameter1: String - Version to search for
'---    Returns Boolean
Function IsSoftwareInstalled(sProduct, sVersion)
	WriteToLog "** IsSoftwareInstalled '" & sProduct & "' Version " & sVersion
	IsSoftwareInstalled = False
	Dim cItems, oItem
	If sVersion = "" Then
	   sVersion = "%"
	End If
    On Error Resume Next
    Set cItems = oWMI.ExecQuery("Select * from Win32_Product Where Name like '" & sProduct & "' And Version like '" & sVersion & "'")
    If Err.Number = 0 Then
    	If cItems.Count = 0 Then
            IsSoftwareInstalled = False
        Else
    		For Each oItem In cItems
    			WriteToLog "  """ & oItem.Name & """ Version """ & oItem.Version & """"
    		Next
            IsSoftwareInstalled = True
        End If
    Else
    	WriteToLog "  ERROR WHEN DETECTING SOFTWARE INSTALLATION"
    End If
    WriteToLog "** IsSoftwareInstalled - Returned: " & IsSoftwareInstalled
End Function

'---    GetCurrentUser
'---    Checks specified PC for current logged in user
'---    use "." to check the local machine
'---    Parameter1: String - name of PC that will be queried for current user logged in        
Function GetCurrentUser(sComputer)
    Dim oRemoteWMIService, colUsers, oItem, sUser
    On Error Resume Next
    Set oRemoteWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & sComputer & "")
    Set colUsers = oRemoteWMIService.ExecQuery("Select * from Win32_ComputerSystem")
	If Err.Number <> 0 Then
	    GetCurrentUser = "error:  " & Err.Description
    Else    
        GetCurrentUser = ""
        sUser = ""
        For Each oItem In colUsers  
            If oItem.UserName <> "" then
                GetCurrentUser = oItem.UserName
	        ElseIf oItem.Username <> "" And oItem.Username <> Null Then
		        GetCurrentUser = oItem.UserName
		    End If
        Next
    End If
    Err.Clear
End Function

'---	NoZoneChecks
'---	Disable zone check to allow cmd and msi to run without user prompt
Sub NoZoneChecks
    Dim oENV, iRet
    Set oEnv = oShell.Environment("PROCESS")
    oEnv("SEE_MASK_NOZONECHECKS") = 1
End Sub

'---	CommonMSIErrors
'---	Writes MSI Error information to logs
Sub CommonMSIErrors(iErr)
	If iErr = 1601 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") The Windows Installer service could not be accessed. Contact your support personnel to verify that the Windows Installer service is properly registered."
	ElseIf iErr = 1602 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") User cancel installation."
	ElseIf iErr = 1603 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Fatal error during installation."
	ElseIf iErr = 1604 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Installation suspended, incomplete."
	ElseIf iErr = 1605 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This action is only valid for products that are currently installed."
	ElseIf iErr = 1606 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Feature ID not registered."
	ElseIf iErr = 1607 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Component ID not registered."
	ElseIf iErr = 1608 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Unknown property."
	ElseIf iErr = 1609 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Handle is in an invalid state."
	ElseIf iErr = 1610 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") The configuration data for this product is corrupt. Contact your support personnel."
	ElseIf iErr = 1611 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Component qualifier not present."
	ElseIf iErr = 1612 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") The installation source for this product is not available. Verify that the source exists and that you can access it."
	ElseIf iErr = 1613 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This installation package cannot be  installed by the Windows Installer service. You must install a Windows service pack that contains a newer version of the Windows Installer service."
	ElseIf iErr = 1614 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Product is uninstalled."
	ElseIf iErr = 1615 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") SQL query syntax invalid or unsupported."
	ElseIf iErr = 1616 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Record field does not exist."
	ElseIf iErr = 1618 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Another installation is already in progress. Complete that installation before proceeding with this install."
	ElseIf iErr = 1619 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This installation package could not be opened. Verify that the package exists and that you can access it, or contact the application vendor to verify that this is a valid Windows Installer package."
	ElseIf iErr = 1620 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This installation package could not be  opened. Contact the application vendor to verify that this is a valid Windows Installer package."
	ElseIf iErr = 1621  Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") There was an error starting the Windows Installer service user interface. Contact your support personnel."
	ElseIf iErr = 1622 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Error opening installation log file. Verify that the specified log file location exists and is writable."
	ElseIf iErr = 1623 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This language of this installation package is not supported by your system."
	ElseIf iErr = 1625 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This installation is forbidden by system policy. Contact your system administrator."
	ElseIf iErr = 1626 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Function could not be executed."
	ElseIf iErr = 1627 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Function failed during execution."
	ElseIf iErr = 1628 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Invalid or unknown table specified."
	ElseIf iErr = 1629 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Data supplied is of wrong type."
	ElseIf iErr = 1630 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Data of this type is not supported."
	ElseIf iErr = 1631 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") The Windows Installer service failed to start. Contact your support personnel."
	ElseIf iErr = 1632 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") The temp folder is either full or inaccessible. Verify that the temp folder exists and that you can write to it."
	ElseIf iErr = 1633 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This installation package is not supported on this platform. Contact your application vendor."
	ElseIf iErr = 1634 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Component not used on this machine."
	ElseIf iErr = 1624 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Error applying transforms. Verify that the specified transform paths are valid."
	ElseIf iErr = 1635 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This patch package could not be opened. Verify that the patch package exists and that you can access it, or contact the application vendor to verify that this is a valid Windows Installer patch package."
	ElseIf iErr = 1636 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This patch package could not be opened. Contact the application vendor to verify that this is a valid Windows Installer patch package."
	ElseIf iErr = 1637 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") This patch package cannot be processed by the Windows Installer service. You must install a Windows service pack that contains a newer version of the Windows Installer service."
	ElseIf iErr = 1638 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Another version of this product is already installed. Installation of this version cannot continue. To configure or remove the existing version of this product, use Add/Remove Programs on the Control Panel."
	ElseIf iErr = 1639 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") Invalid command line argument. Consult the Windows Installer SDK for detailed command line help."
	ElseIf iErr = 3010 Then
		WriteToLog "  ERROR DETAILS (" & iErr & ") A restart is required to complete the install. This does not include installs where the ForceReboot action is run. Note that this error will not be available until future version of the installer. "
	Else
		WriteToLog "  ERROR DETAILS (" & iErr & ") Unknown Error"
	End If
End Sub
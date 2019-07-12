' // ***************************************************************************
' //
' // File:      Pause_TS.vbs
' // 
' // Version:   1
' // 
' // Purpose:   
' // 
' // Usage:     cscript Pause_TS.vbs
' //            	/Uninstall	(Performs Uninstall)
' //
' // Created:   1.1 Brandon Hilgeman (9/11/2014)
' //
' // Modified:  
' // 
' // ***************************************************************************

'Hide the Task Sequence Progress window

Set TsProgressUI = CreateObject("Microsoft.SMS.TsProgressUI")
TsProgressUI.CloseProgressDialog

Dim wshShell, btn
Set wshShell = WScript.CreateObject("WScript.Shell")

' Call the Popup method with a 7 second timeout.
btn = wshShell.Popup("The Task Sequnce has been Paused", ,"## Task Sequence Paused ##", &H1)

Select Case btn
    ' OK button pressed.
    case 1
        WScript.Quit(0)
    ' Cancel button pressed.
    case 2
        WScript.Quit(1)
    
End Select

wscript.quit(0)
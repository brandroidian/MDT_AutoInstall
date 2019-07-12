<# 
        Script Library
  
        File:       ScriptLibrary1.11.psm1
    
        Purpose:    Contains common functions and routines useful in other scripts

        Author: Brandon Hilgeman 
                brandon.hilgeman@gmail.com

#>

Function Set-GlobalVariables{
    $Global:sComputerName = $env:computername
    $Global:sWindowsPath = $env:windir+"\"
    $Global:sDefaultLogPath = $sWindowsPath+"_BPH\"
    $Global:sDefaultLog = $sDefaultLogPath+$sScriptName+".log"
    $Global:MaxLogSizeInKB = 1024*20
    $Global:sTempPath = $env:TEMP+"\"
    $Global:ScriptStatus = 'Success'
    Detect-Runtime
}

Function Detect-Runtime{
    
    $IsInTS = Is-InTS

    If($IsInTS -eq $True){

        If($oENV.Value("DEPLOYMENTMETHOD").ToUpper() -eq "SCCM"){
                $Global:sRuntime = "SCCM"
        }
        ElseIf($oENV.Value("DEPLOYMENTMETHOD").ToUpper() -eq "MDT" -or $oENV.Value("DEPLOYMENTMETHOD").ToUpper() -eq "UNC" -or $oENV.Value("DEPLOYMENTMETHOD").ToUpper() -eq "MEDIA"){
            $Global:sRuntime = "MDT"
        }
        Else{
            $Global:sRuntime = "UNKNOWN"
        }
    }
    Else{
        $Global:sRuntime = "STANDALONE"
    }

    If($sRuntime -eq "SCCM"){
        $Global:sLog = $oENV.Value("LogPath")+"\"+$sScriptName+".log"
        $Global:sLogPath = $oENV.Value("LogPath")+"\"
        $sDeployRoot = $oENV.Value("DeployRoot")+"\"
        $sScriptRoot = $oENV.Value("ScriptRoot")+"\"
    }
    ElseIf($sRuntime -eq "MDT"){
        $Global:sLog = $oENV.Value("LogPath")+"\"+$sScriptName+".log"
        $Global:sLogPath = $oENV.Value("LogPath")+"\"
        $sDeployRoot = $oENV.Value("DeployRoot")+"\"
        $sScriptRoot = $oENV.Value("ScriptRoot")+"\"
    }
    ElseIf($sRuntime -eq "STANDALONE"){
        $Global:sLog = $sDefaultLog
        $Global:sLogPath = $sDefaultLogPath
    }
    Else{
        $Global:sLog = $sTempPath+$sScriptName+".log"
        $Global:sLogPath = $sTempPath
    }

    If(($IsInTS -eq $False) -and ($sWindowsPath.substring(0,3) -eq "X:\")){
        $Global:sRuntime = "MDT"
        $Global:sLog = "X:\MININT\SMSOSD\OSDLOGS\$sScriptName.log"
        $Global:sLogPath = "X:\MININT\SMSOSD\OSDLOGS\"
        $Global:sBackupLog = $sTempPath+$sScriptName+".log"
    }

}

Function Write-Log {
    Param (
        [Parameter(Mandatory=$true)]
        $Message,
        [Parameter(Mandatory=$false)]
        $ErrorMessage,
        [Parameter(Mandatory=$false)]
        $Component,
        [Parameter(Mandatory=$false)]
        [int]$Type,
        [Parameter(Mandatory=$false)]
        $LogFile
    )

    $Time = Get-Date -Format "HH:mm:ss.ffffff"
    $Date = Get-Date -Format "MM-dd-yyyy"
 
    If($ErrorMessage -ne $null){
        $Type = 3
    }
    If($Component -eq $null){
        $Component = $sScriptName
    }
    If($Type -eq $null){
        $Type = 1
    }
    If($LogFile -eq $null){
        $LogFile = $sLog
    }

    If(!(Test-Path -Path $sLogPath)){
        MkDir $sLogPath | Out-Null
    }
 
    $LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"
    $LogMessage | Out-File -Append -Encoding UTF8 -FilePath $sLog

    If ((Get-Item $sLog).Length/1KB -gt $MaxLogSizeInKB){
        $log = $sLog
        Remove-Item ($log.Replace(".log", ".lo_"))
        Rename-Item $sLog($log.Replace(".log", ".lo_")) -Force
    }
}

Function Start-Log {
    $StartDate = ((get-date).toShortDateString() + " " + (get-date).toShortTimeString())
    $global:sScriptStart = get-date
    Write-Log -Message ("-" * 10 + "  Start Script: $sScriptName " + $StartDate + " " + "-" * 10)
}

Function End-Log {
    $EndDate = ((get-date).toShortDateString() + " " + (get-date).toShortTimeString())
    $sScriptEnd = get-date
    $RunTime = ($sScriptEnd.Minute - $sScriptStart.Minute)
    Write-Log -Message ("-" * 10 + "  End Script: " + $EndDate + "   RunTime: " + $RunTime + " Minute(s) " + "-" * 10)
}

Function Scan-Args {

}

Function Set-Mode {

    If($sArgs -eq "Uninstall"){
        Start-Uninstall
    }
    Else{
        Start-Install
    }
}






<#'-------------------------------------------------------------------------------
  '---    GUI
  '-------------------------------------------------------------------------------#>


Function MsgBox {
    <#
    Buttons = AbortRetryIgnore, OK, OKCancel, RetryCancel, YesNo, YesNoCancel
    Icons = Asterisk, Error, Exclamation, Hand, Information, None, Question, Stop, Warning
    #>
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Message,
        [Parameter(Mandatory=$False)]
        [String]$Title,
        [Parameter(Mandatory=$False)]
        $Buttons,
        [Parameter(Mandatory=$False)]
        [String]$Icon
    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    If($Buttons -eq $null){
        $Buttons = 0
    }
    If($Icon -eq ""){
        $Icon = 'Exclamation'
    }

    Write-Log -Message ("  MsgBox: ""$Message"", ""$Title"", ""$Buttons""") -Type 1
    
    $MsgBox = [System.Windows.Forms.MessageBox]::Show("$Message", "$Title", $Buttons, $Icon)

    Write-Log -Message "  MsgBox: $MsgBox pressed" -Type 1

    Return $MsgBox

}


Function Get-XAML{
    Param(
        [Parameter(Mandatory=$True)]
        $sPath,
        [Parameter(Mandatory=$False)]
        $bVariables
    )

    Write-Log -Message "  Loading XAML code from: $sPath"

    Try{
        $InputXML = Get-Content $sPath
    }
    Catch{
        Write-Log -Message "  Unable to locate XAML file at: $sPath"
    }
    $InputXML = $InputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $InputXML
    #Read XAML
 
    $Reader=(New-Object System.Xml.XmlNodeReader $XAML)
    Try{
        $Global:Form=[Windows.Markup.XamlReader]::Load( $Reader )
    }
    Catch{
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
    }
 
    $XAML.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Global:Form.FindName($_.Name) -Scope Global}
    $WPFimage_logo.Source = "$PSScriptRoot\Images\Logo.bmp"
     
    If($bVariables){
        Get-FormVariables
    }

    Start-WPFApp
}


Function Get-FormVariables{
    If ($Global:ReadmeDisplay -ne $True){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
    Write-Host "Found the following interactable elements from our form" -ForegroundColor Cyan
    Get-Variable WPF*
}


Function Show-GUI{
    Param(
        [Parameter(Mandatory=$False)]
        $WPFVariable
    )

    Try{
        $WPFVariable.Visibility = "Visible"
    }
    Catch{
        Write-Log -Message "  Unable to show: $WPFVariable"
    }
}


Function Hide-GUI{
    Param(
        [Parameter(Mandatory=$False)]
        $WPFVariable
    )

    Try{
        $WPFVariable.Visibility = "Hidden"
    }
    Catch{
        Write-Log -Message "  Unable to hide: $WPFVariable"
    }
}


Function Enable-GUI{
    Param(
        [Parameter(Mandatory=$False)]
        $WPFVariable
    )

    Try{
        $WPFVariable.IsEnabled = $True
    }
    Catch{
        Write-Log -Message "  Unable to Enable: $WPFVariable"
    }
}


Function Disable-GUI{
    Param(
        [Parameter(Mandatory=$False)]
        $WPFVariable
    )

    Try{
        $WPFVariable.IsEnabled = $False
    }
    Catch{
        Write-Log -Message "  Unable to Disable: $WPFVariable"
    }
}


Function Clear-GUI{
    Param(
        [Parameter(Mandatory=$False)]
        $WPFVariable,
        [Parameter(Mandatory=$False)]
        $Type
    )

    If($Type -eq "text"){
        Try{$WPFVariable.Clear()}
        Catch{}
    }
    ElseIf($Type -eq "combo"){
        Try{
            $WPFVariable.Items.Clear()
        }
        Catch{}
    }
}

Function Add-GUIText{
    Param(
        [Parameter(Mandatory=$False)]
        $WPFVariable,
        [Parameter(Mandatory=$False)]
        $Text
    )
    Try{
        $WPFVariable.AddText($Text)
    }
    Catch{
        Write-Log "  Failed to add text to: $WPFVariable"
    }
}


Function Show-BalloonTip{          
    Param(
        [Parameter(Mandatory=$True)]
        [string]$Title,
        [Parameter(Mandatory=$False)]
        [string]$Type,
        [Parameter(Mandatory=$True)]
        [string]$Message,
        [Parameter(Mandatory=$False)]
        [Int]$Duration
    )

    If($Type -eq ""){
        $Type = "Info"
    }
    If($Duration -eq $Null){
        $Duration = 5000
    }

    [system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null

    $Balloon = New-Object System.Windows.Forms.NotifyIcon
    $Path = Get-Process -id $pid | Select-Object -ExpandProperty Path
    $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Path)
    $Balloon.Icon = $Icon
    $Balloon.BalloonTipIcon = $Type
    $Balloon.BalloonTipText = $Message
    $Balloon.BalloonTipTitle = $Title
    $Balloon.Visible = $True
    $Balloon.ShowBalloonTip($Duration)
}



<#'-------------------------------------------------------------------------------
  '---    Systems
  '-------------------------------------------------------------------------------#>


Function Test-Ping{
    Param(
        [Parameter(Mandatory=$True)]
        $HostName
    )
    Write-Log -Message "  Test-Ping $HostName"

    $TestPing = Test-Connection -ComputerName $HostName -ErrorAction SilentlyContinue -ErrorVariable iErr;
    If($iErr){
        Write-Log -Message "  Error: Unable to Ping: $HostName"
        $TestPing = $False
    }
    Else{
        Write-Log "  Success: Pinged $HostName"
        $TestPing = $True
        Return $TestPing
    }
}

Function Get-ComputerName {
    $GetComputerName = $env:computername

    Write-Log -Message ("  Get-ComputerName: $GetComputerName")

    Return $GetComputerName
}

Function Get-Make{

    Try{
        $GetMake = (Get-WmiObject -Class Win32_ComputerSystem).Manufacturer
    }
    Catch{
        Write-Log "  Unable to find Manufacturer"
    }

    Write-Log "  Get-Make: '$GetMake'"
    Return $GetMake
}



Function Get-Model{

    Try{
        $GetModel = (Get-WmiObject -Class Win32_ComputerSystem).Model
    }
    Catch{
        Write-Log "  Unable to find Model"
    }

    Write-Log "  Get-Model: '$GetModel'"
    Return $GetModel
}

Function Get-CurrentUser{

    Try{
        $GetUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    Catch{
        Write-Log "  Unable to find Logged on user"
    }

    Write-Log "  Get-CurrentUser: '$GetUser'"
    Return $GetUser
}




<#'-------------------------------------------------------------------------------
  '---    Processes
  '-------------------------------------------------------------------------------#>

Function Is-ProcessRunning{

    Param(
        [Parameter(Mandatory=$True)]
        $sProcess
    )
    
    If($sProcess -contains '"'){
        $sProcess = $sProcess -replace '"',''
    }

    $IsProcessRunning = Get-Process $sProcess -ErrorAction SilentlyContinue

    If ($IsProcessRunning){
        $RunningProcess = $True
    }
    Else{
        $RunningProcess = $False
    }
    Write-Log "  Is-ProcessRunning: '$sProcess' Returned: $RunningProcess"
    Return $RunningProcess 

}


Function End-Process{

    Param(
        [Parameter(Mandatory=$True)]
        $sProcess
    )

    Write-Log "  End-Process($sProcess)"
    
    If($sProcess -contains '"'){
        $sProcess = $sProcess -replace '"',''
    }

    $RunningProcess = Get-Process $sProcess -ErrorAction SilentlyContinue

    If(Is-ProcessRunning -sProcess $sProcess){
        $RunningProcess.CloseMainWindow()
        Sleep 3
    }
    Else{
        Write-Log "  Process: '$sProcess' is not running"
        Return
    }
    
    If (!$RunningProcess.HasExited){
        $RunningProcess | Stop-Process -Force
        Sleep 3
    }
    
    If(Is-ProcessRunning -sProcess $sProcess){
        Write-Log "  Process: '$sProcess' UnSuccessfully terminated"
    }
    Else{
        Write-Log "  Process: '$sProcess' Successfully terminated"
    }

}


<#'-------------------------------------------------------------------------------
  '---    Software
  '-------------------------------------------------------------------------------#>

Function Run-Install {
    Param(
        [String]$sCMD,
        [String]$sArg
     )
    Write-Log -Message "  RunInstall"
    $ENV:SEE_MASK_NOZONECHECKS = 1
    Write-Log -Message ("  Running: $sCMD $sArg") -type 1
    $RunInstall = (Start-Process $sCMD -ArgumentList $sArg -Wait -PassThru -ErrorAction SilentlyContinue -ErrorVariable iErr).ExitCode

    If($RunInstall -ne 0) {
        Write-Log -Message "  ERROR: Command Completed with Error: $RunInstall" -Type 3
    }
    Else{
        Write-Log -Message "  SUCCESS: Command Completed Successfully"
    }
    $ENV:SEE_MASK_NOZONECHECKS = 0
    Return $RunInstall
}

 <#Function Run-Uninstall {
    Param(
        [Parameter(Mandatory=$True)]
        $sName,
        [Parameter(Mandatory=$True)]
        $sVersion
    )
  
    Write-Log -Message "  RunUninstall: ""$sName"", ""$sVersion"""
    $cItems = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '$sName' And Version Like '$sVersion'"
    ForEach($oItem in $cItems){
        Write-Log -Message ("'$oItem.Name' '$oItem.Version'")
        $iRet = $oItem.Uninstall()
        Write-Log "  $iRet"
    }
 }#>

 Function Run-UnInstall {
    Param(
        [String]$sCMD,
        [String]$sArg
     )
    Write-Log -Message "  RunUnInstall"
    $ENV:SEE_MASK_NOZONECHECKS = 1
    Write-Log -Message ("  Running: $sCMD $sArg") -type 1
    $RunInstall = (Start-Process $sCMD -ArgumentList $sArg -Wait -PassThru -ErrorAction SilentlyContinue -ErrorVariable iErr).ExitCode

    If($RunInstall -ne 0) {
        Write-Log -Message "  ERROR: Command Completed with Error: $RunInstall" -Type 3
    }
    Else{
        Write-Log -Message "  SUCCESS: Command Completed Successfully"
    }
    $ENV:SEE_MASK_NOZONECHECKS = 0
    Return $RunInstall
}



 Function Log-InstalledApps {
    Write-Log "  ** Log-InstalledApps **"
    $UninstallRegKeys=@("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")           
    ForEach($Computer in $sComputerName){
        If(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {
            ForEach($UninstallRegKey in $UninstallRegKeys){
                Try {            
                    $HKLM   = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computer)
                    $UninstallRef  = $HKLM.OpenSubKey($UninstallRegKey)
                    $Applications = $UninstallRef.GetSubKeyNames()
                }
                Catch {            
                    Write-Log -Message "  Error: Failed to read $UninstallRegKey"          
                    Continue
                }
            }
        }
    }
            
    ForEach ($App in $Applications) {
        $AppRegistryKey  = $UninstallRegKey + "\\" + $App
        $AppDetails   = $HKLM.OpenSubKey($AppRegistryKey)
        $AppGUID   = $App
        $AppDisplayName  = $($AppDetails.GetValue("DisplayName"))
        $AppVersion   = $($AppDetails.GetValue("DisplayVersion"))
        $AppPublisher  = $($AppDetails.GetValue("Publisher"))
        $AppInstalledDate = $($AppDetails.GetValue("InstallDate"))
        $AppUninstall  = $($AppDetails.GetValue("UninstallString"))
           If($UninstallRegKey -match "Wow6432Node"){
            $Softwarearchitecture = "x86"
           }
           Else {
                $Softwarearchitecture = "x64"
           }
           If(!$AppDisplayName){
            continue
           }
           Write-Log -Message "  ""$AppDisplayName""  Version: ""$AppVersion"""
    }
}

Function IsSoftwareInstalled {
    Param(
        [Parameter(Mandatory=$True)]
        $sProduct,
        [Parameter(Mandatory=$True)]
        $sVersion
    )
    $IsSoftwareInstalled = $False

    Write-Log -Message "  IsSoftwareInstalled: ""$sProduct"" ""$sVersion"""

    $UninstallRegKeys=@("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")           
    ForEach($Computer in $sComputerName) {
        If(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {
            ForEach($UninstallRegKey in $UninstallRegKeys) {
                Try {            
                    $HKLM   = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computer)            
                    $UninstallRef  = $HKLM.OpenSubKey($UninstallRegKey)            
                    $Applications = $UninstallRef.GetSubKeyNames()            
                }
                Catch {            
                    Write-Log -Message "  Failed to read $UninstallRegKey" -Type 3
                    Continue
                }
            }
        }
    }
            
    ForEach ($App in $Applications){
        $AppRegistryKey  = $UninstallRegKey + "\\" + $App
        $AppDetails   = $HKLM.OpenSubKey($AppRegistryKey)
        $AppGUID   = $App
        $AppDisplayName  = $($AppDetails.GetValue("DisplayName"))
        $AppVersion   = $($AppDetails.GetValue("DisplayVersion"))
        $AppPublisher  = $($AppDetails.GetValue("Publisher"))
        $AppInstalledDate = $($AppDetails.GetValue("InstallDate"))
        $AppUninstall  = $($AppDetails.GetValue("UninstallString"))
        If($UninstallRegKey -match "Wow6432Node"){
            $Softwarearchitecture = "x86"
        }
        Else{
            $Softwarearchitecture = "x64"
        }
        If(!$AppDisplayName){
            continue
        }
        If(($AppDisplayName -like $sProduct) -and ($AppVersion -like $sVersion)){
            $IsSoftwareInstalled = $True
            Write-Log -Message "  ""$AppDisplayName"" ""$AppVersion"""
        }
    }

    Write-Log -Message "  IsSoftwareInstalled - Returned: $IsSoftwareInstalled"
    Return $IsSoftwareInstalled
}

 <#'-------------------------------------------------------------------------------
   '---    File System
   '-------------------------------------------------------------------------------#>

Function Create-Folder {
    Param(
        [Parameter(Mandatory=$True)]
        $sPath
        )
        
    Write-Log -Message "  Create-Folder($sPath)"
    If(!(Test-Path -Path $sPath)){
        $CreateFolder = New-Item -ItemType directory -Path $sPath -Force -ErrorAction SilentlyContinue -ErrorVariable iErr;
        If($iErr){
            Write-Log -Message "  Error: Could not create directory: ""$sPath""" -Type 3
        }
        Else{
            Write-Log -Message "  Success: Created directory: ""$sPath"""
        }
    }
    Else{
        Write-Log -Message "  ""$sPath"" already exists"
        }
}

Function Copy-File {
    Param(
        [Parameter(Mandatory=$True)]
        [String]$sSource,
        [Parameter(Mandatory=$True)]
        [String]$sDestination
    )
    
    Write-Log -Message "** CopyFile($sSource, $sDestination)"
    Write-Log -Message "  Source: $sSource"
    Write-Log -Message "  Destination: $sDestination"

    If(!(Test-Path -Path $sDestination)){
        New-Item -ItemType File -Path $sDestination -Force
    }

    If(!(Test-Path -Path $sSource)){
        Write-Log -Message "  Copy Completed: Unsucessful (Invalid Source Path)" -Type 3
    }
    Else{
        $CopyFile = Copy-Item -Path $sSource -Destination $sDestination -Force -ErrorAction SilentlyContinue -ErrorVariable iErr;
        If($iErr){
            Write-Log -Message "  Copy Completed: Unsucessful" -Type 3
        }
        Else {
            Write-Log -Message "  Copy Completed: Success" -Type 1
        }
    }
}

Function Copy-Folder{
    Param(
        [Parameter(Mandatory=$True)]
        [String]$sSource,
        [Parameter(Mandatory=$True)]
        [String]$sDestination
    )
    
    Write-Log -Message "** CopyFolder($sSource, $sDestination)"
    Write-Log -Message "  Source: $sSource"
    Write-Log -Message "  Destination: $sDestination"

    If(!(Test-Path -Path $sDestination)){
        CreateFolder $sDestination
    }
    
    If(!(Test-Path -Path $sSource)){
        Write-Log -Message "  Copy Completed: Unsucessful (Invalid Source Path)" -Type 3
    }
    Else{
        $CopyFolder = Copy-Item -Path $sSource -Destination $sDestination -Force -Recurse -ErrorAction SilentlyContinue -ErrorVariable iErr;
        If ($iErr){
            Write-Log -Message "  Copy Completed: Unsucessful" -Type 3
        }
        Else{
            Write-Log -Message "  Copy Completed: Success" -Type 1
        }
    }
}

Function Delete-Object {
    Param(
        [Parameter(Mandatory=$True)]
        $sPath
    )

    Write-Log -Message "** DeletObject($sPath)"

    If(!(Test-Path -Path $sPath)){
        Write-Log -Message "  Object: $sPath does not exist" -Type 3
    }
    Else{
        $DeleteObject = Remove-Item $sPath -Force -Recurse -ErrorAction SilentlyContinue -ErrorVariable iErr;
        If($iErr){
            Write-Log -Message "  Error: Object Deletion Failed" -Type 3
        }
        Else{
            Write-Log -Message "  Success: Object Deletion Completed Successfully"
        }
    }
}


Function Get-IniContent {  
    <#  
    .Synopsis  
        Gets the content of an INI file  
          
    .Description  
        Gets the content of an INI file and returns it as a hashtable  
          
        #Requires -Version 2.0  
          
    .Inputs  
        System.String  
          
    .Outputs  
        System.Collections.Hashtable  
          
    .Parameter FilePath  
        Specifies the path to the input file.  
          
    .Example  
        $FileContent = Get-IniContent "C:\myinifile.ini"  
        -----------  
        Description  
        Saves the content of the c:\myinifile.ini in a hashtable called $FileContent  
      
    .Example  
        $inifilepath | $FileContent = Get-IniContent  
        -----------  
        Description  
        Gets the content of the ini file passed through the pipe into a hashtable called $FileContent  
      
    .Example  
        C:\PS>$FileContent = Get-IniContent "c:\settings.ini"  
        C:\PS>$FileContent["Section"]["Key"]  
        -----------  
        Description  
        Returns the key "Key" of the section "Section" from the C:\settings.ini file  
          
    .Link  
        Out-IniFile  
    #>  
      
    [CmdletBinding()]  
    Param(  
        [ValidateNotNullOrEmpty()]  
        [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".ini")})]  
        [Parameter(ValueFromPipeline=$True,Mandatory=$True)]  
        [string]$FilePath  
    )  
      
    Begin{
        Write-Log "  $($MyInvocation.MyCommand.Name):: Function started"
    }  
          
    Process{  
        Write-Log "  $($MyInvocation.MyCommand.Name):: Processing file: $Filepath"  
              
        $ini = @{}  
        switch -regex -file $FilePath  
        {  
            "^\[(.+)\]$" # Section  
            {  
                $section = $matches[1]  
                $ini[$section] = @{}  
                $CommentCount = 0  
            }  
            "^(;.*)$" # Comment  
            {  
                if (!($section))  
                {  
                    $section = "No-Section"  
                    $ini[$section] = @{}  
                }  
                $value = $matches[1]  
                $CommentCount = $CommentCount + 1  
                $name = "Comment" + $CommentCount  
                $ini[$section][$name] = $value  
            }   
            "(.+?)\s*=\s*(.*)" # Key  
            {  
                if (!($section))  
                {  
                    $section = "No-Section"  
                    $ini[$section] = @{}  
                }  
                $name,$value = $matches[1..2]  
                $ini[$section][$name] = $value  
            }  
        }  
        Write-Log "  $($MyInvocation.MyCommand.Name):: Finished Processing file: $FilePath"  
        Return $ini  
    }  
          
    End  
        {Write-Log "  $($MyInvocation.MyCommand.Name):: Function ended"}  
} 

Function Out-IniFile {  
    <#  
    .Synopsis  
        Write hash content to INI file  
          
    .Description  
        Write hash content to INI file  
          
        #Requires -Version 2.0  
          
    .Inputs  
        System.String  
        System.Collections.Hashtable  
          
    .Outputs  
        System.IO.FileSystemInfo  
          
    .Parameter Append  
        Adds the output to the end of an existing file, instead of replacing the file contents.  
          
    .Parameter InputObject  
        Specifies the Hashtable to be written to the file. Enter a variable that contains the objects or type a command or expression that gets the objects.  
  
    .Parameter FilePath  
        Specifies the path to the output file.  
       
     .Parameter Encoding  
        Specifies the type of character encoding used in the file. Valid values are "Unicode", "UTF7",  
         "UTF8", "UTF32", "ASCII", "BigEndianUnicode", "Default", and "OEM". "Unicode" is the default.  
          
        "Default" uses the encoding of the system's current ANSI code page.   
          
        "OEM" uses the current original equipment manufacturer code page identifier for the operating   
        system.  
       
     .Parameter Force  
        Allows the cmdlet to overwrite an existing read-only file. Even using the Force parameter, the cmdlet cannot override security restrictions.  
          
     .Parameter PassThru  
        Passes an object representing the location to the pipeline. By default, this cmdlet does not generate any output.  
                  
    .Example  
        Out-IniFile $IniVar "C:\myinifile.ini"  
        -----------  
        Description  
        Saves the content of the $IniVar Hashtable to the INI File c:\myinifile.ini  
          
    .Example  
        $IniVar | Out-IniFile "C:\myinifile.ini" -Force  
        -----------  
        Description  
        Saves the content of the $IniVar Hashtable to the INI File c:\myinifile.ini and overwrites the file if it is already present  
          
    .Example  
        $file = Out-IniFile $IniVar "C:\myinifile.ini" -PassThru  
        -----------  
        Description  
        Saves the content of the $IniVar Hashtable to the INI File c:\myinifile.ini and saves the file into $file  
  
    .Example  
        $Category1 = @{“Key1”=”Value1”;”Key2”=”Value2”}  
        $Category2 = @{“Key1”=”Value1”;”Key2”=”Value2”}  
        $NewINIContent = @{“Category1”=$Category1;”Category2”=$Category2}  
        Out-IniFile -InputObject $NewINIContent -FilePath "C:\MyNewFile.INI"  
        -----------  
        Description  
        Creating a custom Hashtable and saving it to C:\MyNewFile.INI  
    .Link  
        Get-IniContent  
    #>  
      
    [CmdletBinding()]  
    Param(  
        [switch]$Append,  
          
        [ValidateSet("Unicode","UTF7","UTF8","UTF32","ASCII","BigEndianUnicode","Default","OEM")]  
        [Parameter()]  
        [string]$Encoding = "Unicode",  
 
          
        [ValidateNotNullOrEmpty()]  
        [ValidatePattern('^([a-zA-Z]\:)?.+\.ini$')]  
        [Parameter(Mandatory=$True)]  
        [string]$FilePath,  
          
        [switch]$Force,  
          
        [ValidateNotNullOrEmpty()]  
        [Parameter(ValueFromPipeline=$True,Mandatory=$True)]  
        [Hashtable]$InputObject,  
          
        [switch]$Passthru  
    )  
      
    Begin{
        Write-Log "  $($MyInvocation.MyCommand.Name):: Function started"
    }  
          
    Process{  
        Write-Log "  $($MyInvocation.MyCommand.Name):: Writing to file: $Filepath"  
          
        If ($append){
            $outfile = Get-Item $FilePath
        }  
        Else{
            $outFile = New-Item -ItemType file -Path $Filepath -Force:$Force
        }  
        If(!($outFile)){
            Throw "Could not create File"
        }  
        Foreach($i in $InputObject.keys){  
            If (!($($InputObject[$i].GetType().Name) -eq "Hashtable")){  
                #No Sections  
                Write-Log "  $($MyInvocation.MyCommand.Name):: Writing key: $i"  
                Add-Content -Path $outFile -Value "$i=$($InputObject[$i])" -Encoding $Encoding  
            }
            Else{  
                #Sections  
                Write-Log "  $($MyInvocation.MyCommand.Name):: Writing Section: [$i]"  
                Add-Content -Path $outFile -Value "[$i]" -Encoding $Encoding  
                Foreach ($j in $($InputObject[$i].keys | Sort-Object))  
                {  
                    If($j -match "^Comment[\d]+"){  
                        Write-Log "  $($MyInvocation.MyCommand.Name):: Writing comment: $j"  
                        Add-Content -Path $outFile -Value "$($InputObject[$i][$j])" -Encoding $Encoding  
                    } 
                    Else{
                        Write-Log "  $($MyInvocation.MyCommand.Name):: Writing key: $j"  
                        Add-Content -Path $outFile -Value "$j=$($InputObject[$i][$j])" -Encoding $Encoding  
                    }  
                      
                }  
                Add-Content -Path $outFile -Value "" -Encoding $Encoding  
            }  
        }  
        Write-Log "  $($MyInvocation.MyCommand.Name):: Finished Writing to file: $path"  
        If($PassThru){
            Return $outFile
        }  
    }  
          
    End  
        {
            Write-Log "  $($MyInvocation.MyCommand.Name):: Function ended"
        }  
} 



<#'-------------------------------------------------------------------------------
  '---    Operating System
  '-------------------------------------------------------------------------------#>

Function Get-OSArchitecture{
    
    $GetOSArchitecture = (Get-WmiObject Win32_OperatingSystem -computername $env:COMPUTERNAME).OSArchitecture

    If($GetOSArchitecture -eq "64-BIT"){
        $GetOSArchitecture = "64"
    }
    Else{
        $GetOSArchitecture = "32"
    }

    Write-Log -Message "  GetOSArchitecture Returned: ""$GetOSArchitecture"""

    Return $GetOSArchitecture.ToString()
}


Function Get-OSName
{
    $OSName = Get-Itemproperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductName.ProductName
    Write-Log -Message "$OSName"

    Return $OSName

}


<#'-------------------------------------------------------------------------------
  '---    Registry
  '-------------------------------------------------------------------------------#>


Function Read-Registry {
    Param(
        [Parameter(Mandatory=$True)]
        $sPath,
        [Parameter(Mandatory=$True)]
        $sName
    )

    Write-Log -Message "  Read Registry: $sPath\$sName"
    
    If(!(Test-Path -Path $sPath)){
        Write-Log -Message "  Error: Registry Path Not Found"
        Return
    }
    $ReadRegistry = Get-ItemProperty -Path $sPath -Name $sKey -ErrorAction SilentlyContinue -ErrorVariable iErr | ForEach-Object {$_.$sName}
    If($iErr){
        Write-Log -Message "  Error: Unable to read registry"
    }
    Write-Log -Message "  Registry Key Value: '$ReadRegistry'"
    Return $ReadRegistry
}


Function Write-Registry{
    Param(
        [Parameter(Mandatory=$True)]
        $sPath, #Ex "HKLM:\SOFTWARE\Wow6432Node"
        [Parameter(Mandatory=$True)]
        $sName, #Ex "Adobe"
        [Parameter(Mandatory=$False)]
        $sType, #Ex "String"
        [Parameter(Mandatory=$True)]
        $sValue, #Ex "1.1.2.0"
        [Parameter(Mandatory=$True)]
        $Force
    )

    Write-Log -Message "  Write Registry: $sPath : $sName : $sType : $sValue"
    
    If(!(Test-Path -Path $sPath)){
        $sCMD = New-Item -Path $sPath -ErrorAction SilentlyContinue -ErrorVariable iErr;
        If($iErr){
            Write-Log -Message "  Error: Unable to write to registry"
            Return.$sCMD
        }
    }
    
    If((Read-Registry -sPath $sPath -sName $sName) -eq $sValue){
        Write-Log -Message "  Success: Registry Key Already Exists"
        Return
    }
    IF($Force -eq $True){
        Write-Log -Message "  Registry Key will be forcefully overwritten"
    }
    Else{
        Write-Log -Message '  Registry Key will not be overwritten, use -Force $True to forcefully overwrite'
        Return
    }
   
    $WriteRegistry = New-ItemProperty -Path $sPath -Name $sName -PropertyType $sType -Value $sValue -Force
    If((Get-ItemProperty $sPath -Name $sName -ErrorAction SilentlyContinue).$sName){
        Write-Log -Message "  Success: Registry key Written successfully"
    }
    Else{
        Write-Log -Message "  Error: Failed to write registry key"
        Return.$WriteRegistry
    }
}



<#'-------------------------------------------------------------------------------
 '---    Active Directory
 '-------------------------------------------------------------------------------#>


Function AD_ManageGroup{
    Param(
        [Parameter(Mandatory=$False)]
        $sDomain,
        [Parameter(Mandatory=$True)]
        $sFunction,
        [Parameter(Mandatory=$True)]
        $sType,
        [Parameter(Mandatory=$True)]
        $sName,
        [Parameter(Mandatory=$True)]
        $sGroup,
        [Parameter(Mandatory=$False)]
        $sADUser,
        [Parameter(Mandatory=$False)]
        $sADPass
    )

    Write-Log -Message "  AD_ManageGroup('$sDomain', '$sFunction', '$sType', '$sName', '$sGroup', '$sADUser', '$sADPass'"
    Write-Log -Message "  Importing Active Directory Module"
    
    $ImportModule = (Import-Module ActiveDirectory -PassThru -ErrorAction SilentlyContinue -ErrorVariable iErr).ExitCode
    
    If($iErr){Write-Log -Message " Error: Unable to import AD Module Error: $ImportModule" -type 3
        Return}
    Else{
        Write-Log -Message "  Success: Imported AD Module"
    }

    If($sType -eq "User"){
        $GetUser = Get-ADUser -Identity $sName -Properties MemberOf,sAMAccountName -ErrorAction SilentlyContinue -ErrorVariable iErr | Select-Object MemberOf,sAMAccountName

        If($iErr){
            Write-Log -Message "  Error: Unable to locate $sType : '$sName' in AD" -Type 3
            Return
        }

        If($sFunction -eq "Add"){
            If ($GetUser.MemberOf -match $sGroup){
                Write-Log -Message "  Success: '$sName' is already a member of '$sGroup'"
                Return}
            Else{
                $sCMD = Add-ADGroupMember -Identity "$sGroup" -Members "$sName" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
                If($iErr){
                    Write-Log -Message "  Error: Adding '$sName' to '$sGroup' failed" -type 3
                    Return
                } 
                Else{
                    Write-Log -Message "  Success: Added '$sName' to '$sGroup' successfully"
                }
            }
        }
        ElseIf($sFunction -eq "Remove"){
            If (!($GetUser.MemberOf -match $sGroup)){
                Write-Log -Message "  Success: '$sName' is already not a member of '$sGroup'"
                Return
            }
            $sCMD = Remove-ADGroupMember -Identity "$sGroup" -Members "$sName" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
            If($iErr){
                Write-Log -Message "  Error: Removing '$sName' from '$sGroup' failed" -type 3
                Return
            } 
            Else{
                Write-Log -Message "  Success: Removed '$sName' from '$sGroup' successfully"
            }
        }
        ElseIf($sFunction -eq "Query"){
            If ($GetUser.MemberOf -match $sGroup){
                    Write-Log -Message "  Query: '$sName' is a member of '$sGroup'"
                    Return}
            Else{
                Write-Log -Message "  Query: '$sName' is NOT a member of '$sGroup'"
            }
        }
     }


     If($sType -eq "Computer"){
        $GetComputer = Get-ADComputer $sName -Properties MemberOf -ErrorAction SilentlyContinue -ErrorVariable iErr

        If($iErr){
            Write-Log -Message "  Error: Unable to locate $sType : '$sName' in AD" -Type 3
            Return
        }


        If($sFunction -eq "Add"){
            If ($GetComputer.MemberOf -match $sGroup){
                Write-Log -Message "  Success: '$sName' is already a member of '$sGroup'"
                Return}
            Else{
                $sCMD = Add-ADGroupMember -Identity "$sGroup" -Members "$sName$" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
                If($iErr){
                    Write-Log -Message "  Error: Adding '$sName' to '$sGroup' failed" -type 3
                    Return
                } 
                Else{
                    Write-Log -Message "  Success: Added '$sName' to '$sGroup' successfully"
                }
            }
        }
        ElseIf($sFunction -eq "Remove"){
            If (!($GetComputer.Memberof -match $sGroup)){
                Write-Log -Message "  Success: '$sName' is already not a member of '$sGroup'"
                Return
            }
            $sCMD = Remove-ADGroupMember -Identity "$sGroup" -Members "$sName$" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
            If($iErr){
                Write-Log -Message "  Error: Removing '$sName' from '$sGroup' failed" -type 3
                Return
            } 
            Else{
                Write-Log -Message "  Success: Removed '$sName' from '$sGroup' successfully"
            }
        }
        ElseIf($sFunction -eq "Query"){
            If ($GetComputer.Memberof -Match $sGroup){
                Write-Log -Message "  Query: '$sName' is a member of '$sGroup'"
                Return
            }
            Else{
                Write-Log -Message "  Query: '$sName' is NOT a member of '$sGroup'"
            }
        }
     } 


     If($sType -eq "Group"){
        $GetGroup = (Get-ADGroup -Identity $sName -Properties MemberOf -ErrorAction SilentlyContinue -ErrorVariable iErr | Select-Object MemberOf)
        If($iErr){
            Write-Log -Message "  Error: Unable to locate $sType : '$sName' in AD" -Type 3
            Return
        }

        If($sFunction -eq "Add"){
            If ($GetGroup.MemberOf -match $sGroup){
                Write-Log -Message "  Success: '$sName' is already a member of '$sGroup'"
                Return
            }
            Else{
                $sCMD = Add-ADGroupMember -Identity "$sGroup" -Members "$sName" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
                If($iErr){
                    Write-Log -Message "  Error: Adding '$sName' to '$sGroup' failed" -type 3
                    Return
                } 
                Else{
                    Write-Log -Message "  Success: Added '$sName' to '$sGroup' successfully"
                }
            }
        }
        ElseIf($sFunction -eq "Remove"){
            If (!($GetGroup.MemberOf -match $sGroup)){
                Write-Log -Message "  Success: '$sName' is already not a member of '$sGroup'"
                Return
            }
            $sCMD = Remove-ADGroupMember -Identity "$sGroup" -Members "$sName" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
            If($iErr){
                Write-Log -Message "  Error: Removing '$sName' from '$sGroup' failed" -type 3
                Return
            } 
            Else{
                Write-Log -Message "  Success: Removed '$sName' from '$sGroup' successfully"
            }
        }
        ElseIf($sFunction -eq "Query"){
            If ($GetGroup.MemberOf -match $sGroup){
                Write-Log -Message "  Query: '$sName' is a member of '$sGroup'"
                Return
            }
            Else{
                Write-Log -Message "  Query: '$sName' is NOT a member of '$sGroup'"
            }
        }
     } 
}

Function AD_ManageGroupADSI{
<#
  .SYNOPSIS
    Adds/Removes Users/Computers/Groups To/From Groups
  .DESCRIPTION
    Does not require the Active Directory module and allows you to pass alternate credentials to run under.
    sFunction Options Are: Add or Remove
    sType Options Are: User, Computer, or Group
    sADUser and sADPass are optional and are only needed when alternate credentials need supplied
  .EXAMPLE
    AD_ManageGroupADSI -sDomain "contoso.org" -sFunction "Add" -sType "User" -sName "abc20a" -sGroup "GROUP_1"
  .EXAMPLE
    AD_ManageGroupADSI -sDomain "contoso.org" -sFunction "Add" -sType "Computer" -sName "DT12345678" -sGroup "GROUP_1" -sADUser "abc123a" -sADPass "P@ssw0rd"
  .EXAMPLE
    AD_ManageGroupADSI -sDomain "contoso.org" -sFunction "Remove" -sType "Computer" -sName "DT12345678" -sGroup "GROUP_1" -sADUser "abc123a" -sADPass "P@ssw0rd"
  .EXAMPLE
    AD_ManageGroupADSI -sDomain "contoso.org" -sFunction "Add" -sType "User" -sName "czt20b" -sGroup "GROUP_1" -sADUser "abc123a" -sADPass "P@ssw0rd"
  .EXAMPLE
    AD_ManageGroupADSI -sDomain "contoso.org" -sFunction "Add" -sType "Group" -sName "GROUP_42" -sGroup "GROUP_1" -sADUser "abc123a" -sADPass "P@ssw0rd"
  #>
    Param(
        [Parameter(Mandatory=$False)]
        $sDomain,
        [Parameter(Mandatory=$True)]
        $sFunction,
        [Parameter(Mandatory=$True)]
        $sType,
        [Parameter(Mandatory=$True)]
        $sName,
        [Parameter(Mandatory=$True)]
        $sGroup,
        [Parameter(Mandatory=$False)]
        $sADUser,
        [Parameter(Mandatory=$False)]
        $sADPass
    )

    [int]$ADS_PROPERTY_APPEND = 3

    $sGroupPath = Get-ADSPath -sDomain $sDomain -sType "Group" -sName $sGroup -sADUser $sADUser -sADPass $sADPass
    $sObjectPath = Get-ADSPath -sDomain $sDomain -sType $sType -sName $sName -sADUser $sADUser -sADPass $sADPass

    If($sGroupPath -eq $Null){
        Return
    }
    If($sObjectPath -eq $Null){
        Return
    }

    $sObjectDN = $sObjectPath.adspath.Replace("$sDomain/", "")
    #$sObjectDN = $sObjectPath.Replace("$sDomain/", "")
    #Write-Log -Message "  DN: $sObjectDN"

    $sObjectCN = $sObjectPath.adspath.Replace("LDAP://$sDomain/", "")
    #$sObjectCN = $sObjectPath.Replace("LDAP://$sDomain/", "")
    #Write-Log -Message "  CN: $sObjectCN"

    $sGroupDN = $sGroupPath.adspath.Replace("$sDomain/", "")
    #$sGroupDN = $sGroupPath.Replace("$sDomain/", "")
    #Write-Log -Message "  DN: $sGroupDN"
    
    If($sADUser -and $sADPass){
        $oGroup = New-Object DirectoryServices.DirectoryEntry($sGroupDN,$sADUser,$sADPass)
    }
	Else{
        $oGroup = [ADSI]$sGroupDN
    }

    $oComputer = [ADSI]$sObjectDN

    If($sFunction -eq "Add"){
        Try{
	        #Verify if the computer is a member of the Group
	        If ($oGroup.ismember($oComputer.adspath) -eq $False){
		        #Add the the computer to the specified group
		        $oGroup.PutEx($ADS_PROPERTY_APPEND,"member",@("$sObjectCN"))
		        $oGroup.setinfo()
                Write-Log -Message "  Successfully added $sName to $sGroup"
	        }
            Else{
                Write-Log -Message "  $sName is already a member of $sGroup"
            }
        }
        Catch{
            Write-Log -Message "  Unable to query $sGroup for membership status. If credentials were passed check that credentials are valid" -Type 3
            Return}
    }

    If($sFunction -eq "Remove"){
        Try{
            #Verify if the computer is a member of the Group
	        If ($oGroup.ismember($oComputer.adspath) -eq $False){
		        Write-Log -Message "  $sName is not a member of $sGroup"
	        }
            Else{
                #Add the the computer to the specified group
		        $oGroup.Member.Remove($sObjectCN)
		        $oGroup.setinfo()
                Write-Log -Message "  Successfully removed $sName from $sGroup"
            }
        }
        Catch{
            Write-Log -Message "  Unable to query $sGroup for membership status. If credentials were passed check that credentials are valid" -Type 3
            Return
        }
    }
}#End Function


Function AD_ManageComputers{
    Param(
        [Parameter(Mandatory=$True)]
        $Path,
        [Parameter(Mandatory=$False)]
        $Function,
        [Parameter(Mandatory=$False)]
        $DaysInactive
        )

        Write-Log -Message "  AD_ManageComputer($Path, $Function, $DaysInactive)"

        $ImportModule = (Import-Module ActiveDirectory -PassThru -ErrorAction SilentlyContinue -ErrorVariable iErr).ExitCode
    
            If($iErr){
                Write-Log -Message "  ERROR: Unable to import AD Module Error: $ImportModule" -type 3
                Return
            }
            Else{
                Write-Log -Message "  Success: Imported AD Module"
            }

        If($Function -eq "Disable"){
            
            $Machines = Get-Content -Path $Path -ErrorAction SilentlyContinue -ErrorVariable iErr;

            If($iErr){
                Write-Log -Message "  ERROR: Unable to locate or open $Path"
                Return
            }
            Else{
                Write-Log -Message "  SUCCESS: Machine List Loaded Sucessfully."
            }

            $Machines | foreach {
                Try{
                   $DisablePC =  Get-ADComputer -Identity $_ -ErrorAction SilentlyContinue | Disable-ADAccount -Confirm:$false -ErrorAction SilentlyContinue
                    Write-Log -Message "  ACTION: $_ Sucessfully Disabled in Active Directory."
                }
    
                Catch{
                    Write-Log -Message "  Error: ($_) Does not exist in Active Directory." -Type 3
                }
            }
        }

         If($Function -eq "Delete"){
       
            $Machines = Get-Content -Path $Path -ErrorAction SilentlyContinue -ErrorVariable iErr;

            If($iErr){
                Write-Log -Message "  ERROR: Unable to locate or open $Path"
                Return
            }
            Else{
                Write-Log -Message "  SUCCESS: Machine List Loaded Sucessfully."
            }

            $Machines | ForEach {
                Try{
                   $DisablePC =  Get-ADComputer -Identity $_ -ErrorAction SilentlyContinue | Delete-ADAccount -Confirm:$false -ErrorAction SilentlyContinue
                    Write-Log -Message "  ACTION: $_ Sucessfully Disabled in Active Directory."
                }
    
                Catch{
                    Write-Log -Message "  Error: ($_) Does not exist in Active Directory." -Type 3
                }
            }
        }

        If($Function -eq "QueryInactive"){

            $Time = (Get-Date).Adddays(-($DaysInactive))
 
            # Get all AD computers with lastLogonTimestamp less than our time
            Get-ADComputer -server "ghs.org" -Filter {LastLogonTimeStamp -lt $Time} -Properties LastLogonTimeStamp |
 
            # Output hostname and lastLogonTimestamp into CSV
            select-object Name | export-csv $Path -notypeinformation
        }
}



Function Get-ADSPath{
    Param(
        [Parameter(Mandatory = $True)]
            [String]$sDomain,
        [Parameter(Mandatory = $True)]
            [String]$sName,
        [Parameter(Mandatory = $True)]
            [String]$sType,
        [Parameter(Mandatory = $False)]
            [String]$sADUser,
        [Parameter(Mandatory = $False)]
            [String]$sADPass
    )

    If($sADUser -and $sADPass){
        $Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry("LDAP://$sDomain", $sADUser, $sADPass)
        Write-Log -Message "  Credentials supplied for: $sADUser"
    }
    Else{$Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry("LDAP://$sDomain")
        Write-Log -Message "  Running as invoked user"
    }

    If($sType -eq "User"){
        $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher($Domain,"(&(objectCategory=User)(sAMAccountname=$sName))")
        $Searcher.SearchScope = "Subtree"
        $Searcher.SizeLimit = '5000'
        $ADOQuery = $Searcher.FindAll()
        $ADSPath = $ADOQuery.Path
        $ADOProperties = $ADOQuery.Properties
        Write-Log -Message "  User: $sName ADSPath: $ADSPath"
        Return $ADOProperties
    }
    ElseIf($sType -eq "Group"){
        $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher($Domain,"(&(objectCategory=Group)(name=$sName))")
        $Searcher.SearchScope = "Subtree"
        $Searcher.SizeLimit = '5000'
        $ADOQuery = $Searcher.FindAll().Path
        Write-Log -Message "  Group: $sName ADSPath: $ADOQuery"
        Return [ADSI]$ADOQuery
    }
    ElseIf($sType -eq "Computer"){
        $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher($Domain,"(&(objectCategory=Computer)(name=$sName))")
        $Searcher.SearchScope = "Subtree"
        $Searcher.SizeLimit = '5000'
        $ADOQuery = $Searcher.FindAll().Path
        Write-Log -Message "  Computer: $sName ADSPath: $ADOQuery"
        Return [ADSI]$ADOQuery
    }
}




Function Set-AdUserPasswordADSI{ 
    Param(
        [Parameter(Mandatory = $True)]
            [String]$sDomain,
        [Parameter(Mandatory = $True)]
            [String]$sUser,
        [Parameter(Mandatory = $False)]
            [String]$sNewPassword,
        [Parameter(Mandatory = $False)]
            [String]$sADUser,
        [Parameter(Mandatory = $False)]
            [String]$sADPass
    )

    $oUser = Get-ADOProperty -sDomain $sDomain -sName "$sUser" -sType "User" -sADUser $sADUser -sADPass $sADPass
    $oUserDN = $oUser.distinguishedname
    $oUserFullDN = [ADSI]"LDAP://$oUserDN"
    $oUserFullDN.psbase.invoke("SetPassword",$sNewPassword)
    $oUserFullDN.psbase.CommitChanges()

} # end unction Set-AdUserPassword


Function Get-GroupMembershipADSI{
    Param(
        [Parameter(Mandatory = $True)]
            [String]$sDomain,
        [Parameter(Mandatory = $True)]
            [String]$sName,
        [Parameter(Mandatory = $True)]
            [String]$sType,
        [Parameter(Mandatory = $False)]
            [String]$sADUser,
        [Parameter(Mandatory = $False)]
            [String]$sADPass
    )

    Write-Log "  Get-GroupMembershipADSI: $sName"

    $oADObject = Get-ADSPath -sDomain "ghs.org" -sName $sName -sType $sType -sADUser $sADUser -sADPass $sADPass
    $Groups =  $oADObject.memberof | ForEach-Object {[ADSI]"LDAP://$_"}
    
    Return $Groups
}

Function Get-GroupMembership{
    Param(
        [Parameter(Mandatory=$True)]
        $User
    )

    ForEach ($U in $User){
        $UN = Get-ADUser $U -Properties MemberOf
        $Groups = ForEach ($Group in ($UN.MemberOf)){
            (Get-ADGroup $Group).Name
        }
        $Groups = $Groups | Sort
        ForEach ($Group in $Groups){
            New-Object PSObject -Property @{
                Name = $UN.Name
                Group = $Group
            }
        }
    }
}



<#Function Get-ADSPath{
    Param(
        [Parameter(Mandatory = $True)]
            [String]$sDomain,
        [Parameter(Mandatory = $True)]
            [String]$sName,
        [Parameter(Mandatory = $True)]
            [String]$sType,
        [Parameter(Mandatory = $False)]
            [String]$sADUser,
        [Parameter(Mandatory = $False)]
            [String]$sADPass,
        [Parameter(Mandatory = $False)]
            [String]$sProperties
    )

    If($sDomain -eq $null -or $sName -eq $null -or $sType -eq $null -or $sName -eq $null){
        Write-Log -Message "  Invalid Parameters Passed" -Type 3
        Return}

    [int] $ADS_PROPERTY_APPEND = 3
    $ADS_SECURE_AUTHENTICATION = '&H1'
    $ADS_SERVER_BIND = '&H200'
    [int] $ADS_SCOPE_SUBTREE = 2
    If($sADUser -and $sADPass){
        $DomainIP = (Test-Connection -ComputerName "$sDomain" -Count 1).IPV4Address.IPAddressToString
    
        If($sADUser -and $sADPass){
            $Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry("LDAP://$DomainIP", $sADUser, $sADPass)
        }
        Else{$Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry("LDAP://$DomainIP")}

        If($sType -eq "User"){
            $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher($Domain,"(&(objectCategory=User)(sAMAccountname=$sName))")
            $Searcher.SearchScope = "Subtree"
            $Searcher.SizeLimit = '5000'
            $ADOQuery = $Searcher.FindAll()
            $ADOProperties = $ADOQuery.Properties
            Return $ADOProperties
        }
        ElseIf($sType -eq "Group"){
            $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher($Domain,"(&(objectCategory=Group)(name=$sName))")
            $Searcher.SearchScope = "Subtree"
            $Searcher.SizeLimit = '5000'
            $ADOQuery = $Searcher.FindAll().Path
            Return [ADSI]$ADOQuery
        }
        ElseIf($sType -eq "Computer"){
            $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher($Domain,"(&(objectCategory=Computer)(name=$sName))")
            $Searcher.SearchScope = "Subtree"
            $Searcher.SizeLimit = '5000'
            $ADOQuery = $Searcher.FindAll().Path
            Return [ADSI]$ADOQuery
        }
    }
    Else{
        #Create ADODB connection
	    $oAD = New-Object -ComObject "ADODB.Connection"
	    $oAD.Provider = "ADsDSOObject"

	    $oAD.Open("Active Directory Provider")

        Write-Log -Message "  Connecting to AD"
	
        If($sType -eq "User"){
            $sQuery = "SELECT ADsPath,cn,sAMAccountName,manager FROM 'LDAP://$sDomain' WHERE objectCategory='$sType' AND  sAMAccountName='$sName'"}
        ElseIf($sType -eq "Computer" -or $sType -eq "Group"){
		    $sQuery = "SELECT ADsPath,cn,sAMAccountName FROM 'LDAP://$sDomain' WHERE objectCategory='$sType' AND  Name='$sName'"}
        Else{Write-Log -Message "  Invalid Object parameter passed. Must be User, Group, CustomUser, or Computer" -Type 3
            Return}

        Try{$oRS = $oAD.Execute($sQuery)}
        Catch{Write-Log -Message "  Unable to connect to AD" -Type 3
            Return}
        Finally{}

	    If (!$oRs.EOF)
	    {
            $sDomainsPath = $oRs.Fields("ADsPath").value
            $sCN = $oRs.Fields("cn").value
            Write-Log -Message "  CN: $sCN"
            Write-Log -Message "  ADsPath: $sDomainsPath"
	    }
        If($sDomainsPath -eq $null){
            Write-Log -Message "  Unable to locate $sType : $sName in Active Directory" -Type 3
            Return}

        Return $oRS
    }
}#End Get-ADSPath
#>






<#'-------------------------------------------------------------------------------
  '---    SCCM
  '-------------------------------------------------------------------------------#>

Function Is-InTS{

    Try{
        $Global:oENV = New-Object -COMObject Microsoft.SMS.TSEnvironment
    }
    Catch{
        Return $False
    }
    Return $True
}


Function Hide-TSProgress{
    Try{
        $TSProgressUI = New-Object -COMObject Microsoft.SMS.TSProgressUI
        $TSProgressUI.CloseProgressDialog()
        Write-Log -Message "  Successfully hid TS Progress Window"
    }
    Catch{
        Write-Log -Message "  Unable to hide TS Progress Window"
    }
 }


 Function Get-TSVar{
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Variable
    )
    Write-Log "  Get-TSVar($Variable)"

    Try{
        $TSENV = New-Object -COMObject Microsoft.SMS.TSEnvironment
        $TSVariable = $TSENV.Value($Variable)
        Write-Log -Message "  TS Variable: $Variable = $TSVariable"
    }
    Catch{
        Write-Log -Message "  Failed to get value of $Variable"
    }

    Return $TSVariable
}

Function Set-TSVar{
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Variable,
        [Parameter(Mandatory=$True)]
        [String]$Value
    )
    Write-Log -Message "  Set-TSVar($Variable,$Value)"

    Try{
        $TSENV = New-Object -COMObject Microsoft.SMS.TSEnvironment
    }
    Catch{}

    Try{
        $TSENV.Value($Variable) = $Value
        Write-Log -Message "  Successfully Set Variable"
    }
    Catch{
        Write-Log -Message "  Failed to Set Variable"
    }
}

Function Add-ComputerToCollection {
     Param(
        [Parameter(Mandatory=$True)]
        $CompName,
        [Parameter(Mandatory=$True)]
        $ColID,
        [Parameter(Mandatory=$True)]
        $ColName,
        [Parameter(Mandatory=$True)]
        $Server
        )

	$RulesToSkip = $null
	$strMessage = "Do you want to add '$CompName' to '$colname'"
	
        $sccmProviderLocation = Get-WmiObject -query "select * from SMS_ProviderLocation where ProviderForLocalSite = true" -Namespace "root\sms" -computername $Server
        $SiteCode = $sccmProviderLocation.SiteCode
        $Namespace = "root\sms\site_$SiteCode"

		$strQuery = "Select * from SMS_Collection where CollectionID = '$ColID'"
		$Collection = Get-WmiObject -query $strQuery -ComputerName $Server -Namespace $Namespace
		$Collection.Get()

		If($CompName -ne $null){
			$strQuery = "Select * from SMS_R_System where Name like '$CompName'"
			Get-WmiObject -Query $strQuery -Namespace $Namespace -ComputerName $Server | ForEach-Object {
			    $ResourceID = $_.ResourceID
			    $RuleName = $_.Name
			    $CompName = $RuleName
			    If($ResourceID -ne $null){
				    $Error.Clear()
				    $Collection=[WMI]"\\$($Server)\$($Namespace):SMS_Collection.CollectionID='$ColID'"
				    $RuleClass = [wmiclass]"\\$($Server)\$($NameSpace):SMS_CollectionRuleDirect"
				    $newRule = $ruleClass.CreateInstance()
				    $newRule.RuleName = $RuleName
				    $newRule.ResourceClassName = "SMS_R_System"
				    $newRule.ResourceID = $ResourceID
				    $Collection.AddMembershipRule($newRule)
				    If ($Error[0]) {
					    Write-Log -Message "Error adding $CompName - $Error"
					    $ErrorMessage = "$Error"
					    $ErrorMessage = $ErrorMessage.Replace("`n","")
				    }
				    Else {
					    Write-Log -Message "Successfully added $CompName"
                        Return $True
				    }
			    }
			    Else {
				    Write-Log -Message "Could not find $CompName - No rule added"
			    }
			}#End For-Each
			If($ResourceID -eq $null){
				Write-Log -Message "Could not find $CompName - No rule added"
			}
		}
}


<# 
        Script Library
  
        File:       ScriptLibrary1.18.psm1
    
        Purpose:    Contains common functions and routines useful in other scripts

        Author: Brandon Hilgeman 
                brandon.hilgeman@gmail.com

#>

<#
	.SYNOPSIS
		Sets Variables to be used globally throughout the module
	
	.DESCRIPTION
		Should be called immediately after importing this module

	.EXAMPLE
				PS C:\> Set-GlobalVariables
	
	.NOTES
		Additional information about the function.
#>
function Set-GlobalVariables {
	$Global:ComputerName = $env:computername
	$Global:ComputerDomain = "$env:COMPUTERNAME.$env:USERDNSDOMAIN"
	$Global:WindowsPath = $env:windir + "\"
	$Global:DefaultLogPath = $env:TEMP + "\"
	$Global:DefaultLog = $DefaultLogPath + $ScriptName + ".log"
	$Global:MaxLogSizeInKB = 1024 * 20
	$Global:TempPath = $env:TEMP + "\"
	Detect-Runtime
}

<#
	.SYNOPSIS
		Detects environment script is running in
	
	.DESCRIPTION
		Detects whether script is running standalone, in MDT, or in SCCM to determine logging path.
	
	.EXAMPLE
				PS C:\> Detect-Runtime
	
	.NOTES
		Additional information about the function.
#>
function Detect-Runtime {
	$IsInTS = $True
	Try {
		$oENV = New-Object -COMObject Microsoft.SMS.TSEnvironment
	} Catch {
		$IsInTS = $False
	}
	
	If ($IsInTS -eq $True) {
		
		If ($oENV.Value("DEPLOYMENTMETHOD").ToUpper() -eq "SCCM") {
			$Global:RunTime = "SCCM"
		} ElseIf ($oENV.Value("DEPLOYMENTMETHOD").ToUpper() -eq "MDT" -or $oENV.Value("DEPLOYMENTMETHOD").ToUpper() -eq "UNC" -or $oENV.Value("DEPLOYMENTMETHOD").ToUpper() -eq "MEDIA") {
			$Global:RunTime = "MDT"
		} Else {
			$Global:RunTime = "UNKNOWN"
		}
	} Else {
		$Global:RunTime = "STANDALONE"
	}
	
	If ($Runtime -eq "SCCM") {
		$Global:Log = $oENV.Value("LogPath") + "\" + $ScriptName + ".log"
		$Global:LogPath = $oENV.Value("LogPath") + "\"
		$DeployRoot = $oENV.Value("DeployRoot") + "\"
		$ScriptRoot = $oENV.Value("ScriptRoot") + "\"
	} ElseIf ($Runtime -eq "MDT") {
		$Global:Log = $oENV.Value("LogPath") + "\" + $ScriptName + ".log"
		$Global:LogPath = $oENV.Value("LogPath") + "\"
		$DeployRoot = $oENV.Value("DeployRoot") + "\"
		$ScriptRoot = $oENV.Value("ScriptRoot") + "\"
	} ElseIf ($Runtime -eq "STANDALONE") {
		$Global:Log = $DefaultLog
		$Global:LogPath = $DefaultLogPath
	} Else {
		$Global:Log = $TempPath + $ScriptName + ".log"
		$Global:LogPath = $TempPath
	}
	
	If (($IsInTS -eq $False) -and ($WindowsPath.substring(0, 3) -eq "X:\")) {
		$Global:RunTime = "MDT"
		$Global:Log = "X:\MININT\SMSOSD\OSDLOGS\$ScriptName.log"
		$Global:LogPath = "X:\MININT\SMSOSD\OSDLOGS\"
		$Global:BackupLog = $TempPath + $ScriptName + ".log"
	}
}

<#
	.SYNOPSIS
		Writes to a logfile in CMTrace formatting
	
	.DESCRIPTION
		Writes provided parameter to a logfile
	
	.PARAMETER Message
		The information to be written to the log file
	
	.PARAMETER ErrorMessage
		A description of the ErrorMessage parameter.
	
	.PARAMETER Component
		If component is not passed, the script name will be used.
	
	.PARAMETER Type
		1 = Info
		2 = Warning
		3 = Error
		Will mark the line in CMTrace appropriate color based on Type
		If Type is not passed, default is 1
	
	.PARAMETER LogFile
		Full path to logfile to be written to.
		Default is set by Set-GlobalVariables, alternate can be passed to this function.
	
	.EXAMPLE
		Write-Log -Message "Hello World"
		Write-Log -Message "Hello World" -Type 1
		Write-Log -Message "Error Occurred" -Type 3
		Write-Log -Message "Hello World" -Type 1 -LogFile "C:\AlternateLog.log"
	
	
	.NOTES
		Additional information about the function.
#>
function Write-Log {
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$Message,
		[Parameter(Mandatory = $false)]
		[string]$ErrorMessage,
		[Parameter(Mandatory = $false)]
		[string]$Component = $ScriptName,
		[Parameter(Mandatory = $false)]
		[int]$Type = 1,
		[Parameter(Mandatory = $false)]
		$LogFile
	)
	
	$Time = Get-Date -Format "HH:mm:ss.ffffff"
	$Date = Get-Date -Format "MM-dd-yyyy"
	
	<#If ($ErrorMessage -ne $null) {
		$Type = 3
	}#>
	
	If ($LogFile -eq $null) {
		$LogFile = $Log
	} Else {
		$Log = $LogFile
	}
	
	If (!(Test-Path -Path $LogPath)) {
		MkDir $LogPath | Out-Null
	}
	
	$LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"
	$LogMessage | Out-File -Append -Encoding UTF8 -FilePath $Log
	
	If ($LogFile -eq $null) {
		If ((Get-Item $Log).Length/1KB -gt $MaxLogSizeInKB) {
			$log = $Log
			Remove-Item ($log.Replace(".log", ".lo_"))
			Rename-Item $Log($log.Replace(".log", ".lo_")) -Force
		}
	}
}

<#
	.SYNOPSIS
		Writes a timestamp to the first line of the log
	
	.DESCRIPTION
		Writes a timestamp to the first line of the log and captures start time to be compared with Stop time in End-Log
	
	.EXAMPLE
				PS C:\> Start-Log
	
	.NOTES
		Additional information about the function.
#>
function Start-Log {
	$StartDate = ((get-date).toShortDateString() + " " + (get-date).toShortTimeString())
	$global:ScriptStart = get-date
	Write-Log -Message ("-" * 10 + "  Start Script: $ScriptName " + $StartDate + " " + "-" * 10)
}

<#
	.SYNOPSIS
		Last thing to run in a script to stamp log with time stamp
	
	.DESCRIPTION
		Stamps log with time stamp and compares to start time in Start-Log to write script execution length in minutes
	
	.EXAMPLE
				PS C:\> End-Log
	
	.NOTES
		Additional information about the function.
#>
function End-Log {
	$EndDate = ((get-date).toShortDateString() + " " + (get-date).toShortTimeString())
	$ScriptEnd = get-date
	#$RunTime = ($ScriptEnd.Minute - $ScriptStart.Minute)
	$RunTime = ($ScriptEnd - $ScriptStart)
	$RunTime = $RunTime.TotalMinutes
	$RunTime = [math]::round($Runtime, 3)
	Write-Log -Message ("-" * 10 + "  End Script: " + $EndDate + "   RunTime: " + $RunTime + " Minute(s) " + "-" * 10)
}

Function Scan-Args {
	
}

<#
	.SYNOPSIS
		Determines if script should start the install or uninstall function when executed
	
	.DESCRIPTION
		Determines if script should start the install or uninstall function when executed. This isn't used in GUI scripts.
		$sArgs variable is read to determine if an Uninstall parameter was passed to the script.
	
	.EXAMPLE
				PS C:>.\InstallApplication.ps1
				PS C:>.\InstallApplication.ps1 -Uninstall
				PS C:>.\InstallApplication.ps1 /Uninstall
	
	.NOTES
		
#>
function Set-Mode {
	If ($sArgs -eq "/Uninstall" -or $sArgs -eq "-Uninstall") {
		Write-Log "  $($MyInvocation.MyCommand.Name):: UNINSTALL"
		Start-Uninstall
	} Else {
		Write-Log "  $($MyInvocation.MyCommand.Name):: INSTALL"
		Start-Install
	}
}






<#'-------------------------------------------------------------------------------
  '---    GUI
  '-------------------------------------------------------------------------------#>


<#
	.SYNOPSIS
		Displays a dialog box for the user
	
	.DESCRIPTION
		Displays a dialog box for the user with either information and/or buttons to select an option. Button clicked is captured and returned from function
	
	.PARAMETER Message
		Message to be displayed within dialog box
	
	.PARAMETER Title
		Title in dialog box frame
	
	.PARAMETER Buttons
		Buttons to be displayed in dialog box
		Default = 0
		Options are:
		0 = OK
		1 = OKCancel
		2 = AbortRetryIgnore
		3 = YesNoCanel
		4 = YesNo
		5 = RetryCancel
	
	.PARAMETER Icon
		Type of icon to be displayed.
		Default is Exclamation
		Options are:
		Asterisk, Error, Exclamation, Hand, Information, None, Question, Stop, Warning
	
	.EXAMPLE
		PS C:\> Show-MsgBox -Message "Hello User"
		PS C:\> Show-MsgBox -Message "Do You Want To Continue?" -Title "Continue?" -Buttons 4 -Icon "Question"
	
	.NOTES
		Additional information about the function.
#>
function Show-MessageBox {
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$Message,
		[Parameter(Mandatory = $false)]
		[String]$Title,
		[Parameter(Mandatory = $false)]
		[int]$Buttons = 0,
		[Parameter(Mandatory = $false)]
		[String]$Icon = 'Exclamation'
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	
	Write-Log -Message ("  $($MyInvocation.MyCommand.Name):: ""$Message"", ""$Title"", ""$Buttons""") -Type 1
	
	$MessageBox = [System.Windows.Forms.MessageBox]::Show("$Message", "$Title", $Buttons, $Icon)
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $MsgBox pressed" -Type 1
	
	Return $MessageBox
}


<#
	.SYNOPSIS
		Reads XAML file and formats properly
	
	.DESCRIPTION
		Reads XAML file to display GUI from XAML code.
	
	.PARAMETER Path
		Path to the XAML file
	
	.PARAMETER bVariables
		Boolean, whether all variables in XAML should be output to console
	
	.EXAMPLE
		PS C:\> Get-XAML -Path $value1
	
	.NOTES
		Additional information about the function.
#>
function Get-XAML {
	param
	(
		[Parameter(Mandatory = $true)]
		$Path,
		[Parameter(Mandatory = $false)]
		[boolean]$bVariables = $False
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Path"
	
	Try {
		$InputXML = Get-Content $Path
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to locate XAML file at: $Path" -Type 3
	}
	$InputXML = $InputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
	
	[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
	[xml]$XAML = $InputXML
	#Read XAML
	
	$Reader = (New-Object System.Xml.XmlNodeReader $XAML)
	Try {
		$Global:Form = [Windows.Markup.XamlReader]::Load($Reader)
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed." -Type 3
	}
	
	$XAML.SelectNodes("//*[@Name]") | %{
		Set-Variable -Name "WPF$($_.Name)" -Value $Global:Form.FindName($_.Name) -Scope Global
	}
	#$WPFimage_logo.Source = "$ScriptDirectory\Images\Logo.bmp"
	
	If ($bVariables) {
		Get-FormVariables
	}
	
	Start-WPFApp
}


Function Get-FormVariables {
	If ($Global:ReadmeDisplay -ne $True) {
		Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow; $global:ReadmeDisplay = $true
	}
	Write-Host "Found the following interactable elements from our form" -ForegroundColor Cyan
	Get-Variable WPF*
}


<#
	.SYNOPSIS
		Displays part of XAML GUI
	
	.DESCRIPTION
		Makes a part of a XAML GUI visible
	
	.PARAMETER WPFVariable
		A description of the WPFVariable parameter.
	
	.EXAMPLE
				PS C:\> Show-GUI -WPFVariable $Button1
	
	.NOTES
		Additional information about the function.
#>
function Show-GUI {
	param
	(
		[Parameter(Mandatory = $false)]
		$WPFVariable
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Try {
		$WPFVariable.Visibility = "Visible"
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to show: $WPFVariable"
	}
}


<#
	.SYNOPSIS
		Hides part of a XAML GUI
	
	.DESCRIPTION
		A detailed description of the Hide-GUI function.
	
	.PARAMETER WPFVariable
		A description of the WPFVariable parameter.
	
	.EXAMPLE
				Hide-GUI -WPFVariable $Button1
	
	.NOTES
		Additional information about the function.
#>
function Hide-GUI {
	param
	(
		[Parameter(Mandatory = $false)]
		$WPFVariable
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Try {
		$WPFVariable.Visibility = "Hidden"
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to hide: $WPFVariable"
	}
}


Function Enable-GUI {
	Param (
		[Parameter(Mandatory = $False)]
		$WPFVariable
	)
	
	Try {
		$WPFVariable.IsEnabled = $True
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to Enable: $WPFVariable"
	}
}


Function Disable-GUI {
	Param (
		[Parameter(Mandatory = $False)]
		$WPFVariable
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Try {
		$WPFVariable.IsEnabled = $False
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to Disable: $WPFVariable"
	}
}


Function Clear-GUI {
	Param (
		[Parameter(Mandatory = $False)]
		$WPFVariable,
		[Parameter(Mandatory = $False)]
		$Type
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	If ($Type -eq "text") {
		Try {
			$WPFVariable.Clear()
		} Catch {
		}
	} ElseIf ($Type -eq "combo") {
		Try {
			$WPFVariable.Items.Clear()
		} Catch {
		}
	}
}

Function Add-GUIText {
	Param (
		[Parameter(Mandatory = $False)]
		$WPFVariable,
		[Parameter(Mandatory = $False)]
		$Text
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Try {
		$WPFVariable.AddText($Text)
	} Catch {
		Write-Log "  $($MyInvocation.MyCommand.Name):: Failed to add text to: $WPFVariable"
	}
}


<#
	.SYNOPSIS
		Displays a balloon tip (Toast)
	
	.DESCRIPTION
		A detailed description of the Show-BalloonTip function.
	
	.PARAMETER Title
		Title to Balloon Tip
	
	.PARAMETER Type
		Type of balloon tip
	
	.PARAMETER Message
		Message to be displayed in balloon tip
	
	.PARAMETER Duration
		Duration balloon tip should be displayed (In Milliseconds)
	
	.EXAMPLE
				PS C:\> Show-BalloonTip -Title 'Value1' -Message 'Value2'
	
	.NOTES
		Additional information about the function.
#>
function Show-BalloonTip {
	param
	(
		[Parameter(Mandatory = $true)]
		[AllowNull()]
		[string]$Title,
		[Parameter(Mandatory = $false)]
		[string]$Type = 'Info',
		[Parameter(Mandatory = $true)]
		[string]$Message,
		[Parameter(Mandatory = $false)]
		[Int]$Duration = 5000
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
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


<#
	.SYNOPSIS
		Test if computer is pingable
	
	.DESCRIPTION
		Test if computer is pingable to determine if remote computer is accessible over the network or online.
	
	.PARAMETER HostName
		Hostname of computer to ping
	
	.EXAMPLE
				PS C:\> Test-Ping -HostName "PC1"
	
	.NOTES
		Returns Boolean
#>
function Test-Ping {
	param
	(
		[Parameter(Mandatory = $true)]
		$HostName
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Pinging $HostName"
	
	$TestPing = Test-Connection -ComputerName $HostName -ErrorAction SilentlyContinue -ErrorVariable iErr;
	If ($iErr) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name)::ERROR Unable to Ping: $HostName" -Type 3
		$TestPing = $False
	} Else {
		Write-Log "  $($MyInvocation.MyCommand.Name):: $HostName SUCCESS"
		$TestPing = $True
		Return $TestPing
	}
}

<#
	.SYNOPSIS
		Gets the name of the local host
	
	.DESCRIPTION
		Gets the name of the local host
	
	.EXAMPLE
				PS C:\> Get-ComputerName
	
	.NOTES
		Returns hostname of local computer
#>
function Get-ComputerName {
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	$GetComputerName = $env:computername
	
	Write-Log "  $($MyInvocation.MyCommand.Name)::$GetComputerName"
	
	Return $GetComputerName
}

<#
	.SYNOPSIS
		Gets manufacturer of computer from WMI
	
	.DESCRIPTION
		Gets manufacturer of computer from WMI
	
	.EXAMPLE
				PS C:\> Get-Manufacturer
	
	.NOTES
		Returns local computer manufacturer as string
#>
function Get-Manufacturer {
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Try {
		$GetManufacturer = (Get-WmiObject -Class Win32_ComputerSystem).Manufacturer
	} Catch {
		Write-Log "  $($MyInvocation.MyCommand.Name)::Unable to find Manufacturer" -Type 3
	}
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: '$GetManufacturer'"
	Return $GetManufacturer
}



<#
	.SYNOPSIS
		Returns model of local computer from WMI
	
	.DESCRIPTION
		Returns model of local computer from WMI
	
	.EXAMPLE
				PS C:\> Get-Model
	
	.NOTES
		Returns local computer model as string
#>
function Get-Model {
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Try {
		$GetModel = (Get-WmiObject -Class Win32_ComputerSystem).Model
	} Catch {
		Write-Log "  $($MyInvocation.MyCommand.Name)::Unable to find Model" -Type 3
		Return
	}
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: '$GetModel'"
	Return $GetModel
}

<#
	.SYNOPSIS
		Returns the username of the currently logged on user
	
	.DESCRIPTION
		Returns the username of the currently logged on user
	
	.EXAMPLE
				PS C:\> Get-CurrentUser
	
	.NOTES
		Returns the username of the currently logged on user as string
#>
function Get-CurrentUser {
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Try {
		$GetCurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
	} Catch {
		Write-Log "  $($MyInvocation.MyCommand.Name)::Unable to find Logged on user" -Type 3
	}
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: '$GetCurrentUser'"
	Return $GetCurrentUser
}




<#'-------------------------------------------------------------------------------
  '---    Processes
  '-------------------------------------------------------------------------------#>

<#
	.SYNOPSIS
		Determines if a given process is running
	
	.DESCRIPTION
		Determines if a given process is running. Returns boolean
	
	.PARAMETER Process
		Name of process to check
	
	.EXAMPLE
				PS C:\> Is-ProcessRunning -Process "Notepad.exe"
	
	.NOTES
		Automatically trims .exe from end of process name
		Returns boolean
#>
function Is-ProcessRunning {
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$Process
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	If ($Process -contains '"') {
		$Process = $Process -replace '"', ''
	}
	If ($Process.substring($Process.length - 4, 4) -eq ".exe") {
		$Process = $Process -replace '.exe', ''
	}
	
	$IsProcessRunning = Get-Process $Process -ErrorAction SilentlyContinue
	#$IsProcessRunning
	If ($IsProcessRunning) {
		$RunningProcess = $True
	} Else {
		$RunningProcess = $False
	}
	Write-Log "  $($MyInvocation.MyCommand.Name)::'$Process' Returned: $RunningProcess"
	Return $RunningProcess
}


<#
	.SYNOPSIS
		Terminates a process
	
	.DESCRIPTION
		Terminates a given process
	
	.PARAMETER Process
		Name of process to terminate
	
	.EXAMPLE
				PS C:\> End-Process -Process "Notepad.exe"
	
	.NOTES
		Automatically trims .exe from end of process name
#>
function End-Process {
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$Process
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Ending $Process"
	
	If ($Process -contains '"') {
		$Process = $Process -replace '"', ''
	}
	If ($Process -contains '"') {
		$Process = $Process -replace '"', ''
	}
	If ($Process.substring($Process.length - 4, 4) -eq ".exe") {
		$Process = $Process -replace '.exe', ''
	}
	
	$RunningProcess = Get-Process $Process -ErrorAction SilentlyContinue
	
	If (Is-ProcessRunning -Process $Process) {
		$RunningProcess.CloseMainWindow()
		Sleep 3
	} Else {
		Write-Log "  $($MyInvocation.MyCommand.Name)::'$Process' is not running"
		Return
	}
	
	If (!$RunningProcess.HasExited) {
		$RunningProcess | Stop-Process -Force
		Sleep 3
	}
	
	If (Is-ProcessRunning -Process $Process) {
		Write-Log "  $($MyInvocation.MyCommand.Name)::'$Process' UnSuccessfully terminated"
	} Else {
		Write-Log "  $($MyInvocation.MyCommand.Name)::'$Process' Successfully terminated"
	}
}


<#'-------------------------------------------------------------------------------
  '---    Software
  '-------------------------------------------------------------------------------#>

<#
	.SYNOPSIS
		Runs a process or installs an application
	
	.DESCRIPTION
		Runs a process or installs an application
	
	.PARAMETER CMD
		Process to execute
	
	.PARAMETER Parameter
		Arguments to pass to the process
	
	.PARAMETER Hidden
		Boolean
		Is the process run hidden or visible
	
	.PARAMETER Wait
		Boolean
		Determines if the function should wait for the process to exit before returning
	
	.PARAMETER StatusOutput
		Boolean
		Should the log output also write to a Status Output windows in a GUI
	
	.PARAMETER StatusOutputTextBox
		Variable for the Status Output (Richtextbox) in a GUI
	
	.EXAMPLE
				Run-Install -CMD "$PSScriptRoot\Source\Install.exe" -Parameter "/silent /wait"
				Run-Install -CMD "$PSScriptRoot\Source\Install.msi" -Parameter "/qn /noreboot"
				Run-Install -CMD "$PSScriptRoot\Source\Install.exe" -Parameter "/silent /wait" -StatusOutPut $True -StatusOutputTextBox $StatusOutputTextBox
	
	.NOTES
		Returns exit code from process if -wait $True
#>
function Run-Install {
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$CMD,
		[String]$Parameter,
		[Boolean]$Hidden = $True,
		[boolean]$Wait = $True,
		[Boolean]$StatusOutput = $False,
		$StatusOutputTextBox
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	$ENV:SEE_MASK_NOZONECHECKS = 1
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: ""$CMD"" $Parameter" -type 1
	If ($StatusOutput) {
		Write-StatusOutput -Textbox $StatusOutputTextBox -Text """$CMD"" $Parameter" -Append $true -Tabbed $false
	}
	
	If ($Hidden) {
		If ($Wait) {
			$RunInstall = Start-Process $CMD -ArgumentList $Parameter -PassThru -Wait -WindowStyle Hidden
		} Else {
			$RunInstall = Start-Process $CMD -ArgumentList $Parameter -PassThru -WindowStyle Hidden
		}
	} Else {
		If ($Wait) {
			$RunInstall = Start-Process $CMD -ArgumentList $Parameter -PassThru -Wait
		} Else {
			$RunInstall = Start-Process $CMD -ArgumentList $Parameter -PassThru
		}
	}
	If ($Wait) {
		$ErrorCode = $RunInstall.ExitCode
		
		If ($ErrorCode -ne 0) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Command Completed with Error: $ErrorCode" -Type 3
			If ($StatusOutput) {
				Write-StatusOutput -Textbox $StatusOutputTextBox -Text "Command Completed with Error: $ErrorCode" -Append $true -Tabbed $false -Color "Red"
			}
		} Else {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Command Completed Successfully RETURN CODE: $ErrorCode" -Type 1
			If ($StatusOutput) {
				Write-StatusOutput -Textbox $StatusOutputTextBox -Text "Command Completed Successfully RETURN CODE: $ErrorCode" -Append $true -Tabbed $false
			}
		}
		
		$ENV:SEE_MASK_NOZONECHECKS = 0
		
		Return $ErrorCode
	}
}


<#
	.SYNOPSIS
		Uninstalls an application using WMI
	
	.DESCRIPTION
		Uninstalls an application using WMI.
		Use * as wildcard in name or version
	
	.PARAMETER Name
		Name of the applciation as noted in WMI
	
	.PARAMETER Version
		Version of the application as noted in WMI
	
	.EXAMPLE
		Run-Uninstall -Name "*Adobe Reader*"
		Run-Uninstall -Name "*Adobe Reader*" -Version "*.123"
	
	.NOTES
		Returns ExitCode from uninstall process
#>
function Run-Uninstall {
	param
	(
		[Parameter(Mandatory = $true)]
		$Name,
		[Parameter(Mandatory = $false)]
		$Version = '*'
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: ""$Name"", ""$Version"""
	
	$cItems = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE '$Name' And Version Like '$Version'"
	If (!$cItems) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Application: ""$Name"" Version: ""$Version"" - Not found on system"
		Return
	}
	
	ForEach ($oItem in $cItems) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: "($oItem.Name + ", " + $oItem.Version) -Type 1
		Try {
			$RunUninstall = $oItem.Uninstall()
			$RunUninstall = $true
		} Catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Uninstall Completed with Error: "$RunUninstall.ExitCode -Type 3
			$RunUninstall = $false
			Return $RunUninstall
		}
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Uninstall Completed Successfully" -Type 1
		Return $RunUninstall
	}
}


<#
	.SYNOPSIS
		Logs All Software Installed on System using Registry entries
	
	.DESCRIPTION
		Logs All Software Installed on System using Registry entries
	
	.EXAMPLE
		Log-InstalledSoftware
	
	.NOTES
		Check Log for list of Installed Software
#>
function Log-InstalledSoftware {
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	$UninstallRegKeys = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
	
	ForEach ($Computer in $ComputerName) {
		If (Test-Connection -ComputerName $Computer -Count 1 -ea 0) {
			ForEach ($UninstallRegKey in $UninstallRegKeys) {
				Try {
					$HKLM = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $computer)
					$UninstallRef = $HKLM.OpenSubKey($UninstallRegKey)
					$Applications = $UninstallRef.GetSubKeyNames()
				} Catch {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to read $UninstallRegKey"
					Continue
				}
			}
		}
	}
	
	ForEach ($App in $Applications) {
		$AppRegistryKey = $UninstallRegKey + "\\" + $App
		$AppDetails = $HKLM.OpenSubKey($AppRegistryKey)
		$AppGUID = $App
		$AppDisplayName = $($AppDetails.GetValue("DisplayName"))
		$AppVersion = $($AppDetails.GetValue("DisplayVersion"))
		$AppPublisher = $($AppDetails.GetValue("Publisher"))
		$AppInstalledDate = $($AppDetails.GetValue("InstallDate"))
		$AppUninstall = $($AppDetails.GetValue("UninstallString"))
		If ($UninstallRegKey -match "Wow6432Node") {
			$Softwarearchitecture = "x86"
		} Else {
			$Softwarearchitecture = "x64"
		}
		If (!$AppDisplayName) {
			continue
		}
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: ""$AppDisplayName""  Version: ""$AppVersion"""
	}
}

<#
	.SYNOPSIS
		Queries registry if provided software name and version are installed
	
	.DESCRIPTION
		Queries registry to determine if proviced software name and version are installed.
	
	.PARAMETER Product
		Name of the product to query the registry for
	
	.PARAMETER Version
		Version of the software to query the registry for
	
	.EXAMPLE
		Is-SoftwareInstalled -Product "*Adobe Reader*"
		Is-SoftwareInstalled -Product "*Adobe Reader*" -Version "*.123"
	
	.NOTES
		Returns Boolean
#>
function Is-SoftwareInstalled {
	param
	(
		[Parameter(Mandatory = $true)]
		$Product,
		[Parameter(Mandatory = $false)]
		$Version = '*'
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	$IsSoftwareInstalled = $False
	
	$Product.Replace("%", "*")
	$Version.Replace("%", "*")
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: ""$Product"" ""$Version"""
	
	$UninstallRegKeys = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
	
	ForEach ($Computer in $ComputerName) {
		If (Test-Connection -ComputerName $Computer -Count 1 -ErrorAction 0) {
			ForEach ($UninstallRegKey in $UninstallRegKeys) {
				Try {
					[hashtable]$Return = @{
					}
					$HKLM = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $computer)
					$UninstallRef = $HKLM.OpenSubKey($UninstallRegKey)
					$Applications = $UninstallRef.GetSubKeyNames()
					ForEach ($App in $Applications) {
						$AppRegistryKey = $UninstallRegKey + "\\" + $App
						$AppDetails = $HKLM.OpenSubKey($AppRegistryKey)
						$AppGUID = $App
						$AppDisplayName = $($AppDetails.GetValue("DisplayName"))
						$AppVersion = $($AppDetails.GetValue("DisplayVersion"))
						$AppPublisher = $($AppDetails.GetValue("Publisher"))
						$AppInstalledDate = $($AppDetails.GetValue("InstallDate"))
						$AppUninstall = $($AppDetails.GetValue("UninstallString"))
						If ($UninstallRegKey -match "Wow6432Node") {
							$Softwarearchitecture = "x86"
						} Else {
							$Softwarearchitecture = "x64"
						}
						If (!$AppDisplayName) {
							Continue
						}
						If (($AppDisplayName -like $Product) -and ($AppVersion -like $Version)) {
							$Return.AppDisplayName = $AppDisplayName
							$Return.AppVersion = $AppVersion
							$IsSoftwareInstalled = $True
							$Return.IsSoftwareInstalled = $IsSoftwareInstalled
							Write-Log -Message "  $($MyInvocation.MyCommand.Name)::  ""$AppDisplayName"" ""$AppVersion"""
							Write-Log -Message "  $($MyInvocation.MyCommand.Name)::  ""True"""
							Return $IsSoftwareInstalled
						}
					}
				} Catch {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to read $UninstallRegKey" -Type 3
					Continue
				}
			}
		}
	}
}

<#
	.SYNOPSIS
		Queries WMI for all software installed
	
	.DESCRIPTION
		Queries WMI for all software installed
	
	.PARAMETER ProgressBar
		Variable of Progress Bar in GUI
	
	.EXAMPLE
		PS C:\> Is-SoftwareInstalledWMI
	
	.NOTES
		Returns an array of all software names and versions installed
#>
function Is-SoftwareInstalledWMI {
	param
	(
		[Parameter(Mandatory = $false)]
		$ProgressBar
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	$Win32_Products = Get-WmiObject -query "Select Name,Version from Win32_Product"
	
	$InstalledApplications = @()
	
	foreach ($Product in $Win32_Products) {
		$InstalledApplications += New-Object psobject -Property @{
			Name = $Product.Name;
			Version = $Product.Version
		}
		If ($ProgressBar) {
			$ProgressBar.PerformStep()
			Start-Sleep -Milliseconds 25
		}
		[String]$ProductName = $Product.Name
		[String]$ProductVersion = $Product.Version
		Write-Log -Message "Name:"$ProductName"    Version:"$ProductVersion"" -Type 1
	}
	Return $InstalledApplications
}

 <#'-------------------------------------------------------------------------------
   '---    File System
   '-------------------------------------------------------------------------------#>

<#
	.SYNOPSIS
		Creates a folder
	
	.DESCRIPTION
		Creates a folder at the provided path if it doesn't exist
	
	.PARAMETER Path
		Full path to folder to be created
	
	.EXAMPLE
		Create-Folder -Path "C:\ProgramData\NewFolder"
	
	.NOTES
		Additional information about the function.
#>
function Create-Folder {
	param
	(
		[Parameter(Mandatory = $true)]
		$Path
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Function started"
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Path"
	If (!(Test-Path -Path $Path)) {
		$CreateFolder = New-Item -ItemType directory -Path $Path -Force -ErrorAction SilentlyContinue -ErrorVariable iErr;
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Could not create directory: ""$Path""" -Type 3
		} Else {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Created directory: ""$Path"""
		}
	} Else {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: ""$Path"" already exists"
	}
}

<#
	.SYNOPSIS
		Copies a file
	
	.DESCRIPTION
		Copies a file from a provided source path to a provided destination path
	
	.PARAMETER Source
		Full path to file to be copied
	
	.PARAMETER Destination
		Full path to destination file should be copied to
	
	.EXAMPLE
		Copy-File -Source "C:\Notepad.exe" -Destination "C:\Windows\Notepad.exe"
	
	.NOTES
		Additional information about the function.
#>
function Copy-File {
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$Source,
		[Parameter(Mandatory = $true)]
		[String]$Destination
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Function started"
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Source, $Destination"
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Source: $Source"
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Destination: $Destination"
	
	If (!(Test-Path -Path $Destination)) {
		New-Item -ItemType File -Path $Destination -Force
	}
	
	If (!(Test-Path -Path $Source)) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unsucessful (Invalid Source Path)" -Type 3
	} Else {
		$CopyFile = Copy-Item -Path $Source -Destination $Destination -Force -ErrorAction SilentlyContinue -ErrorVariable iErr;
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unsucessful" -Type 3
		} Else {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Success" -Type 1
		}
	}
}

<#
	.SYNOPSIS
		Copies a folder
	
	.DESCRIPTION
		Copies a folder from a provided source path to a provided destination path
	
	.PARAMETER Source
		Full path to folder to be copied
	
	.PARAMETER Destination
		Full path to destination folder should be copied to
	
	.EXAMPLE
		Copy-Folder -Source "C:\Folder" -Destination "C:\Windows\Folder"
	
	.NOTES
		
#>
Function Copy-Folder {
	Param (
		[Parameter(Mandatory = $True)]
		[String]$Source,
		[Parameter(Mandatory = $True)]
		[String]$Destination
	)
	Write-Log "  $($MyInvocation.MyCommand.Name):: Function started"
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Source, $Destination"
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Source: $Source"
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Destination: $Destination"
	
	If (!(Test-Path -Path $Destination)) {
		Create-Folder $Destination
	}
	
	If (!(Test-Path -Path $Source)) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unsucessful (Invalid Source Path)" -Type 3
	} Else {
		$CopyFolder = Copy-Item -Path $Source -Destination $Destination -Force -Recurse -ErrorAction SilentlyContinue -ErrorVariable iErr;
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unsucessful" -Type 3
		} Else {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Success" -Type 1
		}
	}
}

<#
	.SYNOPSIS
		Deletes a file or folder
	
	.DESCRIPTION
		Deletes a provided file or folder
	
	.PARAMETER Path
		Full path to file or folder to be deleted
	
	.EXAMPLE
		Delete-Object -Path "C:\Notepad.exe"
	
	.NOTES
		
#>
function Delete-Object {
	param
	(
		[Parameter(Mandatory = $true)]
		$Path
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Function started"
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Path"
	
	If (!(Test-Path -Path $Path)) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Path does not exist" -Type 3
	} Else {
		$DeleteObject = Remove-Item $Path -Force -Recurse -ErrorAction SilentlyContinue -ErrorVariable iErr;
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Object Deletion Failed" -Type 3
		} Else {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Object Deletion Completed Successfully"
		}
	}
}


<#
	.SYNOPSIS
		A brief description of the Get-IniContent function.
	
	.DESCRIPTION
		A detailed description of the Get-IniContent function.
	
	.PARAMETER FilePath
		Path to the .ini file to process
		Path is validated that file extension ends in .ini
	
	.EXAMPLE
		Get-IniContent -FilePath "$PSScriptRoot\Config.ini"
	
	.NOTES
		Returns contents of .ini in an array
#>
function Get-IniContent {
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[ValidateScript({
				(Test-Path $_) -and ((Get-Item $_).Extension -eq ".ini")
			})]
		[ValidateNotNullOrEmpty()]
		[string]$FilePath
	)
	
	Begin {
		Write-Log "  $($MyInvocation.MyCommand.Name):: Function started"
	}
	
	Process {
		Write-Log "  $($MyInvocation.MyCommand.Name):: Processing file: $Filepath"
		
		$ini = @{
		}
		Switch -Regex -File $FilePath
		{
			"^\[(.+)\]$" # Section  
			{
				$section = $matches[1]
				$ini[$section] = @{
				}
				$CommentCount = 0
			}
			"^(;.*)$" # Comment  
			{
				If (!($section)) {
					$section = "No-Section"
					$ini[$section] = @{
					}
				}
				$value = $matches[1]
				$CommentCount = $CommentCount + 1
				$name = "Comment" + $CommentCount
				$ini[$section][$name] = $value
			}
			"(.+?)\s*=\s*(.*)" # Key  
			{
				if (!($section)) {
					$section = "No-Section"
					$ini[$section] = @{
					}
				}
				$name, $value = $matches[1 .. 2]
				$ini[$section][$name] = $value
			}
		}
		Write-Log "  $($MyInvocation.MyCommand.Name):: Finished Processing file: $FilePath"
		Return $ini
	}
	
	End {
		Write-Log "  $($MyInvocation.MyCommand.Name):: Function ended"
	}
}

<#  
    .Synopsis  
        Write hash content to INI file  
          
    .Description  
        Write hash content to INI file  
          
    .Notes  
        Author        : Oliver Lipkau <oliver@lipkau.net>  
        Blog        : http://oliver.lipkau.net/blog/  
        Source        : https://github.com/lipkau/PsIni 
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91 
        Version        : 1.0 - 2010/03/12 - Initial release  
                      1.1 - 2012/04/19 - Bugfix/Added example to help (Thx Ingmar Verheij)  
                      1.2 - 2014/12/11 - Improved handling for missing output file (Thx SLDR) 
          
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
        $Category1 = @{"Key1"="Value1";"Key2"="Value2"}  
    	$Category2 = @{"Key1"="Value1";"Key2"="Value2"}  
    	$NewINIContent = @{"Category1"=$Category1;"Category2"=$Category2}  
    	Out-IniFile -InputObject $NewINIContent -FilePath "C:\MyNewFile.INI"  
        -----------  
        Description  
        Creating a custom Hashtable and saving it to C:\MyNewFile.INI  
    .Link  
        Get-IniContent  
    #>
function Out-IniFile {
	[CmdletBinding()]
	param
	(
		[switch]$Append,
		[ValidateSet('Unicode', 'UTF7', 'UTF8', 'UTF32', 'ASCII', 'BigEndianUnicode', 'Default', 'OEM')]
		[string]$Encoding = "Unicode",
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^([a-zA-Z]\:)?.+\.ini$')]
		[ValidateNotNullOrEmpty()]
		[string]$FilePath,
		[switch]$Force,
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true)]
		[ValidateNotNullOrEmpty()]
		[Hashtable]$InputObject,
		[switch]$Passthru
	)
	
	Begin {
		Write-Log "  $($MyInvocation.MyCommand.Name):: Function started"
	}
	
	Process {
		Write-Log "  $($MyInvocation.MyCommand.Name):: Writing to file: $Filepath"
		
		If ($append) {
			$outfile = Get-Item $FilePath
		} Else {
			$outFile = New-Item -ItemType file -Path $Filepath -Force:$Force
		}
		If (!($outFile)) {
			Throw "Could not create File"
		}
		Foreach ($i in $InputObject.keys) {
			If (!($($InputObject[$i].GetType().Name) -eq "Hashtable")) {
				#No Sections  
				Write-Log "  $($MyInvocation.MyCommand.Name):: Writing key: $i"
				Add-Content -Path $outFile -Value "$i=$($InputObject[$i])" -Encoding $Encoding
			} Else {
				#Sections  
				Write-Log "  $($MyInvocation.MyCommand.Name):: Writing Section: [$i]"
				Add-Content -Path $outFile -Value "[$i]" -Encoding $Encoding
				Foreach ($j in $($InputObject[$i].keys | Sort-Object)) {
					If ($j -match "^Comment[\d]+") {
						Write-Log "  $($MyInvocation.MyCommand.Name):: Writing comment: $j"
						Add-Content -Path $outFile -Value "$($InputObject[$i][$j])" -Encoding $Encoding
					} Else {
						Write-Log "  $($MyInvocation.MyCommand.Name):: Writing key: $j"
						Add-Content -Path $outFile -Value "$j=$($InputObject[$i][$j])" -Encoding $Encoding
					}
					
				}
				Add-Content -Path $outFile -Value "" -Encoding $Encoding
			}
		}
		Write-Log "  $($MyInvocation.MyCommand.Name):: Finished Writing to file: $Filepath"
		If ($PassThru) {
			Return $outFile
		}
	}
	
	End {
		Write-Log "  $($MyInvocation.MyCommand.Name):: Function ended"
	}
}



<#'-------------------------------------------------------------------------------
  '---    Operating System
  '-------------------------------------------------------------------------------#>

<#
	.SYNOPSIS
		Is operating system architecture 32 or 64
	
	.DESCRIPTION
		Queries WMI to determine if operating system is 32-bit or 64-bit
	
	.EXAMPLE
		Get-OSArchitecture
	
	.NOTES
		Returns 32 or 64
#>
function Get-OSArchitecture {
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	$GetOSArchitecture = (Get-WmiObject Win32_OperatingSystem -computername $env:COMPUTERNAME).OSArchitecture
	
	If ($GetOSArchitecture -like "*64*") {
		$GetOSArchitecture = "64"
	} Else {
		$GetOSArchitecture = "32"
	}
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: ""$GetOSArchitecture"""
	
	Return $GetOSArchitecture.ToString()
}


<#
	.SYNOPSIS
		Gets name of operating system installed from WMI
	
	.DESCRIPTION
		Gets name of operating system installed from WMI
	
	.EXAMPLE
		Get-OSName
	
	.NOTES
		Returns Operating System Name
#>
function Get-OSName {
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	$wmiOS = Get-WmiObject -ComputerName $ComputerName -Class Win32_OperatingSystem;
	
	$OSName = $wmiOS.Caption
	
	Return $OSName
}


<#'-------------------------------------------------------------------------------
  '---    Registry
  '-------------------------------------------------------------------------------#>


Function Read-Registry {
	Param (
		[Parameter(Mandatory = $True)]
		$Path,
		[Parameter(Mandatory = $True)]
		$Name
	)
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Path\$Name"
	
	If (!(Test-Path -Path $Path)) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Registry Path Not Found" -Type 3
		Return 42
	}
	
	$ReadRegistry = Get-ItemProperty -Path $Path -Name $Key -ErrorAction SilentlyContinue -ErrorVariable iErr | ForEach-Object {
		$_.$Name
	}
	
	If ($iErr) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to read registry" -Type 3
	}
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$ReadRegistry'"
	
	Return $ReadRegistry
}


<#
	.SYNOPSIS
		Writes to a registry path
	
	.DESCRIPTION
		Writes to a registry path
	
	.PARAMETER Path
		Path in the registry to write to
	
	.PARAMETER Name
		Ex "HKLM:\SOFTWARE\Wow6432Node"
	
	.PARAMETER Type
		Ex "Adobe"
	
	.PARAMETER Value
		Ex "String"
	
	.PARAMETER Force
		Ex "1.1.2.0"
	
	.EXAMPLE
				PS C:\> Write-Registry -Path 'Value1' -Name $value2 -Value $value3 -Force $value4
	
	.NOTES
		Additional information about the function.
#>
function Write-Registry {
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$Path,
		[Parameter(Mandatory = $true)]
		$Name,
		[Parameter(Mandatory = $false)]
		$Type,
		[Parameter(Mandatory = $true)]
		$Value,
		[Parameter(Mandatory = $true)]
		$Force
	)
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Path : $Name : $Type : $Value"
	
	If (!(Test-Path -Path $Path)) {
		$CMD = New-Item -Path $Path -ErrorAction SilentlyContinue -ErrorVariable iErr;
		
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to write to registry" -Type 3
			Return.$CMD
		}
	}
	
	If ((Read-Registry -Path $Path -Name $Name) -eq $Value) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Registry Key Already Exists"
		Return
	}
	If ($Force -eq $True) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Registry Key will be forcefully overwritten"
	} Else {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Registry Key will not be overwritten, use -Force $True to forcefully overwrite"
		Return
	}
	
	$WriteRegistry = New-ItemProperty -Path $Path -Name $Name -PropertyType $Type -Value $Value -Force
	
	If ((Get-ItemProperty $Path -Name $Name -ErrorAction SilentlyContinue).$Name) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Registry key Written successfully"
	} Else {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to write registry key"
		Return.$WriteRegistry
	}
}



<#'-------------------------------------------------------------------------------
 '---    Active Directory
 '-------------------------------------------------------------------------------#>


Function AD_ManageGroup {
	Param (
		[Parameter(Mandatory = $False)]
		$Domain,
		[Parameter(Mandatory = $True)]
		$Function,
		[Parameter(Mandatory = $True)]
		$Type,
		[Parameter(Mandatory = $True)]
		$Name,
		[Parameter(Mandatory = $True)]
		$Group,
		[Parameter(Mandatory = $False)]
		$ADUser,
		[Parameter(Mandatory = $False)]
		$ADPass
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Domain', '$Function', '$Type', '$Name', '$Group', '$ADUser', '$ADPass'"
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Importing Active Directory Module"
	
	$ImportModule = (Import-Module ActiveDirectory -PassThru -ErrorAction SilentlyContinue -ErrorVariable iErr).ExitCode
	
	If ($iErr) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to import AD Module Error: $ImportModule" -type 3
		Return
	} Else {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Imported AD Module"
	}
	
	If ($Type -eq "User") {
		$GetUser = Get-ADUser -Identity $Name -Properties MemberOf, sAMAccountName -ErrorAction SilentlyContinue -ErrorVariable iErr | Select-Object MemberOf, sAMAccountName
		
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to locate $Type : '$Name' in AD" -Type 3
			Return
		}
		
		If ($Function -eq "Add") {
			If ($GetUser.MemberOf -match $Group) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is already a member of '$Group'"
				Return
			} Else {
				$CMD = Add-ADGroupMember -Identity "$Group" -Members "$Name" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
				If ($iErr) {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Adding '$Name' to '$Group' failed" -type 3
					Return
				} Else {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Added '$Name' to '$Group' successfully"
				}
			}
		} ElseIf ($Function -eq "Remove") {
			If (!($GetUser.MemberOf -match $Group)) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is already not a member of '$Group'"
				Return
			}
			$CMD = Remove-ADGroupMember -Identity "$Group" -Members "$Name" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
			If ($iErr) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Removing '$Name' from '$Group' failed" -type 3
				Return
			} Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Removed '$Name' from '$Group' successfully"
			}
		} ElseIf ($Function -eq "Query") {
			If ($GetUser.MemberOf -match $Group) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is a member of '$Group'"
				Return
			} Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is NOT a member of '$Group'"
			}
		}
	}
	
	
	If ($Type -eq "Computer") {
		$GetComputer = Get-ADComputer $Name -Properties MemberOf -ErrorAction SilentlyContinue -ErrorVariable iErr
		
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to locate $Type : '$Name' in AD" -Type 3
			Return
		}
		
		
		If ($Function -eq "Add") {
			If ($GetComputer.MemberOf -match $Group) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is already a member of '$Group'"
				Return
			} Else {
				$CMD = Add-ADGroupMember -Identity "$Group" -Members "$Name$" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
				If ($iErr) {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Adding '$Name' to '$Group' failed" -type 3
					Return
				} Else {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Added '$Name' to '$Group' successfully"
				}
			}
		} ElseIf ($Function -eq "Remove") {
			If (!($GetComputer.Memberof -match $Group)) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is already not a member of '$Group'"
				Return
			}
			$CMD = Remove-ADGroupMember -Identity "$Group" -Members "$Name$" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
			If ($iErr) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Removing '$Name' from '$Group' failed" -type 3
				Return
			} Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Removed '$Name' from '$Group' successfully"
			}
		} ElseIf ($Function -eq "Query") {
			If ($GetComputer.Memberof -Match $Group) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is a member of '$Group'"
				Return
			} Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is NOT a member of '$Group'"
			}
		}
	}
	
	
	If ($Type -eq "Group") {
		$GetGroup = (Get-ADGroup -Identity $Name -Properties MemberOf -ErrorAction SilentlyContinue -ErrorVariable iErr | Select-Object MemberOf)
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to locate $Type : '$Name' in AD" -Type 3
			Return
		}
		
		If ($Function -eq "Add") {
			If ($GetGroup.MemberOf -match $Group) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is already a member of '$Group'"
				Return
			} Else {
				$CMD = Add-ADGroupMember -Identity "$Group" -Members "$Name" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
				If ($iErr) {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Adding '$Name' to '$Group' failed" -type 3
					Return
				} Else {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Added '$Name' to '$Group' successfully"
				}
			}
		} ElseIf ($Function -eq "Remove") {
			If (!($GetGroup.MemberOf -match $Group)) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is already not a member of '$Group'"
				Return
			}
			$CMD = Remove-ADGroupMember -Identity "$Group" -Members "$Name" -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable iErr;
			If ($iErr) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Removing '$Name' from '$Group' failed" -type 3
				Return
			} Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Removed '$Name' from '$Group' successfully"
			}
		} ElseIf ($Function -eq "Query") {
			If ($GetGroup.MemberOf -match $Group) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is a member of '$Group'"
				Return
			} Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: '$Name' is NOT a member of '$Group'"
			}
		}
	}
}


<#
  .SYNOPSIS
    Adds/Removes Users/Computers/Groups To/From Groups
  .DESCRIPTION
    Does not require the Active Directory module and allows you to pass alternate credentials to run under.
    sFunction Options Are: Add or Remove
    sType Options Are: User, Computer, or Group
    sADUser and sADPass are optional and are only needed when alternate credentials need supplied
  .EXAMPLE
    AD_ManageGroupADSI -Domain "contoso.org" -Function "Add" -Type "User" -Name "abc20a" -Group "GROUP_1"
  .EXAMPLE
    AD_ManageGroupADSI -Domain "contoso.org" -Function "Add" -Type "Computer" -Name "DT12345678" -Group "GROUP_1" -ADUser "abc123a" -ADPass "P@ssw0rd"
  .EXAMPLE
    AD_ManageGroupADSI -Domain "contoso.org" -Function "Remove" -Type "Computer" -Name "DT12345678" -Group "GROUP_1" -ADUser "abc123a" -ADPass "P@ssw0rd"
  .EXAMPLE
    AD_ManageGroupADSI -Domain "contoso.org" -Function "Add" -Type "User" -Name "czt20b" -Group "GROUP_1" -ADUser "abc123a" -ADPass "P@ssw0rd"
  .EXAMPLE
    AD_ManageGroupADSI -Domain "contoso.org" -Function "Add" -Type "Group" -Name "GROUP_42" -Group "GROUP_1" -ADUser "abc123a" -ADPass "P@ssw0rd"
  #>
Function AD_ManageGroupADSI {

	Param (
		[Parameter(Mandatory = $False)]
		$Domain,
		[Parameter(Mandatory = $True)]
		$Function,
		[Parameter(Mandatory = $True)]
		$Type,
		[Parameter(Mandatory = $True)]
		$Name,
		[Parameter(Mandatory = $True)]
		$Group,
		[Parameter(Mandatory = $False)]
		$ADUser,
		[Parameter(Mandatory = $False)]
		$ADPass
	)
	
	[int]$ADS_PROPERTY_APPEND = 3
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: Starting Function"
	
	$GroupPath = Get-ADSPath -DomainName $Domain -Type "Group" -Name $Group -ADUser $ADUser -ADPass $ADPass
	$ObjectPath = Get-ADSPath -DomainName $Domain -Type $Type -Name $Name -ADUser $ADUser -ADPass $ADPass
	
	If ($GroupPath -eq $Null) {
		Return
	}
	If ($ObjectPath -eq $Null) {
		Return
	}
	
	$ObjectDN = $ObjectPath.adspath.Replace("$Domain/", "")
	#$ObjectDN = $ObjectPath.Replace("$Domain/", "")
	#Write-Log -Message "  DN: $ObjectDN"
	
	$ObjectCN = $ObjectPath.adspath.Replace("LDAP://$Domain/", "")
	#$ObjectCN = $ObjectPath.Replace("LDAP://$Domain/", "")
	#Write-Log -Message "  CN: $ObjectCN"
	
	$GroupDN = $GroupPath.adspath.Replace("$Domain/", "")
	#$GroupDN = $GroupPath.Replace("$Domain/", "")
	#Write-Log -Message "  DN: $GroupDN"
	
	If ($ADUser -and $ADPass) {
		$oGroup = New-Object DirectoryServices.DirectoryEntry($GroupDN, $ADUser, $ADPass)
	} Else {
		$oGroup = [ADSI]$GroupDN
	}
	
	$oComputer = [ADSI]$ObjectDN
	
	If ($Function -eq "Add") {
		Try {
			#Verify if the computer is a member of the Group
			If ($oGroup.ismember($oComputer.adspath) -eq $False) {
				#Add the the computer to the specified group
				$oGroup.PutEx($ADS_PROPERTY_APPEND, "member", @("$ObjectCN"))
				$oGroup.setinfo()
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Added $Name to $Group"
			} Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Name is already a member of $Group"
			}
		} Catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name)::  Unable to query $Group for membership status. If credentials were passed check that credentials are valid" -Type 3
			Return
		}
	}
	
	If ($Function -eq "Remove") {
		Try {
			#Verify if the computer is a member of the Group
			If ($oGroup.ismember($oComputer.adspath) -eq $False) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Name is not a member of $Group"
			} Else {
				#Add the the computer to the specified group
				$oGroup.Member.Remove($ObjectCN)
				$oGroup.setinfo()
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Removed $Name from $Group"
			}
		} Catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to query $Group for membership status. If credentials were passed check that credentials are valid" -Type 3
			Return
		}
	}
} #End Function

<#****** NEED TO FIX *********#
Function AD_ManageComputers {
	Param (
		[Parameter(Mandatory = $True)]
		$Path,
		[Parameter(Mandatory = $False)]
		$Function,
		[Parameter(Mandatory = $False)]
		$DaysInactive
	)
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Path, $Function, $DaysInactive"
	
	$ImportModule = (Import-Module ActiveDirectory -PassThru -ErrorAction SilentlyContinue -ErrorVariable iErr).ExitCode
	
	If ($iErr) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to import AD Module Error: $ImportModule" -type 3
		Return
	} Else {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Imported AD Module"
	}
	
	If ($Function -eq "Disable") {
		
		$Machines = Get-Content -Path $Path -ErrorAction SilentlyContinue -ErrorVariable iErr;
		
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to locate or open $Path"
			Return
		} Else {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Machine List Loaded Sucessfully."
		}
		
		$Machines | foreach {
			Try {
				$DisablePC = Get-ADComputer -Identity $_ -ErrorAction SilentlyContinue | Disable-ADAccount -Confirm:$false -ErrorAction SilentlyContinue
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $_ Sucessfully Disabled in Active Directory."
			} Catch {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: ($_) Does not exist in Active Directory." -Type 3
			}
		}
	}
	
	If ($Function -eq "Delete") {
		
		$Machines = Get-Content -Path $Path -ErrorAction SilentlyContinue -ErrorVariable iErr;
		
		If ($iErr) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to locate or open $Path"
			Return
		} Else {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Machine List Loaded Sucessfully."
		}
		
		$Machines | ForEach {
			Try {
				$DisablePC = Get-ADComputer -Identity $_ -ErrorAction SilentlyContinue | Delete-ADAccount -Confirm:$false -ErrorAction SilentlyContinue
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $_ Sucessfully Disabled in Active Directory."
			} Catch {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: ($_) Does not exist in Active Directory." -Type 3
			}
		}
	}
	
	If ($Function -eq "QueryInactive") {
		
		$Time = (Get-Date).Adddays(- ($DaysInactive))
		
		# Get all AD computers with lastLogonTimestamp less than our time
		Get-ADComputer -Server "ghs.org" -Filter {
			LastLogonTimeStamp -lt $Time
		} -Properties LastLogonTimeStamp |
		
		# Output hostname and lastLogonTimestamp into CSV
		select-object Name | export-csv $Path -notypeinformation
	}
}#>



<#
	.SYNOPSIS
		Returns object from Active Directory
	
	.DESCRIPTION
		Returns a User, Group, or Computer object from Active Directory
		Great for connecting to AD with hardcoded credentials, running as invoked user without prompting for credentials,  or when Active directory module is unavailable
	
	.PARAMETER DomainName
		FQDN
	
	.PARAMETER Name
		Name of the object to query in AD
	
	.PARAMETER Type
		Type of object:
		User
		Computer
		Group
	
	.PARAMETER ADUser
		Username to connect to Active Directory
	
	.PARAMETER ADPass
		Password for ADUser paramater to connect to Active Directory
	
	.EXAMPLE
		Get-ADSPath -DomainName "Contoso.com" -Name "PC0001" -Type "Computer"
		Get-ADSPath -DomainName "Contoso.com" -Name "PC0001" -Type "Computer" -ADUser "BPH123" -ADPass "P@ssw0rd"		
	
	.NOTES
		Returns an object from AD
		Can be run from Workgroup computer with credentials supplied but,
		functions that call this function may not be able to run from Workgroup computer
#>
function Get-ADSPath {
	param
	(
		[Parameter(Mandatory = $false)]
		[String]$DomainName = $ComputerDomain,
		[Parameter(Mandatory = $true)]
		[String]$Name,
		[Parameter(Mandatory = $true)]
		[String]$Type,
		[Parameter(Mandatory = $false)]
		[String]$ADUser = "",
		[Parameter(Mandatory = $false)]
		[String]$ADPass = ""
	)
	
	If ($ADUser -and $ADPass) {
		$Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry("LDAP://$DomainName", $ADUser, $ADPass)
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Credentials supplied for: $ADUser"
	} Else {
		$Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry("LDAP://$DomainName")
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Running as invoked user"
	}
	
	If ($Type -eq "User") {
		$Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher($Domain, "(&(objectCategory=User)(sAMAccountname=$Name))")
		$Searcher.SearchScope = "Subtree"
		$Searcher.SizeLimit = '5000'
		$ADOQuery = $Searcher.FindAll()
		$ADSPath = $ADOQuery.Path
		$ADOProperties = $ADOQuery.Properties
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: User: $Name ADSPath: $ADSPath"
		Return $ADOProperties
	} ElseIf ($Type -eq "Group") {
		$Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher($Domain, "(&(objectCategory=Group)(name=$Name))")
		$Searcher.SearchScope = "Subtree"
		$Searcher.SizeLimit = '5000'
		$ADOQuery = $Searcher.FindAll().Path
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Group: $Name ADSPath: $ADOQuery"
		Return [ADSI]$ADOQuery
	} ElseIf ($Type -eq "Computer") {
		$Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher($Domain, "(&(objectCategory=Computer)(name=$Name))")
		$Searcher.SearchScope = "Subtree"
		$Searcher.SizeLimit = '5000'
		$ADOQuery = $Searcher.FindAll().Path
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Computer: $Name ADSPath: $ADOQuery"
		Return [ADSI]$ADOQuery
	}
}



<#
	.SYNOPSIS
		Used to change a domain user's password with using the Active Directory Module
	
	.DESCRIPTION
		Used to change a domain user's password with using the Active Directory Module
		Cannot be run from a Workgroup computer
	
	.PARAMETER Domain
		FQDN of the user
	
	.PARAMETER User
		Domain User whose password should be changed
	
	.PARAMETER NewPassword
		New password to be set for domain user account
	
	.PARAMETER ADUser
		Username to connect to Active Directory
	
	.PARAMETER ADPass
		Password for ADUser paramater to connect to Active Directory
	
	.EXAMPLE
				PS C:\> Set-AdUserPasswordADSI -Domain 'Value1' -User 'Value2'
	
	.NOTES
		Additional information about the function.
#>
function Set-ADUserPasswordADSI {
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$Domain,
		[Parameter(Mandatory = $true)]
		[String]$User,
		[Parameter(Mandatory = $false)]
		[String]$NewPassword,
		[Parameter(Mandatory = $false)]
		[String]$ADUser,
		[Parameter(Mandatory = $false)]
		[String]$ADPass
	)
	
	$oUser = Get-ADSPath -DomainName $Domain -Name $User -Type "User" -ADUser $ADUser -ADPass $ADPass
	$oUserDN = $oUser.distinguishedname
	$oUserFullDN = [ADSI]"LDAP://$oUserDN"
	$oUserFullDN.psbase.invoke("SetPassword", $NewPassword)
	$oUserFullDN.psbase.CommitChanges()
} # end function Set-ADUserPasswordADSI


Function Get-GroupMembershipADSI {
	Param (
		[Parameter(Mandatory = $True)]
		[String]$Domain,
		[Parameter(Mandatory = $True)]
		[String]$Name,
		[Parameter(Mandatory = $True)]
		[String]$Type,
		[Parameter(Mandatory = $False)]
		[String]$ADUser,
		[Parameter(Mandatory = $False)]
		[String]$ADPass
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name):: $Name"
	
	$oADObject = Get-ADSPath -DomainName $Domain -Name $Name -Type $Type -ADUser $ADUser -ADPass $ADPass
	$Groups = $oADObject.memberof | ForEach-Object {
		[ADSI]"LDAP://$_"
	}
	
	Return $Groups
}

<#
	.SYNOPSIS
		Gets the AD groups a user is a member of
	
	.DESCRIPTION
		Gets the AD groups a user is a member of
	
	.PARAMETER User
		Username to get group memberships for
	
	.EXAMPLE
				PS C:\> Get-GroupMembership -User $value1
	
	.NOTES
		Additional information about the function.
#>
function Get-GroupMembership {
	param
	(
		[Parameter(Mandatory = $true)]
		$User
	)
	
	ForEach ($U in $User) {
		$UN = Get-ADUser $U -Properties MemberOf
		$Groups = ForEach ($Group in ($UN.MemberOf)) {
			(Get-ADGroup $Group).Name
		}
		$Groups = $Groups | Sort
		ForEach ($Group in $Groups) {
			New-Object PSObject -Property @{
				Name = $UN.Name
				Group = $Group
			}
		}
	}
}


<#'-------------------------------------------------------------------------------
  '---    SCCM
  '-------------------------------------------------------------------------------#>

<#
	.SYNOPSIS
		Determine if system is currently running a task sequence
	
	.DESCRIPTION
		Determine if system is currently running a task sequence
	
	.EXAMPLE
				PS C:\> Is-InTS
	
	.NOTES
		Additional information about the function.
#>
function Is-InTS {
	Try {
		$Global:oENV = New-Object -COMObject Microsoft.SMS.TSEnvironment
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $False"
		Return $False
	}
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $True"
	Return $True
}


<#
	.SYNOPSIS
		Hide the Task Sequence progress window
	
	.DESCRIPTION
		Hide the Task Sequence progress window
	
	.EXAMPLE
				PS C:\> Hide-TSProgress
	
	.NOTES
		Additional information about the function.
#>
function Hide-TSProgress {
	Write-Log -Message "  $($MyInvocation.MyCommand.Name)::"
	Try {
		$TSProgressUI = New-Object -COMObject Microsoft.SMS.TSProgressUI
		$TSProgressUI.CloseProgressDialog()
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Hid TS Progress Window"
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to hide TS Progress Window" -Type 3
	}
}


<#
	.SYNOPSIS
		Gets the value of a Task Sequence variable
	
	.DESCRIPTION
		Gets the value of a Task Sequence variable
	
	.PARAMETER Variable
		Task Sequence Variable to get the value of
	
	.EXAMPLE
				PS C:\> Get-TSVar -Variable 'Value1'
	
	.NOTES
		Additional information about the function.
#>
function Get-TSVar {
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$Variable
	)
	
	Write-Log "  $($MyInvocation.MyCommand.Name)::$Variable"
	
	Try {
		$TSENV = New-Object -COMObject Microsoft.SMS.TSEnvironment
		$TSVariable = $TSENV.Value($Variable)
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Variable = $TSVariable"
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to get value of $Variable" -Type 3
	}
	
	Return $TSVariable
}

<#
	.SYNOPSIS
		Sets a Task Sequence variable during OSD
	
	.DESCRIPTION
		Sets a Task Sequence variable during OSD
	
	.PARAMETER Variable
		Task Sequence Variable to set a value for in OSD
	
	.PARAMETER Value
		Value of the Task Sequence Variable to be set
	
	.EXAMPLE
				PS C:\> Set-TSVar -Variable 'Value1' -Value 'Value2'
	
	.NOTES
		Additional information about the function.
#>
function Set-TSVar {
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$Variable,
		[Parameter(Mandatory = $true)]
		[String]$Value
	)
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name)::$Variable=$Value"
	
	Try {
		$oENV = New-Object -COMObject Microsoft.SMS.TSEnvironment
	} Catch {
	}
	
	Try {
		$oENV.Value($Variable) = $Value
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to Set Variable" -Type 3
		Return
	}
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Successfully Set Variable"
}

<#
	.SYNOPSIS
		Adds a computer to a collection in SCCM
	
	.DESCRIPTION
		Adds a computer to a collection in SCCM
	
	.PARAMETER ComputerName
		Computer record name to add to a collection
	
	.PARAMETER CollectionID
		Collection ID of the collection to add computer to
	
	.PARAMETER CollectionName
		Name of the collection in SCCM to add computer to
	
	.PARAMETER SMSServer
		FQDN of SCCM server to connect to to perform operation
	
	.EXAMPLE
				PS C:\> Add-ComputerToCollection -ComputerName $value1 -CollectionID $value2 -CollectionName $value3 -SMSServer $value4
	
	.NOTES
		Additional information about the function.
#>
function Add-ComputerToCollection {
	param
	(
		[Parameter(Mandatory = $true)]
		$ComputerName,
		[Parameter(Mandatory = $true)]
		$CollectionID,
		[Parameter(Mandatory = $true)]
		$CollectionName,
		[Parameter(Mandatory = $true)]
		$SMSServer
	)
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $ComputerName,$CollectionID,$CollectionName,$SMSServer"
	
	$RulesToSkip = $null
	$strMessage = "Do you want to add '$ComputerName' to '$CollectionName'"
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Connecting to Site Server: $SMSServer"
	Try {
		$sccmProviderLocation = Get-WmiObject -query "select * from SMS_ProviderLocation where ProviderForLocalSite = true" -Namespace "root\sms" -computername $SMSServer
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to connect to Site Server: $SMSServer"
		Return
	}
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Successfully connected to Site Server: $SMSServer"
	$SiteCode = $sccmProviderLocation.SiteCode
	$Namespace = "root\sms\site_$SiteCode"
	
	Write-Log -Message "  Query $SMSServer for CollectionID: $CollectionID"
	$strQuery = "Select * from SMS_Collection where CollectionID = '$CollectionID'"
	$Collection = Get-WmiObject -query $strQuery -ComputerName $SMSServer -Namespace $Namespace
	$Collection.Get()
	
	If ($ComputerName -ne $null) {
		$strQuery = "Select * from SMS_R_System where Name like '$ComputerName'"
		Get-WmiObject -Query $strQuery -Namespace $Namespace -ComputerName $SMSServer | ForEach-Object {
			$ResourceID = $_.ResourceID
			$RuleName = $_.Name
			$ComputerName = $RuleName
			If ($ResourceID -ne $null) {
				$Error.Clear()
				$Collection = [WMI]"\\$($SMSServer)\$($Namespace):SMS_Collection.CollectionID='$CollectionID'"
				$RuleClass = [wmiclass]"\\$($SMSServer)\$($NameSpace):SMS_CollectionRuleDirect"
				$newRule = $ruleClass.CreateInstance()
				$newRule.RuleName = $RuleName
				$newRule.ResourceClassName = "SMS_R_System"
				$newRule.ResourceID = $ResourceID
				$Collection.AddMembershipRule($newRule)
				If ($Error[0]) {
					Write-Log -Message "Error adding $ComputerName - $Error"
					$ErrorMessage = "$Error"
					$ErrorMessage = $ErrorMessage.Replace("`n", "")
				} Else {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Successfully added $ComputerName"
					Return $True
				}
			} Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Could not find $ComputerName - No rule added" -Type 2
			}
		} #End For-Each
		If ($ResourceID -eq $null) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Could not find $ComputerName - No rule added" -Type 2
		}
	}
}

<#
	.SYNOPSIS
		Removes a computer record from SCCM
	
	.DESCRIPTION
		Removes a computer record from SCCM
	
	.PARAMETER SMSComputerName
		Computer record name to be removed
	
	.PARAMETER SMSServer
		FQDN of SMS server to connect to
	
	.PARAMETER Credentials
		Credentials to use to connect to server and perform operations
	
	.PARAMETER StatusOutPut
		Boolean
		Output log info to Status Output Box
	
	.PARAMETER StatusOutPutTextBox
		Variable of textbox in GUI to write status to
	
	.EXAMPLE
				PS C:\> Remove-CMComputer -SMSComputerName $value1 -SMSServer $value2
	
	.NOTES
		Additional information about the function.
#>
function Remove-CMComputer {
	param
	(
		[Parameter(Mandatory = $true)]
		[Array]$SMSComputerName,
		[Parameter(Mandatory = $true)]
		$SMSServer,
		[Parameter(Mandatory = $false)]
		$Credentials,
		[Parameter(Mandatory = $false)]
		[Boolean]$StatusOutPut,
		[Parameter(Mandatory = $false)]
		$StatusOutPutTextBox
	)
	
	$StartDate = ((get-date).toShortDateString() + " " + (get-date).toShortTimeString())
	$global:ScriptStart = get-date
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $SMSComputerName,$SMSServer"
	
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Connecting to Site Server: $SMSServer"
	
	Try {
		If ($Credentials) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Using Supplied Credentials"
			$sccmProviderLocation = Get-WmiObject -query "select * from SMS_ProviderLocation where ProviderForLocalSite = true" -Namespace "root\sms" -computername $SMSServer -Credential $Credentials
		} Else {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Using Invoked Credentials"
			$sccmProviderLocation = Get-WmiObject -query "select * from SMS_ProviderLocation where ProviderForLocalSite = true" -Namespace "root\sms" -computername $SMSServer
		}
	} Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to connect to Site Server: $SMSServer"
		If ($StatusOutPut -eq $true) {
			Write-StatusOutput -Textbox $StatusOutPutTextBox -Text ("-" * 10 + "  Unable to connect to Site Server: $SMSServer " + $StartDate + " " + "-" * 10) -Append $true
		}
		Return
	}
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Successfully connected to Site Server: $SMSServer"
	If ($StatusOutPut -eq $true) {
		Write-StatusOutput -Textbox $StatusOutPutTextBox -Text ("-" * 10 + "  Successfully connected to Site Server: $SMSServer " + $StartDate + " " + "-" * 10) -Append $true
	}
	$SiteCode = $sccmProviderLocation.SiteCode
	$Namespace = "root\sms\site_$SiteCode"
	
	# Get Resource ID
	ForEach ($SMSComputer in $SMSComputerName) {
		
		Try {
			If ($Credentials) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Using Supplied Credentials"
				$ComputerObject = Get-WmiObject -ComputerName $SMSServer -Namespace "root\sms\site_$($SiteCode)" -Class 'SMS_R_SYSTEM' -Filter "Name='$SMSComputer'" -Credential $Credentials
			} Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Using Invoked Credentials"
				$ComputerObject = Get-WmiObject -ComputerName $SMSServer -Namespace "root\sms\site_$($SiteCode)" -Class 'SMS_R_SYSTEM' -Filter "Name='$SMSComputer'"
			}
		} Catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed To Query SCCM"
			If ($StatusOutPut -eq $true) {
				Write-StatusOutput -Textbox $StatusOutPutTextBox -Text ("-" * 10 + "  Failed To Query SCCM " + $StartDate + " " + "-" * 10) -Append $true -Color 'Red'
			}
		}
		
		If ($ComputerObject.ResourceId -eq $NULL) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Computer: $SMSComputer has not been found in SCCM."
			If ($StatusOutPut -eq $true) {
				Write-StatusOutput -Textbox $StatusOutPutTextBox -Text ("-" * 10 + "  Computer: $SMSComputer has not been found in SCCM. " + $StartDate + " " + "-" * 10) -Append $true -Color 'Red'
			}
		}
		
		# Delete computer account
		Try {
			$ComputerObject.PSBase.Delete()
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Computer: $SMSComputer successfully deleted"
			If ($StatusOutPut -eq $true) {
				Write-StatusOutput -Textbox $StatusOutPutTextBox -Text ("-" * 10 + "  Computer: $SMSComputer successfully deleted " + $StartDate + " " + "-" * 10) -Append $true
			}
		} Catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Could not delete $SMSComputer."
			If ($StatusOutPut -eq $true) {
				Write-StatusOutput -Textbox $StatusOutPutTextBox -Text ("-" * 10 + "  Could not delete $SMSComputer. " + $StartDate + " " + "-" * 10) -Append $true -Color 'Red'
			}
		}
	}
}

<#
	.SYNOPSIS
		A brief description of the Write-StatusOutput function.
	
	.DESCRIPTION
		A detailed description of the Write-StatusOutput function.
	
	.PARAMETER Textbox
		Variable of textbox from GUI to write to
	
	.PARAMETER Text
		Text to write to status output
	
	.PARAMETER Append
		Boolean
		Should text append the current text or replace it
	
	.PARAMETER Tabbed
		Should text be tab indented
	
	.PARAMETER Color
		Color of output text
	
	.EXAMPLE
		Write-StatusOutput -Textbox $value1 -Text $value2
	
	.NOTES
		Additional information about the function.
#>
function Write-StatusOutput {
	param
	(
		[Parameter(Mandatory = $true)]
		$Textbox,
		[Parameter(Mandatory = $true)]
		$Text,
		[Parameter(Mandatory = $false)]
		[Boolean]$Append = $True,
		[Parameter(Mandatory = $false)]
		[Boolean]$Tabbed = $False,
		$Color = 'Black'
	)
	
	$Textbox.SelectionColor = $Color
	
	If ($Append) {
		If ($tabbed) {
			
			$Textbox.AppendText("            " + $Text + "`n")
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Text" -Type 1
		}
		Else {
	
			$Textbox.AppendText($Text + "`n")
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Text" -Type 1
		}
	}
	Else {
		$Textbox.AppendText($Text)
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $Text" -Type 1
	}
}

function Get-XmlNamespaceManager {
	param
	(
		[xml]$XmlDocument,
		[string]$NamespaceURI = ""
	)
	
	# If a Namespace URI was not given, use the Xml document's default namespace.
	if ([string]::IsNullOrEmpty($NamespaceURI)) {
		$NamespaceURI = $XmlDocument.DocumentElement.NamespaceURI
	}
	
	# In order for SelectSingleNode() to actually work, we need to use the fully qualified node path along with an Xml Namespace Manager, so set them up.
	[System.Xml.XmlNamespaceManager]$xmlNsManager = New-Object System.Xml.XmlNamespaceManager($XmlDocument.NameTable)
	$xmlNsManager.AddNamespace("ns", $NamespaceURI)
	return, $xmlNsManager # Need to put the comma before the variable name so that PowerShell doesn't convert it into an Object[].
}

function Get-FullyQualifiedXmlNodePath {
	param
	(
		[string]$NodePath,
		[string]$NodeSeparatorCharacter = '.'
	)
	
	return "/ns:$($NodePath.Replace($($NodeSeparatorCharacter), '/ns:'))"
}

function Get-XmlNode {
	param
	(
		[xml]$XmlDocument,
		[string]$NodePath,
		[string]$NamespaceURI = "",
		[string]$NodeSeparatorCharacter = '.'
	)
	
	$xmlNsManager = Get-XmlNamespaceManager -XmlDocument $XmlDocument -NamespaceURI $NamespaceURI
	[string]$fullyQualifiedNodePath = Get-FullyQualifiedXmlNodePath -NodePath $NodePath -NodeSeparatorCharacter $NodeSeparatorCharacter
	
	# Try and get the node, then return it. Returns $null if the node was not found.
	$node = $XmlDocument.SelectSingleNode($fullyQualifiedNodePath, $xmlNsManager)
	return $node
}

function Get-XmlNodes {
	param
	(
		[xml]$XmlDocument,
		[string]$NodePath,
		[string]$NamespaceURI = "",
		[string]$NodeSeparatorCharacter = '.'
	)
	
	$xmlNsManager = Get-XmlNamespaceManager -XmlDocument $XmlDocument -NamespaceURI $NamespaceURI
	[string]$fullyQualifiedNodePath = Get-FullyQualifiedXmlNodePath -NodePath $NodePath -NodeSeparatorCharacter $NodeSeparatorCharacter
	
	# Try and get the nodes, then return them. Returns $null if no nodes were found.
	$nodes = $XmlDocument.SelectNodes($fullyQualifiedNodePath, $xmlNsManager)
	return $nodes
}

function Get-XmlElementsTextValue {
	param
	(
		[xml]$XmlDocument,
		[string]$ElementPath,
		[string]$NamespaceURI = "",
		[string]$NodeSeparatorCharacter = '.'
	)
	
	# Try and get the node.	
	$node = Get-XmlNode -XmlDocument $XmlDocument -NodePath $ElementPath -NamespaceURI $NamespaceURI -NodeSeparatorCharacter $NodeSeparatorCharacter
	
	# If the node already exists, return its value, otherwise return null.
	if ($node) {
		return $node.InnerText
	} else {
		return $null
	}
}

function Set-XmlElementsTextValue {
	param
	(
		[xml]$XmlDocument,
		[string]$ElementPath,
		[string]$TextValue,
		[string]$NamespaceURI = "",
		[string]$NodeSeparatorCharacter = '.'
	)
	
	# Try and get the node.	
	$node = Get-XmlNode -XmlDocument $XmlDocument -NodePath $ElementPath -NamespaceURI $NamespaceURI -NodeSeparatorCharacter $NodeSeparatorCharacter
	
	# If the node already exists, update its value.
	if ($node) {
		$node.InnerText = $TextValue
	}
	# Else the node doesn't exist yet, so create it with the given value.
	else {
		# Create the new element with the given value.
		$elementName = $ElementPath.Substring($ElementPath.LastIndexOf($NodeSeparatorCharacter) + 1)
		$element = $XmlDocument.CreateElement($elementName, $XmlDocument.DocumentElement.NamespaceURI)
		$textNode = $XmlDocument.CreateTextNode($TextValue)
		$element.AppendChild($textNode) > $null
		
		# Try and get the parent node.
		$parentNodePath = $ElementPath.Substring(0, $ElementPath.LastIndexOf($NodeSeparatorCharacter))
		$parentNode = Get-XmlNode -XmlDocument $XmlDocument -NodePath $parentNodePath -NamespaceURI $NamespaceURI -NodeSeparatorCharacter $NodeSeparatorCharacter
		
		if ($parentNode) {
			$parentNode.AppendChild($element) > $null
		} else {
			throw "$parentNodePath does not exist in the xml."
		}
	}
}

function Get-XmlElementsAttributeValue {
	param
	(
		[xml]$XmlDocument,
		[string]$ElementPath,
		[string]$AttributeName,
		[string]$NamespaceURI = "",
		[string]$NodeSeparatorCharacter = '.'
	)
	
	# Try and get the node. 
	$node = Get-XmlNode -XmlDocument $XmlDocument -NodePath $ElementPath -NamespaceURI $NamespaceURI -NodeSeparatorCharacter $NodeSeparatorCharacter
	
	# If the node already exists, return its value, otherwise return null.
	if ($node -and $node.$AttributeName) {
		return $node.$AttributeName
	} else {
		return $null
	}
}

function Set-XmlElementsAttributeValue {
	param
	(
		[xml]$XmlDocument,
		[string]$ElementPath,
		[string]$AttributeName,
		[string]$AttributeValue,
		[string]$NamespaceURI = "",
		[string]$NodeSeparatorCharacter = '.'
	)
	
	# Try and get the node. 
	$node = Get-XmlNode -XmlDocument $XmlDocument -NodePath $ElementPath -NamespaceURI $NamespaceURI -NodeSeparatorCharacter $NodeSeparatorCharacter
	
	# If the node already exists, create/update its attribute's value.
	if ($node) {
		$attribute = $XmlDocument.CreateNode([System.Xml.XmlNodeType]::Attribute, $AttributeName, $NamespaceURI)
		$attribute.Value = $AttributeValue
		$node.Attributes.SetNamedItem($attribute) > $null
	}
	# Else the node doesn't exist yet, so create it with the given attribute value.
	else {
		# Create the new element with the given value.
		$elementName = $ElementPath.SubString($ElementPath.LastIndexOf($NodeSeparatorCharacter) + 1)
		$element = $XmlDocument.CreateElement($elementName, $XmlDocument.DocumentElement.NamespaceURI)
		$element.SetAttribute($AttributeName, $NamespaceURI, $AttributeValue) > $null
		
		# Try and get the parent node.
		$parentNodePath = $ElementPath.SubString(0, $ElementPath.LastIndexOf($NodeSeparatorCharacter))
		$parentNode = Get-XmlNode -XmlDocument $XmlDocument -NodePath $parentNodePath -NamespaceURI $NamespaceURI -NodeSeparatorCharacter $NodeSeparatorCharacter
		
		if ($parentNode) {
			$parentNode.AppendChild($element) > $null
		} else {
			throw "$parentNodePath does not exist in the xml."
		}
	}
}

<#
	.SYNOPSIS
		Encrypts a string
	
	.DESCRIPTION
		Encrypts a string into a given encoding type.
		Returns encrypted string
	
	.PARAMETER String
		String to encrypt
	
	.PARAMETER HashType
		Type of encryption to use
	
	.EXAMPLE
		PS C:\> Encrypt-String -String "EncryptThisString" -HashType "BASE64"
		PS C:\> Encrypt-String -String "EncryptThisString" -HashType "SHA256"
	
	.NOTES
		Returns encrypted string
#>
function Encrypt-String {
	param
	(
		[Parameter(Mandatory = $true)]
		[String]$String,
		[Parameter(Mandatory = $true)]
		$HashType
	)
	
	If ($HashType -eq "BASE64") {
		$Bytes = [System.Text.Encoding]::UTF8.GetBytes("$String")
		$StringBuilder = [System.Convert]::ToBase64String($Bytes)
		Return $StringBuilder
	}
	
	$StringBuilder = New-Object System.Text.StringBuilder
	[System.Security.Cryptography.HashAlgorithm]::Create($HashType).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($String)) | %{
		[Void]$StringBuilder.Append($_.ToString("x2"))
	}
	Return $StringBuilder.ToString()
}

<#
	.SYNOPSIS
		Returns information about disk space
	
	.DESCRIPTION
		Returns information about disk space
	
	.EXAMPLE
				PS C:\> Get-DiskSpace
	
	.NOTES
		Returns an array about disk usage information
#>
function Get-DiskSpace {
	$DiskSpace += Get-WmiObject Win32_Volume -Filter "DriveType='3'" -ComputerName $ComputerName | ForEach {
		New-Object PSObject -Property @{
			Name = $_.Name
			Label = $_.Label
			Computer = $ComputerName
			FreeSpace_GB = ([Math]::Round($_.FreeSpace /1GB, 2))
			TotalSize_GB = ([Math]::Round($_.Capacity /1GB, 2))
			UsedSpace_GB = ([Math]::Round($_.Capacity /1GB, 2)) - ([Math]::Round($_.FreeSpace /1GB, 2))
		}
	}
	
	Return $DiskSpace
}
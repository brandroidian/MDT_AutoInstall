<#
' // ***************************************************************************
' // 
' // FileName:  MDT_AutoInstall.ps1
' //            
' // Version:   1.00
' //            
' // Usage:     powershell.exe -executionpolicy bypass -file "%~dp0MDT_AutoInstall.ps1" -SvcAccountPassword P@ssw0rd -DSDrive C:\ -DeploymentType Capture -UpdateDeploymentShare
' //			powershell.exe -executionpolicy bypass -file "%~dp0MDT_AutoInstall.ps1" -SvcAccountPassword P@ssw0rd -DSDrive C:\ -DeploymentType Capture -IncludeApplications -UpdateDeploymentShare
' //			powershell.exe -executionpolicy bypass -file "%~dp0MDT_AutoInstall.ps1" -SvcAccountPassword P@ssw0rd -DSDrive C:\ -DeploymentType Capture -IncludeApplications
' //            powershell.exe -executionpolicy bypass -file "%~dp0MDT_AutoInstall.ps1" -SvcAccountPassword P@ssw0rd -DSDrive C:\ -DeploymentType Deploy -UpdateDeploymentShare
' //            powershell.exe -executionpolicy bypass -file "%~dp0MDT_AutoInstall.ps1" -SvcAccountPassword P@ssw0rd -DSDrive C:\ -DeploymentType NewCaptureOS
' //            powershell.exe -executionpolicy bypass -file "%~dp0MDT_AutoInstall.ps1" -SvcAccountPassword P@ssw0rd -DSDrive C:\ -InstallOnly
' //
' //            
' // Created:   1.0 (2019.07.12)
' //            Brandon Hilgeman
' //            brandon.hilgeman@gmail.com
' // ***************************************************************************
#>


<#-------------------------------------------------------------------------------
'---    Initialize Objects
'-------------------------------------------------------------------------------#>

param (
	[Parameter(Mandatory = $true)]
	[string]$SvcAccountPassword,
	
	[Parameter(Mandatory = $true)]
	[ValidateScript({Test-Path $_})]
	[string]$DSDrive,
	
	[Parameter(Mandatory = $false)]
	[string]$DeploymentType,
	
	[Parameter(Mandatory = $false)]
	[switch]$InstallOnly,
	
	[Parameter(Mandatory = $false)]
	[switch]$IncludeApplications,
	
	[Parameter(Mandatory = $false)]
	[switch]$InstallWDS,
	
	[Parameter(Mandatory = $false)]
	[switch]$UpdateDeploymentShare
)

<#-------------------------------------------------------------------------------
'---    Configure
'-------------------------------------------------------------------------------#>

$Global:ScriptName = $MyInvocation.MyCommand.Name

$Global:ComputerFQDN = (Get-WmiObject win32_computersystem).DNSHostName + "." + (Get-WmiObject win32_computersystem).Domain

$DSDrive = $DSDrive.TrimEnd("\")


<#-------------------------------------------------------------------------------
'---    Functions
'-------------------------------------------------------------------------------#>


Function Exit-Script {
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Set Execution Policy So MDT Console Will Open"
	Set-ExecutionPolicy bypass -Scope CurrentUser -Confirm:$false -WarningAction SilentlyContinue
	End-Log
	Pause
	Exit
}

function Import-MDTModule {
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Importing MDT Module"
	$ModulePath = "$MDTInstallPath" +
	"\bin\MicrosoftDeploymentToolkit.psd1"
	Import-Module $ModulePath
}

function Connect-DeploymentShare {
		
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Creating PSDrive: DS001"
	$params = @{
		Name	   = "DS001"
		PSProvider = "MDTProvider"
		Scope	   = "Global"
		Root	   = "$DSDrive\$DeploymentShareLocalPath"
		Description = "$DeploymentShareDescription"
		NetworkPath = "\\$ComputerFQDN\$DeploymentShareNetworkPath"
	}
	New-PSDrive @params -ErrorAction SilentlyContinue | Add-MDTPersistentDrive
	
}

function Import-Configuration {
	#Import configuration.ps1
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Importing configuration.ps1"
	$Configuration = Test-Path "$PSScriptRoot\configuration.ps1"
	if (!$Configuration) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Configuration.ps1 not found in script directory"
		Exit-Script
	}
	
	try {
		. "$PSScriptRoot\configuration.ps1"
	}
	catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Check configuration.ps1 for syntax errors"
		Exit-Script
	}
}

function Set-DeploymentSharePaths {
	If (($DeployMentType -eq "Capture") -or ($DeploymentType -eq "NewCaptureOS")) {
		$Global:DeploymentShareLocalPath = $CaptureDeploymentShareLocalPath
		$Global:DeploymentShareNetworkPath = $CaptureDeploymentShareNetworkPath
		$Global:DeploymentShareDescription = $CaptureDeploymentShareDescription
		$Global:CustomSettingsIni = $CaptureCustomSettingsIni
	}
	Else {
		$Global:DeploymentShareLocalPath = $DeployDeploymentShareLocalPath
		$Global:DeploymentShareNetworkPath = $DeployDeploymentShareNetworkPath
		$Global:DeploymentShareDescription = $DeployDeploymentShareDescription
		$Global:CustomSettingsIni = $DeployCustomSettingsIni
	}
	
	If ($DeploymentType -eq "NewCaptureOS") {
		Create-CaptureTaskSequences
	}
}

Function Import-ApplicationsFile {
	#Import applications.json and download Office Deployment Tool
	if ($IncludeApplications) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Importing applications.json"
		$Applications = Test-Path -Path "$PSScriptRoot\applications.json"
		if (!$Applications) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: No applications.json file found in script directory" -Type 3
			Exit-Script
		}
		else {
			try {
				$Global:Applist = Get-Content "$PSScriptRoot\applications.json" | ConvertFrom-Json
			}
			catch {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to load applications.json. Please check syntax and try again" -Type 3
				Exit-Script
			}
		}
	}
}

function Download-Installers {
	
	#If needed, create folder for installation file downloads
	If (!(Test-Path -Path "$PSScriptRoot\Installs")) {
		Create-Folder -Path "$PSScriptRoot\Installs"
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Installs\MicrosoftDeploymentToolkit_x64.msi")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading MDT $MDTVersion"
		$params = @{
			Source	    = $MDTUrl
			Destination = "$PSScriptRoot\Installs\MicrosoftDeploymentToolkit_x64.msi"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download MDT. ($($_))"
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Installs\adksetup.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading ADK $ADKVersion"
		$params = @{
			Source	    = $ADKUrl
			Destination = "$PSScriptRoot\Installs\adksetup.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download ADK. ($($_))" -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Installs\adkwinpesetup.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading ADK $ADKVersion WinPE Addon"
		$params = @{
			Source	    = $ADKWinPEUrl
			Destination = "$PSScriptRoot\Installs\adkwinpesetup.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download ADK WinPE Addon. ($($_))" -Type 3
			Exit-Script
		}
	}
}

function Install-Applications {
	If (!($MDTInstalled -eq $true)) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Installing MDT $MDTVersion"
		$params = @{
			Wait								  = $True
			PassThru							  = $True
			NoNewWindow						      = $True
			FilePath							  = "msiexec"
			ArgumentList						  = "/i ""$PSScriptRoot\Installs\MicrosoftDeploymentToolkit_x64.msi"" /qn INSTALLDIR=""$($MDTInstallPath)"" " +
			"/l*v ""$DefaultLogPath\MDT_Install.log"""
		}
		$Return = Start-Process @params
		if (@(0, 3010, 1641) -notcontains $Return.ExitCode) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to install MDT. Exit code: $($Return.ExitCode)" -Type 3
			Exit-Script
		}
		$Return = $null
	}
	
	If (!($ADKInstalled -eq $true)) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Installing ADK $ADKVersion"
		$params = @{
			Wait	    = $True
			PassThru    = $True
			NoNewWindow = $True
			FilePath    = "$PSScriptRoot\Installs\adksetup.exe"
			ArgumentList = "/quiet /installpath ""$ADKInstallPath"" /features OptionId.DeploymentTools OptionId.UserStateMigrationTool " +
			"/log ""$DefaultLogPath\ADK.log"""
		}
		$Return = Start-Process @params
		if ($Return.ExitCode -ne 0) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to install ADK. Exit code: $($Return.ExitCode)" -Type 3
			Exit-Script
		}
		$Return = $null
	}
	
	If (!($ADKWinPEInstalled -eq $true)) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Installing ADK $ADKVersion WinPE Addon"
		$params = @{
			Wait	    = $True
			PassThru    = $True
			NoNewWindow = $True
			FilePath    = "$PSScriptRoot\Installs\adkwinpesetup.exe"
			ArgumentList = "/quiet /features OptionId.WindowsPreinstallationEnvironment " +
			"/log ""$DefaultLogPath\ADK_WinPE.log"""
		}
		$Return = Start-Process @params
		if ($Return.ExitCode -ne 0) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to install ADK WinPE Addon. Exit code: $($Return.ExitCode)" -Type 3
			Exit-Script
		}
		$Return = $null
	}
}

function Set-LocalServiceAccount {
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Creating local Service Account for DeploymentShare"
	$params = @{
		Name = "$MDTServiceAccountName"
		Password = (ConvertTo-SecureString $SvcAccountPassword -AsPlainText -Force)
		AccountNeverExpires = $true
		PasswordNeverExpires = $true
	}
	Try {
		New-LocalUser @params -ErrorAction SilentlyContinue
	}
	Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to create local user, may already exist" -Type 3
	}
}

function Create-DeploymentShareShare{
	If (!(Test-Path -Path "$DSDrive\$DeploymentShareLocalPath")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Creating Deployment Share Directory"
		New-Item -Path "$DSDrive\$DeploymentShareLocalPath" -ItemType Directory -Force
	}
	
	$params = @{
		Name = "$DeploymentShareNetworkPath"
		Path = "$DSDrive\$DeploymentShareLocalPath"
		FullAccess = "Everyone"
	}
	Try {
		New-SmbShare @params -ErrorAction SilentlyContinue
	}
	Catch {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Unable to create share: $DeploymentShareNetworkPath, Share may already exist"
	}
}

function Set-Bootstrap {
	#Edit Bootstrap.ini
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Setting Bootstrap.ini"
	$BootstrapIni = @"
[Settings]
Priority=Default
[Default]
DeployRoot=\\$ComputerFQDN\$DeploymentShareNetworkPath
SkipBDDWelcome=YES
Userdomain=$env:COMPUTERNAME
UserID=$MDTServiceAccountName
UserPassword=$SvcAccountPassword
"@
	
	$params = @{
		Path = "$DSDrive\$DeploymentShareLocalPath\Control\Bootstrap.ini"
		Value = $BootstrapIni
		Force = $True
	}
	Set-Content @params -Confirm:$False | Out-Null
}

function Set-CustomSettings {
	#Edit CustomSettings.ini
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Setting CustomSettings.ini"
	$params = @{
		Path = "$DSDrive\$DeploymentShareLocalPath\Control\CustomSettings.ini"
		Value = $CustomSettingsIni
		Force = $True
	}
	Set-Content @params -Confirm:$False | Out-Null
}

function Disable-x86Support {
	if ($DisableX86Support) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Disabling x86 Support"
		$DeploymentShareSettings = "$DSDrive\$DeploymentShareLocalPath\Control\Settings.xml"
		$xmldoc = [XML](Get-Content $DeploymentShareSettings)
		$xmldoc.Settings.SupportX86 = "False"
		$xmldoc.Save($DeploymentShareSettings)
	}
}

function Import-Applications {
	if ($IncludeApplications) {
		foreach ($Application in $AppList) {
			If (!(Test-Path -Path "DS001:\Applications\$($Application.name)")) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading and importing $($Application.Name)"
				New-Item -Path "$PSScriptRoot\Installs\$($application.name)" -ItemType Directory -Force | Out-Null
				$params = @{
					Source = $Application.download
					Destination = "$PSScriptRoot\Installs\$($application.name)\$($Application.filename)"
				}
				try {
					Download-File @params -ErrorAction Stop
					$params = @{
						Path = "DS001:\Applications"
						Name = $Application.Name
						ShortName = $Application.Name
						Publisher = ""
						Language = ""
						Enable = "True"
						Version = $Application.version
						CommandLine = $Application.install
						WorkingDirectory = ".\Applications\$($Application.name)"
						ApplicationSourcePath = "$PSScriptRoot\Installs\$($application.name)"
						DestinationFolder = $Application.name + " " + $Application.version
					}
					Import-MDTApplication @params | Out-Null
				}
				catch {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download $($Application.name). Check URL is valid in applications.json. ($($_))" -Type 3
				}
			}
			Else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: $($Application.name) already exists. Will not import."
			}
			
		}
	}
}

function Install-WDS {
	#Install WDS
	if ($InstallWDS) {
		$OSInfo = Get-CimInstance -ClassName Win32_OperatingSystem
		if ($OSInfo.ProductType -eq 1) {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Workstation OS - WDS Not available" -Type 2
		}
		else {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Server OS - Checking if WDS available on this version"
			$WDSCheck = Get-WindowsFeature -Name WDS
			if ($WDSCheck) {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: WDS Role Available - Installing"
				Add-WindowsFeature -Name WDS -IncludeAllSubFeature -IncludeManagementTools | Out-Null
				$WDSInit = wdsutil /initialize-server /remInst:"$DSDrive\remInstall" /standalone
				$WDSConfig = wdsutil /Set-Server /AnswerClients:All
				$params = @{
					Path = "$DSDrive\$DeploymentShareLocalPath\Boot\LiteTouchPE_x64.wim"
					SkipVerify = $True
					NewImageName = "MDT Litetouch"
					
				}
				Import-WdsBootImage @params | Out-Null
			}
			else {
				Write-Log -Message "  $($MyInvocation.MyCommand.Name):: WDS Role not available on this version of Server"
			}
		}
	}
}

function Install-Capture {
	#Give service account $MDTServiceAccountName "Modify" rights to Captures directory
	Start-Process "icacls.exe" -ArgumentList "$DSDrive\$DeploymentShareLocalPath\Captures /grant:r $env:COMPUTERNAME\$($MDTServiceAccountName):(OI)(CI)M /T"
	
	Create-CaptureTaskSequences
	
	#Set WinPE Values in Settings.xml
	$XMLFilePath = "$DSDrive\$DeploymentShareLocalPath\Control\Settings.xml"
	[XML]$XMLFile = Get-Content -Path $XMLFilePath
	
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.ScratchSpace" -TextValue "512" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeAllDrivers" -TextValue "True" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeNetworkDrivers" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeMassStorageDrivers" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeVideoDrivers" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeSystemDrivers" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.ExtraDirectory" -TextValue "%DEPLOYROOT%/ExtraFiles/x64" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.GenerateGenericWIM" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.GenerateGenericISO" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.GenerateLiteTouchISO" -TextValue "True" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.LiteTouchISOName" -TextValue "Capture_x64.iso" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.SelectionProfile" -TextValue "Nothing" -NodeSeparatorCharacter "/"
	$XMLFile.Save($XMLFilePath)
	
	#download & Import Office 365 2016
	if ($IncludeO365 -eq $true) {
		If (!(Test-Path -Path "$DSDrive\$DeploymentShareLocalPath\Applications\Microsoft Office*")) {
			If (!(Test-Path -Path "$PSScriptRoot\Applications\O365\officedeploymenttool.exe")) {
				New-Item -ItemType Directory -Path "$PSScriptRoot\Applications\O365" -Force | Out-Null
				try {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading Office Deployment Toolkit"
					Download-File -Source $OfficeDeploymentToolUrl -Destination "$PSScriptRoot\Applications\O365\officedeploymenttool.exe"
				}
				catch {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download Office Deployment Toolkit. ($($_))"
					Exit-Script
				}
			}
			
			
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Extracting Office Deployment Toolkit"
			$params = @{
				FilePath = "$PSScriptRoot\Applications\O365\officedeploymenttool.exe"
				ArgumentList = "/quiet /extract:$PSScriptRoot\Applications\O365"
			}
			Start-Process @params -Wait
			
			Set-Content -Path "$PSScriptRoot\Applications\O365\configuration.xml" -Value $Office365Configurationxml -Force -Confirm:$false | Out-Null
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Importing Office 365 into MDT"
			$params = @{
				Path				  = "DS001:\Applications"
				Name				  = "Microsoft Office 365"
				ShortName			  = "Office 365 2016"
				Publisher			  = "Microsoft"
				Language			  = ""
				Enable			      = "True"
				Version			      = "Semi-Annual"
				Verbose			      = $true
				CommandLine		      = "setup.exe /configure configuration.xml"
				WorkingDirectory	  = ".\Applications\Microsoft Office 365"
				ApplicationSourcePath = "$PSScriptRoot\Applications\O365"
				DestinationFolder	  = "Microsoft Office 365"
			}
			Import-MDTApplication @params | Out-Null
		}
		
	}
	
	#Install C++ Redistributables
	if ($IncludeCPlusPlus -eq $true) {
		If (!(Test-Path -Path "$DSDrive\$DeploymentShareLocalPath\Applications\Microsoft C++ *")) {
			Download-Redistributables
			
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Importing C++ Redistributables"
			$params = @{
				Path				  = "DS001:\Applications"
				Name				  = "Microsoft C++ Redistributables"
				ShortName			  = "C++ Redistributables"
				Publisher			  = "Microsoft"
				Language			  = ""
				Enable			      = "True"
				Version			      = "2008 - 2019"
				Verbose			      = $true
				CommandLine		      = "cscript.exe Install-MicrosoftVisualC++x86x64.wsf"
				WorkingDirectory	  = ".\Applications\Microsoft C++ Redistributables"
				ApplicationSourcePath = "$PSScriptRoot\Applications\Microsoft C++ Redistributables"
				DestinationFolder	  = "Microsoft C++ Redistributables"
			}
			Import-MDTApplication @params | Out-Null
		}
	}
	
	If ($OperatingSystems -ne $false) {
		Set-CommonTaskSequence
	}
	
}

function Set-CommonTaskSequence {
	#Get Application GUID for each application
	$CommonApplications = Get-ChildItem -Path "DS001:\Applications"
	foreach ($CommonApplication in $CommonApplications) {
		If ($CommonApplication.Name -eq "Microsoft C++ Redistributables") {
			$RedistributableGUID = $CommonApplication.guid
		}
		If ($CommonApplication.Name -eq "Microsoft Office 365") {
			$O365GUID = $CommonApplication.guid
		}
		
		If ($IncludeCPlusPlus -and $IncludeO365) {
			#Hide COMMON Task Sequence From The Deployment Wizard
			$CommonTaskSequenceXML = @"
<?xml version="1.0"?>
<sequence version="3.00" name="Custom Task Sequence" description="Sample Custom Task Sequence">
  <group expand="true" name="Preinstall" description="" disable="false" continueOnError="false">
    <action />
    <condition>
      <operator type="and">
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">Phase</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">Preinstall</variable>
        </expression>
      </operator>
    </condition>
    <step type="SMS_TaskSequence_RunCommandLineAction" name="Disable Store Updates (Offline)" description="" disable="false" continueOnError="true" startIn="" successCodeList="0 3010" runIn="WinPEandFullOS">
      <defaultVarList>
        <variable name="PackageID" property="PackageID" />
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\Custom\DisableStoreUpdates\DisableStoreUpdates.wsf"</action>
      <condition>
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">_SMSTSInWinPE</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">True</variable>
        </expression>
      </condition>
    </step>
  </group>
  <group expand="true" name="State Restore" description="" disable="false" continueOnError="false">
    <action />
    <condition>
      <operator type="and">
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">Phase</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">StateRestore</variable>
        </expression>
      </operator>
    </condition>
    <step type="SMS_TaskSequence_RunCommandLineAction" name="Copy CMTrace" description="" disable="false" continueOnError="false" startIn="" successCodeList="0 3010" runIn="WinPEandFullOS">
      <defaultVarList>
        <variable name="PackageID" property="PackageID" />
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <action>xcopy /q /y "%DEPLOYROOT%\ExtraFiles\x64\Windows\System32" "C:\Windows\System32"</action>
      <condition></condition>
    </step>
    <step type="BDD_InstallRoles" name="Install .Net 3.5" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="OSRoleIndex" property="OSRoleIndex">13</variable>
        <variable name="OSRoles" property="OSRoles"></variable>
        <variable name="OSRoleServices" property="OSRoleServices"></variable>
        <variable name="OSFeatures" property="OSFeatures">NetFx3</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\ZTIOSRole.wsf"</action>
      <condition></condition>
    </step>
    <step type="SMS_TaskSequence_RebootAction" name="Restart computer" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="SMSRebootMessage" property="Message" />
        <variable name="SMSRebootTimeout" property="MessageTimeout">60</variable>
        <variable name="SMSRebootTarget" property="Target" />
      </defaultVarList>
      <action>smsboot.exe /target:WinPE</action>
      <condition></condition>
    </step>
    <step name="Windows Update (Pre-Application Installation)" disable="false" continueOnError="true" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIWindowsUpdate.wsf"</action>
      <defaultVarList>
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <condition></condition>
    </step>
    <step type="BDD_InstallApplication" name="Install Microsoft C++ Redistributables" disable="false" continueOnError="false" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIApplications.wsf"</action>
      <defaultVarList>
        <variable name="ApplicationGUID" property="ApplicationGUID">$RedistributableGUID</variable>
        <variable name="ApplicationSuccessCodes" property="ApplicationSuccessCodes">0 3010</variable>
      </defaultVarList>
      <condition></condition>
    </step>
    <step type="BDD_InstallApplication" name="Install Office 365 2016" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="ApplicationGUID" property="ApplicationGUID">$O365GUID</variable>
        <variable name="ApplicationSuccessCodes" property="ApplicationSuccessCodes">0 3010</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\ZTIApplications.wsf"</action>
      <condition></condition>
    </step>
    <step type="SMS_TaskSequence_RebootAction" name="Restart computer" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="SMSRebootMessage" property="Message"></variable>
        <variable name="SMSRebootTimeout" property="MessageTimeout">60</variable>
        <variable name="SMSRebootTarget" property="Target"></variable>
      </defaultVarList>
      <action>smsboot.exe /target:WinPE</action>
    </step>
    <step name="Windows Update (Post-Application Installation)" disable="false" continueOnError="true" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIWindowsUpdate.wsf"</action>
      <defaultVarList>
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
    </step>
  </group>
</sequence>
"@
		}
		
		If ($IncludeCPlusPlus -and (!($IncludeO365))) {
			#Hide COMMON Task Sequence From The Deployment Wizard
			$CommonTaskSequenceXML = @"
<?xml version="1.0"?>
<sequence version="3.00" name="Custom Task Sequence" description="Sample Custom Task Sequence">
  <group expand="true" name="Preinstall" description="" disable="false" continueOnError="false">
    <action />
    <condition>
      <operator type="and">
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">Phase</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">Preinstall</variable>
        </expression>
      </operator>
    </condition>
    <step type="SMS_TaskSequence_RunCommandLineAction" name="Disable Store Updates (Offline)" description="" disable="false" continueOnError="true" startIn="" successCodeList="0 3010" runIn="WinPEandFullOS">
      <defaultVarList>
        <variable name="PackageID" property="PackageID" />
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\Custom\DisableStoreUpdates\DisableStoreUpdates.wsf"</action>
      <condition>
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">_SMSTSInWinPE</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">True</variable>
        </expression>
      </condition>
    </step>
  </group>
  <group expand="true" name="State Restore" description="" disable="false" continueOnError="false">
    <action />
    <condition>
      <operator type="and">
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">Phase</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">StateRestore</variable>
        </expression>
      </operator>
    </condition>
    <step type="SMS_TaskSequence_RunCommandLineAction" name="Copy CMTrace" description="" disable="false" continueOnError="false" startIn="" successCodeList="0 3010" runIn="WinPEandFullOS">
      <defaultVarList>
        <variable name="PackageID" property="PackageID" />
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <action>xcopy /q /y "%DEPLOYROOT%\ExtraFiles\x64\Windows\System32" "C:\Windows\System32"</action>
      <condition></condition>
    </step>
    <step type="BDD_InstallRoles" name="Install .Net 3.5" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="OSRoleIndex" property="OSRoleIndex">13</variable>
        <variable name="OSRoles" property="OSRoles"></variable>
        <variable name="OSRoleServices" property="OSRoleServices"></variable>
        <variable name="OSFeatures" property="OSFeatures">NetFx3</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\ZTIOSRole.wsf"</action>
      <condition></condition>
    </step>
    <step type="SMS_TaskSequence_RebootAction" name="Restart computer" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="SMSRebootMessage" property="Message" />
        <variable name="SMSRebootTimeout" property="MessageTimeout">60</variable>
        <variable name="SMSRebootTarget" property="Target" />
      </defaultVarList>
      <action>smsboot.exe /target:WinPE</action>
      <condition></condition>
    </step>
    <step name="Windows Update (Pre-Application Installation)" disable="false" continueOnError="true" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIWindowsUpdate.wsf"</action>
      <defaultVarList>
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <condition></condition>
    </step>
    <step type="BDD_InstallApplication" name="Install Microsoft C++ Redistributables" disable="false" continueOnError="false" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIApplications.wsf"</action>
      <defaultVarList>
        <variable name="ApplicationGUID" property="ApplicationGUID">$RedistributableGUID</variable>
        <variable name="ApplicationSuccessCodes" property="ApplicationSuccessCodes">0 3010</variable>
      </defaultVarList>
      <condition></condition>
    </step>
    <step type="SMS_TaskSequence_RebootAction" name="Restart computer" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="SMSRebootMessage" property="Message"></variable>
        <variable name="SMSRebootTimeout" property="MessageTimeout">60</variable>
        <variable name="SMSRebootTarget" property="Target"></variable>
      </defaultVarList>
      <action>smsboot.exe /target:WinPE</action>
    </step>
    <step name="Windows Update (Post-Application Installation)" disable="false" continueOnError="true" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIWindowsUpdate.wsf"</action>
      <defaultVarList>
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
    </step>
  </group>
</sequence>
"@
		}
		
		If ((!($IncludeCPlusPlus)) -and $IncludeO365) {
			$CommonTaskSequenceXML = @"
<?xml version="1.0"?>
<sequence version="3.00" name="Custom Task Sequence" description="Sample Custom Task Sequence">
  <group expand="true" name="Preinstall" description="" disable="false" continueOnError="false">
    <action />
    <condition>
      <operator type="and">
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">Phase</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">Preinstall</variable>
        </expression>
      </operator>
    </condition>
    <step type="SMS_TaskSequence_RunCommandLineAction" name="Disable Store Updates (Offline)" description="" disable="false" continueOnError="true" startIn="" successCodeList="0 3010" runIn="WinPEandFullOS">
      <defaultVarList>
        <variable name="PackageID" property="PackageID" />
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\Custom\DisableStoreUpdates\DisableStoreUpdates.wsf"</action>
      <condition>
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">_SMSTSInWinPE</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">True</variable>
        </expression>
      </condition>
    </step>
  </group>
  <group expand="true" name="State Restore" description="" disable="false" continueOnError="false">
    <action />
    <condition>
      <operator type="and">
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">Phase</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">StateRestore</variable>
        </expression>
      </operator>
    </condition>
    <step type="SMS_TaskSequence_RunCommandLineAction" name="Copy CMTrace" description="" disable="false" continueOnError="false" startIn="" successCodeList="0 3010" runIn="WinPEandFullOS">
      <defaultVarList>
        <variable name="PackageID" property="PackageID" />
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <action>xcopy /q /y "%DEPLOYROOT%\ExtraFiles\x64\Windows\System32" "C:\Windows\System32"</action>
      <condition></condition>
    </step>
    <step type="BDD_InstallRoles" name="Install .Net 3.5" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="OSRoleIndex" property="OSRoleIndex">13</variable>
        <variable name="OSRoles" property="OSRoles"></variable>
        <variable name="OSRoleServices" property="OSRoleServices"></variable>
        <variable name="OSFeatures" property="OSFeatures">NetFx3</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\ZTIOSRole.wsf"</action>
      <condition></condition>
    </step>
    <step type="SMS_TaskSequence_RebootAction" name="Restart computer" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="SMSRebootMessage" property="Message" />
        <variable name="SMSRebootTimeout" property="MessageTimeout">60</variable>
        <variable name="SMSRebootTarget" property="Target" />
      </defaultVarList>
      <action>smsboot.exe /target:WinPE</action>
      <condition></condition>
    </step>
    <step name="Windows Update (Pre-Application Installation)" disable="false" continueOnError="true" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIWindowsUpdate.wsf"</action>
      <defaultVarList>
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <condition></condition>
    </step>
    <step type="BDD_InstallApplication" name="Install Office 365 2016" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="ApplicationGUID" property="ApplicationGUID">$O365GUID</variable>
        <variable name="ApplicationSuccessCodes" property="ApplicationSuccessCodes">0 3010</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\ZTIApplications.wsf"</action>
      <condition></condition>
    </step>
    <step type="SMS_TaskSequence_RebootAction" name="Restart computer" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="SMSRebootMessage" property="Message"></variable>
        <variable name="SMSRebootTimeout" property="MessageTimeout">60</variable>
        <variable name="SMSRebootTarget" property="Target"></variable>
      </defaultVarList>
      <action>smsboot.exe /target:WinPE</action>
    </step>
    <step name="Windows Update (Post-Application Installation)" disable="false" continueOnError="true" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIWindowsUpdate.wsf"</action>
      <defaultVarList>
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
    </step>
  </group>
</sequence>
"@
		}
		
		If ((!($IncludeCPlusPlus)) -and (!($IncludeO365))) {
			$CommonTaskSequenceXML = @"
<?xml version="1.0"?>
<sequence version="3.00" name="Custom Task Sequence" description="Sample Custom Task Sequence">
  <group expand="true" name="Preinstall" description="" disable="false" continueOnError="false">
    <action />
    <condition>
      <operator type="and">
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">Phase</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">Preinstall</variable>
        </expression>
      </operator>
    </condition>
    <step type="SMS_TaskSequence_RunCommandLineAction" name="Disable Store Updates (Offline)" description="" disable="false" continueOnError="true" startIn="" successCodeList="0 3010" runIn="WinPEandFullOS">
      <defaultVarList>
        <variable name="PackageID" property="PackageID" />
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\Custom\DisableStoreUpdates\DisableStoreUpdates.wsf"</action>
      <condition>
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">_SMSTSInWinPE</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">True</variable>
        </expression>
      </condition>
    </step>
  </group>
  <group expand="true" name="State Restore" description="" disable="false" continueOnError="false">
    <action />
    <condition>
      <operator type="and">
        <expression type="SMS_TaskSequence_VariableConditionExpression">
          <variable name="Variable">Phase</variable>
          <variable name="Operator">equals</variable>
          <variable name="Value">StateRestore</variable>
        </expression>
      </operator>
    </condition>
    <step type="SMS_TaskSequence_RunCommandLineAction" name="Copy CMTrace" description="" disable="false" continueOnError="false" startIn="" successCodeList="0 3010" runIn="WinPEandFullOS">
      <defaultVarList>
        <variable name="PackageID" property="PackageID" />
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <action>xcopy /q /y "%DEPLOYROOT%\ExtraFiles\x64\Windows\System32" "C:\Windows\System32"</action>
      <condition></condition>
    </step>
    <step type="BDD_InstallRoles" name="Install .Net 3.5" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="OSRoleIndex" property="OSRoleIndex">13</variable>
        <variable name="OSRoles" property="OSRoles"></variable>
        <variable name="OSRoleServices" property="OSRoleServices"></variable>
        <variable name="OSFeatures" property="OSFeatures">NetFx3</variable>
      </defaultVarList>
      <action>cscript.exe "%SCRIPTROOT%\ZTIOSRole.wsf"</action>
      <condition></condition>
    </step>
    <step type="SMS_TaskSequence_RebootAction" name="Restart computer" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="SMSRebootMessage" property="Message" />
        <variable name="SMSRebootTimeout" property="MessageTimeout">60</variable>
        <variable name="SMSRebootTarget" property="Target" />
      </defaultVarList>
      <action>smsboot.exe /target:WinPE</action>
      <condition></condition>
    </step>
    <step name="Windows Update (Pre-Application Installation)" disable="false" continueOnError="true" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIWindowsUpdate.wsf"</action>
      <defaultVarList>
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
      <condition></condition>
    </step>
    <step type="SMS_TaskSequence_RebootAction" name="Restart computer" description="" disable="false" continueOnError="false" runIn="WinPEandFullOS" successCodeList="0 3010">
      <defaultVarList>
        <variable name="SMSRebootMessage" property="Message"></variable>
        <variable name="SMSRebootTimeout" property="MessageTimeout">60</variable>
        <variable name="SMSRebootTarget" property="Target"></variable>
      </defaultVarList>
      <action>smsboot.exe /target:WinPE</action>
    </step>
    <step name="Windows Update (Post-Application Installation)" disable="false" continueOnError="true" successCodeList="0 3010" description="" startIn="">
      <action>cscript.exe "%SCRIPTROOT%\ZTIWindowsUpdate.wsf"</action>
      <defaultVarList>
        <variable name="RunAsUser" property="RunAsUser">false</variable>
        <variable name="SMSTSRunCommandLineUserName" property="SMSTSRunCommandLineUserName"></variable>
        <variable name="SMSTSRunCommandLineUserPassword" property="SMSTSRunCommandLineUserPassword"></variable>
        <variable name="LoadProfile" property="LoadProfile">false</variable>
      </defaultVarList>
    </step>
  </group>
</sequence>
"@
		}
		
		Copy-Folder -Source "$PSScriptRoot\Scripts\Capture\Custom" -Destination "$DSDrive\$DeploymentShareLocalPath\Scripts"
		
		Set-Content -Path "$DSDrive\$DeploymentShareLocalPath\Control\COMMON\TS.xml" -Value $CommonTaskSequenceXML -Force -Confirm:$false | Out-Null
		
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Make Common TS Hidden"
		$XMLFilePath = "$DSDrive\$DeploymentShareLocalPath\Control\TaskSequences.xml"
		[XML]$XMLFile = Get-Content -Path $XMLFilePath
		Set-XmlElementsAttributeValue -XmlDocument $XMLFile -ElementPath "tss.ts[@guid='$CommonTaskSequenceGUID']" -AttributeName "hide" -AttributeValue "True"
		$XMLFile.Save($XMLFilePath)
		
	}
	
}

function Create-CaptureTaskSequences {
	Import-MDTModule
	Connect-DeploymentShare
	
	#Find Operating System wims to import
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Checking for wim files to import"
	$Wims = Get-ChildItem $PSScriptRoot -Filter "install.wim" -Recurse | Select -ExpandProperty FullName
	if ($Wims) {
		foreach ($Wim in $Wims) {
			$WimName = (Split-Path $Wim -Leaf).TrimEnd(".wim")
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Operating System .wim found - will import"
			$WimBuild = Get-WindowsImage -ImagePath $Wim -Name "Windows 10 Enterprise"

			ForEach ($WimBuildVersion in $WimBuildList.Keys) {
				If ($WimBuild.version -like $WimBuildVersion + "*") {
					$OSBuild = $WimBuildList[$WimBuildVersion]
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Importing Windows 10 Enterprise $OSBuild."
				}
			}
			If ($WimBuild.ImageName -eq "Windows 10 Enterprise") {
				$params = @{
					Path			  = "DS001:\Operating Systems"
					SourceFile	      = $Wim
					DestinationFolder = "Windows 10 Enterprise x64 $OSBuild"
				}
				If (!(Test-Path -Path "$DSDrive\$DeploymentShareLocalPath\Operating Systems\Windows 10 Enterprise x64 $OSBuild")) {
					$OSData = Import-MDTOperatingSystem @params | Out-Null
				}
				Else {
					Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Windows 10 Enterprise $OSBuild already exists. Skipping operating system import."
				}
			}
		}
	}
	else {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: No WIM files found to import" -Type 2
		$Global:OperatingSystems = $false
	}
	
	
	If ($OperatingSystems -ne $false) {
		#Create Task Sequence for each Operating System
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Creating Task Sequence for each imported Operating System"
		$OperatingSystems = Get-ChildItem -Path "DS001:\Operating Systems"
	}
	
	
	
	if ($OperatingSystems -ne $false) {
		Copy-File -Source "$PSScriptRoot\TaskSequences\Capture\Capture.xml" -Destination "$MDTInstallPath\Templates\Capture.xml"
		foreach ($OS in $OperatingSystems) {
			If ($OS.Flags -eq "Enterprise") {
				$WimName = Split-Path -Path $OS.Source -Leaf

				ForEach ($WimBuildVersion in $WimBuildList.Keys) {
					If ($OS.Build -like $WimBuildVersion + "*") {
						$OSBuild = $WimBuildList[$WimBuildVersion]
						Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Importing Windows 10 Enterprise $OSBuild."
					}
				}
			
				$params = @{
					Path = "DS001:\Task Sequences"
					Name = "Capture Windows 10 $OSBuild"
					Template = "Capture.xml"
					Comments = ""
					ID   = "Capture10_$OSBuild"
					Version = "1.0"
					OperatingSystemPath = "DS001:\Operating Systems\$($OS.Name)"
					FullName = $OSFullName
					OrgName = $OSOrgName
					HomePage = $OSHomePage
				}
				If (!(Test-Path -Path "$DSDrive\$DeploymentShareLocalPath\Control\Capture10_$OSBuild")) {
					Import-MDTTaskSequence @params | Out-Null
				}
			}
			Else {
				remove-item -path "DS001:\Operating Systems\$($OS.Name)" -verbose
			}
		}
		
		
		$params = @{
			Path	 = "DS001:\Task Sequences"
			Name	 = "Capture Common Steps"
			Template = "Custom.xml"
			Comments = ""
			ID	     = "COMMON"
			Version  = "1.0"
		}
		
		If (!(Test-Path -Path "$DSDrive\$DeploymentShareLocalPath\Control\COMMON")) {
			Import-MDTTaskSequence @params | Out-Null
		}
		
		$TaskSequences = Get-ChildItem -Path "DS001:\Task Sequences"
		foreach ($TaskSequence in $TaskSequences) {
			If ($TaskSequence.Name -eq "Capture Common Steps") {
				$Global:CommonTaskSequenceGUID = $TaskSequence.guid
			}
		}
		
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Set Common TS GUID for Each Capture TS"
		$TSXMLs = Get-ChildItem "$DSDrive\$DeploymentShareLocalPath\Control" -Filter "ts.xml" -Recurse | Select -ExpandProperty FullName
		ForEach ($TSXML in $TSXMLs) {
			If ($TSXML -notlike "*Common*") {
				$XMLFilePath = $TSXML
				[XML]$XMLFile = Get-Content -Path $XMLFilePath
				
				Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "sequence.group.step.defaultVarList.variable[@property='SubTaskSequenceGUID']" -TextValue $CommonTaskSequenceGUID #-NamespaceURI "//ns:variable[@property='SubTaskSequenceGUID']"
				$XMLFile.Save($XMLFilePath)
			}
		}
	}
	
	
	If ($DeploymentType -eq "NewCaptureOS") {
		Set-CommonTaskSequence
		Exit-Script
	}
}

function Install-Deploy {
	$XMLFilePath = "$DSDrive\$DeploymentShareLocalPath\Control\Settings.xml"
	[XML]$XMLFile = Get-Content -Path $XMLFilePath
	
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.ScratchSpace" -TextValue "512" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeAllDrivers" -TextValue "True" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeNetworkDrivers" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeMassStorageDrivers" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeVideoDrivers" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.IncludeSystemDrivers" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.ExtraDirectory" -TextValue "%DEPLOYROOT%/ExtraFiles/x64" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.GenerateGenericWIM" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.GenerateGenericISO" -TextValue "False" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.GenerateLiteTouchISO" -TextValue "True" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.LiteTouchISOName" -TextValue "Deploy_x64.iso" -NodeSeparatorCharacter "/"
	Set-XmlElementsTextValue -XmlDocument $XMLFile -ElementPath "Settings/Boot.x64.SelectionProfile" -TextValue "Nothing" -NodeSeparatorCharacter "/"
	$XMLFile.Save($XMLFilePath)
	
	Copy-Folder -Source "$PSScriptRoot\Scripts\Deploy\Custom" -Destination "$DSDrive\$DeploymentShareLocalPath\Scripts"
	
}

function Download-Redistributables {
	#Download C++ Redistributables
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2008\vcredist_x86.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2008 x86"
		$params = @{
			Source	    = $CPP2008x86URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2008\vcredist_x86.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2008 x86." -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2008\vcredist_x64.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2008 x64"
		$params = @{
			Source	    = $CPP2008x64URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2008\vcredist_x64.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2008 x64" -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2010\vcredist_x86.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2010 x86"
		$params = @{
			Source	    = $CPP2010x86URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2010\vcredist_x86.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2010 x86" -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2010\vcredist_x64.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2010 x64"
		$params = @{
			Source	    = $CPP2010x64URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2010\vcredist_x64.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2010 x64" -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2012\vcredist_x86.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2012 x86"
		$params = @{
			Source	    = $CPP2012x86URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2012\vcredist_x86.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2012 x86" -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2012\vcredist_x64.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2012 x64"
		$params = @{
			Source	    = $CPP2012x64URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2012\vcredist_x64.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2012 x64" -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2013\vcredist_x86.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2013 x86"
		$params = @{
			Source	    = $CPP2013x86URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2013\vcredist_x86.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2013 x86" -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2013\vcredist_x64.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2013 x64"
		$params = @{
			Source	    = $CPP2013x64URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2013\vcredist_x64.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2013 x64" -Type 3
			Exit-Script
		}
	}
	<#
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2017\VC_redistx86.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2017 x86"
		$params = @{
			Source	    = $CPP2017x86URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2017\VC_redistx86.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2017 x86" -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2017\VC_redistx64.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2017 x64"
		$params = @{
			Source	    = $CPP2017x64URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2017\VC_redistx64.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2017 x64" -Type 3
			Exit-Script
		}
	}
	
	#>
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2019\VC_redistx86.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2019 x86"
		$params = @{
			Source	    = $CPP2019x86URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2019\VC_redistx86.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2019 x86" -Type 3
			Exit-Script
		}
	}
	
	If (!(Test-Path -Path "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2019\VC_redistx64.exe")) {
		Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Downloading C++ 2019 x64"
		$params = @{
			Source	    = $CPP2019x64URL
			Destination = "$PSScriptRoot\Applications\Microsoft C++ Redistributables\Source\VS2019\VC_redistx64.exe"
		}
		try {
			Download-File @params -ErrorAction Stop
		}
		catch {
			Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Failed to download C++ 2019 x64" -Type 3
			Exit-Script
		}
	}
}

<#-------------------------------------------------------------------------------
'---    Install
'-------------------------------------------------------------------------------#>

Function Start-Install {
	
	Import-Configuration
	Set-DeploymentSharePaths
	
	#Check if provided deployment share path already exists
<#	If (Test-Path -Path "$DSDrive\$DeploymentShareLocalPath") {
		Write-Log -Message "  WARNING: Deployment Share Path ""$DSDrive\$DeploymentShareLocalPath"" Already Exists. Enter A New Path in Configuration.ps1" -Type 2
		Exit-Script
	}
#>
	
	Import-ApplicationsFile
	
	$Global:MDTInstalled = Is-SoftwareInstalled -Product $MDTProductSearch -Version "*"
	$Global:ADKInstalled = Is-SoftwareInstalled -Product $ADKProductSearch -Version "*"
	$Global:ADKWinPEInstalled = Is-SoftwareInstalled -Product $ADKWinPEProductSearch -Version "*"
	
	If ((!($MDTInstalled -eq $true)) -or (!($ADKInstalled -eq $True)) -or (!($ADKWinPEInstalled -eq $true))){
		Download-Installers
		Install-Applications
	}
	
	If ($InstallOnly) {
		Write-Log -Message "  Install Only Complete"
		Exit-Script
	}
	
	Import-MDTModule	
	Set-LocalServiceAccount	
	Create-DeploymentShareShare		
	Connect-DeploymentShare
	
	If ($DeploymentType -eq "Capture") {
		Install-Capture
	}
	ElseIf($DeploymentType -eq "Deploy") {
		Install-Deploy
	}
	Else {
		Exit-Script
	}
	
	Set-BootStrap
	Set-CustomSettings
	Disable-x86Support
	
	Copy-Folder -Source "$PSScriptRoot\ExtraFiles" -Destination "$DSDrive\$DeploymentShareLocalPath"
	
	#Create LiteTouch Boot WIM & ISO
	Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Creating LiteTouch Boot Media"
	If ($UpdateDeploymentShare) {
		Update-MDTDeploymentShare -Path "DS001:" -Force -Verbose | Out-Null
	}
	
	
	Import-Applications
	
	Install-WDS
	
	Exit-Script
}



<#-------------------------------------------------------------------------------
'---    UnInstall
'-------------------------------------------------------------------------------#>

Function Start-Uninstall{

    #Run-Uninstall -Name $ProductSearch -Version $ProductVersion
}

Function Start-WPFApp{
	
	#Get-XAML -Path "$PSScriptRoot\Example\MainWindow.xaml"
    
}


<#-------------------------------------------------------------------------------
'---    Start
'-------------------------------------------------------------------------------#>

Import-Module -WarningAction SilentlyContinue "$PSScriptRoot\ScriptLibrary1.23.psm1"
Set-GlobalVariables
Start-Log
Write-Log "  Runtime: $Runtime" -Type 1
Set-Mode
#$Null = $Global:Form.ShowDialog() #Uncomment for WPF Apps
End-Log


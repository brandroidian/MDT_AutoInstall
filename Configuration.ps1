# Download URLs and versions for MDT & ADK
$Global:MDTUrl = "https://download.microsoft.com/download/3/3/9/339BE62D-B4B8-4956-B58D-73C4685FC492/MicrosoftDeploymentToolkit_x64.msi"
$Global:MDTVersion = "8456"
$Global:ADKUrl = "http://download.microsoft.com/download/0/1/C/01CC78AA-B53B-4884-B7EA-74F2878AA79F/adk/adksetup.exe"
$Global:ADKWinPEUrl = "http://download.microsoft.com/download/D/7/E/D7E22261-D0B3-4ED6-8151-5E002C7F823D/adkwinpeaddons/adkwinpesetup.exe"
$Global:ADKVersion = "1809"

# Download URLs and versions for C++ Redistributables
$Global:CPP2008x86URL = "http://download.microsoft.com/download/1/1/1/1116b75a-9ec3-481a-a3c8-1777b5381140/vcredist_x86.exe"
$Global:CPP2008x64URL = "http://download.microsoft.com/download/d/2/4/d242c3fb-da5a-4542-ad66-f9661d0a8d19/vcredist_x64.exe"
$Global:CPP2010x86URL = "http://download.microsoft.com/download/5/B/C/5BC5DBB3-652D-4DCE-B14A-475AB85EEF6E/vcredist_x86.exe"
$Global:CPP2010x64URL = "http://download.microsoft.com/download/3/2/2/3224B87F-CFA0-4E70-BDA3-3DE650EFEBA5/vcredist_x64.exe"
$Global:CPP2012x86URL = "http://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU3/vcredist_x86.exe"
$Global:CPP2012x64URL = "http://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU3/vcredist_x64.exe"
$Global:CPP2013x86URL = "http://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x86.exe"
$Global:CPP2013x64URL = "http://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x64.exe"
$Global:CPP2013x86URL = "https://go.microsoft.com/fwlink/?LinkId=746571"
$Global:CPP2013x64URL = "https://go.microsoft.com/fwlink/?LinkId=746572"
#Removed 2015 and 2017 as it's been replaced with 2019
#$Global:CPP2017x86URL = "https://go.microsoft.com/fwlink/?LinkId=746571"
#$Global:CPP2017x64URL = "https://go.microsoft.com/fwlink/?LinkId=746572"
$Global:CPP2019x86URL = "https://aka.ms/vs/16/release/VC_redist.x86.exe"
$Global:CPP2019x64URL = "https://aka.ms/vs/16/release/VC_redist.x64.exe"

#Download URL for Office Deployment Tool
$Global:OfficeDeploymentToolUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11617-33601.exe"

#Custom Paths For Installs.
#Default Values to use
#$Global:MDTInstallPath = "C:\Program Files\Microsoft Deployment Toolkit"
#$Global:ADKInstallPath = "C:\Program Files (x86)\Windows Kits\10"

$Global:MDTInstallPath = "C:\Program Files\Microsoft Deployment Toolkit"
$Global:ADKInstallPath = "C:\Program Files (x86)\Windows Kits\10"

#Include O365 and/or C++ in Capture
$Global:IncludeO365 = $false
$Global:IncludeCPlusPlus = $true

#Capture Config
$Global:CaptureDeploymentShareLocalPath = "OSD\MDT\Image_Capture"
$Global:CaptureDeploymentShareNetworkPath = "Image_Capture$"
$Global:CaptureDeploymentShareDescription = "Image_Capture"

#Deploy Config
$Global:DeployDeploymentShareLocalPath = "OSD\MDT\Image_Deploy"
$Global:DeployDeploymentShareNetworkPath = "Image_Deploy$"
$Global:DeployDeploymentShareDescription = "Image_Deploy"

#Operating System Import Config
$Global:OSFullName = "Viamonstra"
$Global:OSOrgName = "Viamonstra"
$Global:OSHomePage = "http:\\www.google.com"
$Global:OSLocalAdminPassword = "P@ssw0rd"

#Local Service Account to be created
$Global:MDTServiceAccountName = "SVC_MDT"


# Disable x86 Support - If you are only using 64 bit OS, set this to $True. If not, set to $False
# If x86 is disabled, MDT is faster at re-generating boot images
$Global:DisableX86Support = $True

#Search query to skip installs if already installed (Only modify if MSFT changes in future versions of MDT and ADK)
$Global:MDTProductSearch = "Microsoft Deployment Toolkit *"
$Global:ADKProductSearch = "Kits Configuration Installer"
$Global:ADKWinPEProductSearch = "Windows PE *"

# To get Wim Build version for this list
# Run $WimBuild = Get-WindowsImage -ImagePath <Path to Install.wim> -Name "Windows 10 Enterprise"
# $WimBuild.Version
# WimBuild List
$Global:WIMBuildList = @{
	"10.0.17134" = '1803'
	"10.0.17763" = '1809'
	"10.0.18362" = '1903'
}

# Change according to requirements, remove if not required (customsettings.ini)
# CustomSettings.ini options can be found here: https://docs.microsoft.com/en-us/sccm/mdt/toolkit-reference#properties-60
$Global:CaptureCustomSettingsIni = @"
[Settings]
Priority=TaskSequenceID,Default
Properties=MyCustomProperty

[Default]
OSInstall=YES
SkipCapture=YES
SkipAdminPassword=YES
SkipComputerBackup=YES
SkipBitLocker=YES
SkipApplications=YES
SkipComputerName=YES
SkipPackageDisplay=YES
SkipSummary=YES
SkipFinalSummary=YES
SkipUserData=YES
SkipWizard=NO
SkipComputerBackup=NO
SkipProductKey=YES
SkipTaskSequence=NO
SkipDomainMembership=YES
SkipLocaleSelection=YES
SkipTimeZone=YES
SkipRoles=YES
SkipRearm=YES

HideShell=YES

OSDComputerName=CAPTURE

_SMSTSOrgName=$OSOrgName Capture

DoCapture=YES
ComputerBackupLocation=NETWORK
BackupShare=%DEPLOYROOT%
BackupDir=Captures
BackupFile=%TaskSequenceID%_#year(date) & "-" & month(date) & "-" & day(date)#.wim

;SLShare=\\$ComputerFQDN\OSDLogs$\#Year(date) & "_" & Month(date) & "_" & Day(date)#

KeyboardLocale=0409:00000409
UserLocale=en-US
UILanguage=en-US

TimeZoneName=Eastern Standard Time

AdminPassword=$OSLocalAdminPassword

DeploymentType=NewComputer
DoNotCreateExtraPartition=YES
FinishAction=SHUTDOWN
;EventService=http://$($ComputerFQDN):9800

"@

$Global:DeployCustomSettingsIni = @"
[Settings]
Priority=ByVM,Make,UserExit,RunLTI,ReadLTI,Default
Properties=cDomain,cImagingOU,cTargetOU,cOS,cTSList,RunWizard,cDC,cOSDComputerName

[ByVM]
Subsection=VM-%IsVM%

[VM-True]
DoNotCreateExtraPartition=YES
cOSDComputerName=VM-#Left("%SerialNumber%",12)#

[VM-False]
Subsection=Laptop-%IsLaptop%

[Laptop-True]
cOSDComputerName=LT-#Left("%SerialNumber%",12)#

[Laptop-False]
cOSDComputerName=DT-#Left("%SerialNumber%",12)#

[HP]
[Lenovo]
[Microsoft Corporation]

[USEREXIT]
UserExit=Custom\LTIWizard\UserExit.vbs

[RUNLTI]
RunWizard=#RunLTI()#

[READLTI]
SkipTaskSequence=#cTSList()#
cOS=#cOperatingSystem()#
cAppProfile=#cAppProfile()#
OSDCOMPUTERNAME=#OSDCOMPUTERNAME()#

[Default]
SkipCapture=YES
SkipAdminPassword=YES
SkipProductKey=YES
SkipComputerBackup=YES
SkipBitLocker=YES
SkipApplications=NO
SkipComputerName=NO
SkipPackageDisplay=YES
SkipSummary=YES
SkipFinalSummary=YES
SkipWizard=NO
SkipUserData=YES
SkipDomainMembership=YES
SkipLocaleSelection=YES
SkipComputerBackup=YES
SkipTaskSequence=YES
SkipTimeZone=YES

DeploymentType=NewComputer
AdminPassword=$OSLocalAdminPassword
_SMSTSOrgName=$OSOrgName Deployment

OSInstall=Y

;WSUSServer=http://WSUS.contoso.com:8530
;SLShare=\\$ComputerFQDN\OSDLogs$\#Year(date) & "_" & Month(date) & "_" & Day(date)#
;SLShareDynamicLogging=\\$ComputerFQDN\OSDLogs$\#Year(date) & "_" & Month(date) & "_" & Day(date)#\%OSDCOMPUTERNAME%

;TaskSequenceID=DEPLOY
TimeZoneName=Eastern Standard Time
;FinishAction=Restart

DoCapture=NO

cDC=DC01.contoso.com
cImagingOU=OU=Imaging,OU=ViaMonstra,DC=corp,DC=viamonstra,DC=com
cDomain=contoso.com
cTargetOU=OU=Workstations,OU=ViaMonstra,DC=corp,DC=viamonstra,DC=com

;JoinWorkgroup=WORKGROUP
JoinDomain=contoso.com
MachineObjectOU=OU=Imaging,OU=ViaMonstra,DC=corp,DC=viamonstra,DC=com

DomainAdmin=CM_jd
DomainAdminPassword=P@ssw0rd
DomainAdminDomain=contoso.com

KeyboardLocale=en-us
UserLocale=en-us
UILanguage=en-us

;HideShell=Yes
;DisableTaskMgr=YES

;EventService=http://$($ComputerFQDN):9800


"@

#Office 365 Configuration XML
$Global:Office365ConfigurationXml = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="Broad">
        <Product ID="O365ProPlusRetail">
            <Language ID="en-us"/>
        </Product>
    </Add>
    <Updates Enabled="FALSE"/>
    <Display Level="None" AcceptEULA="TRUE"/>
    <Logging Level="Standard" Path="%TEMP%"/>
    <Property Name="FORCEAPPSHUTDOWN" Value="FALSE"/>
    <Property Name="SharedComputerLicensing" Value="0"/>
    <Property Name="PinIconsToTaskbar" Value="FALSE"/>
</Configuration>
"@




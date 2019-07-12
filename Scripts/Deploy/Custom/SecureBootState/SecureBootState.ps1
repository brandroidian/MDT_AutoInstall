<#
' // ***************************************************************************
' // 
' // FileName:  SecureBootState.ps1
' //            
' // Version:   1.00
' //            
' // Usage:     powershell.exe -executionpolicy bypass -file SecureBootState.ps1
' //          
' //
' //            
' // Created:   1.18 (2019.02.06)
' //            Brandon Hilgeman
' //            brandon.hilgeman@gmail.com
' // ***************************************************************************
#>


<#-------------------------------------------------------------------------------
'---    Initialize Objects
'-------------------------------------------------------------------------------#>

$Global:sArgs = $args[0]
$Global:ScriptName = $MyInvocation.MyCommand.Name

<#-------------------------------------------------------------------------------
'---    Configure
'-------------------------------------------------------------------------------#>



<#-------------------------------------------------------------------------------
'---    Install
'-------------------------------------------------------------------------------#>


Function Start-Install {
	
	$SecureBootEnabled = Read-Registry -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State" -Name "UEFISecureBootEnabled"
	
	If ($SecureBootEnabled -eq 0) {
		Set-TSVar -Variable "SecureBootState" -Value "Disabled"
	}
	elseif ($SecureBootEnabled -eq 1) {
		Set-TSVar -Variable "SecureBootState" -Value "Enabled"
	}
	elseif ($SecureBootEnabled -eq 42) {
		Set-TSVar -Variable "SecureBootState" -Value "Unknown"
	}
	
}


<#-------------------------------------------------------------------------------
'---    UnInstall
'-------------------------------------------------------------------------------#>

Function Start-Uninstall{

   
}

<#-------------------------------------------------------------------------------
'---    Functions
'-------------------------------------------------------------------------------#>

Function Start-WPFApp{

   
    
}


<#-------------------------------------------------------------------------------
'---    Start
'-------------------------------------------------------------------------------#>

Import-Module -WarningAction SilentlyContinue "$PSScriptRoot\ScriptLibrary1.22.psm1"
Set-GlobalVariables
Start-Log
Write-Log "  Runtime: $Runtime" -Type 1
Set-Mode
#$Null = $Global:Form.ShowDialog() #Uncomment for WPF Apps
End-Log


<#-------------------------------------------------------------------------------
'---    Function Templates
'-------------------------------------------------------------------------------#>


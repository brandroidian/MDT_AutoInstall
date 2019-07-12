<#
' // ***************************************************************************
' // 
' // FileName:  LTIWizard.ps1
' //            
' // Version:   1
' //            
' // Usage:     powershell.exe -file .\LTIWizard.ps1
' //          
' //
' //            
' // Created:   1.0 (2016.04.27)
' //            Brandon Hilgeman
' //            brandon.hilgeman@gmail.com
' // ***************************************************************************
#>


<#-------------------------------------------------------------------------------
'---    Initialize Objects
'-------------------------------------------------------------------------------#>

$sArg = $args[0]
$Global:sScriptName = $MyInvocation.MyCommand.Name
$Global:LogFile = $Null

<#-------------------------------------------------------------------------------
'---    Configure
'-------------------------------------------------------------------------------#>

$sPublisher = ""
$sProductName = ""
$sProductVersion = ""
$sProductSearch = ""

<#-------------------------------------------------------------------------------
'---    Install
'-------------------------------------------------------------------------------#>


Function Start-Install {
    
    Get-XAML -spath "$PSScriptRoot\LTIWizard\MainWindow.xaml" -bvariables $True
    
}


<#-------------------------------------------------------------------------------
'---    UnInstall
'-------------------------------------------------------------------------------#>

Function Start-Uninstall{

    Run-Uninstall -sName $sProductSearch -sVersion $sProductVersion
}

<#-------------------------------------------------------------------------------
'---    Functions
'-------------------------------------------------------------------------------#>

Function Start-WPFApp{
    #Hide the task sequence progress window
    Hide-TSProgress

    #Get the computername
    If($sRunTime -eq "STANDALONE"){
        $OSDComputerName = Get-ComputerName
    }
    ElseIf($sRunTime -eq "SCCM"){
        $OSDComputerName = Get-TSVar -Variable "OSDCOMPUTERNAME"
    }
    ElseIf($sRunTime -eq "MDT"){
        $OSDComputerName = $sArg
    }
    Else{
        $OSDComputerName = ""
    }

    #Add ComputerName to Option1 text field
    Add-GUIText -WPFVariable $WPFtextBox_Option1 -Text $OSDComputerName

    #Set focus on Option1 text field
    $WPFtextbox_Option1.Focus()

    #Set variable for Settings.ini
    $Global:SettingsFile = Get-IniContent "$PSScriptRoot\Files\Settings.ini"

    #Set variable for ShowLogo (boolean)
    $DisplayLogo = $SettingsFile["Settings"]["ShowLogo"]

    #Show or Hide the logo
    If($DisplayLogo -eq "True"){
        Show-GUI -WPFVariable $WPFImage_Logo
    }
    Else{
        Hide-GUI -WPFVariable $WPFImage_Logo
    }

    #Set variable for ShowTSList
    $ShowTSList = $SettingsFile["Settings"]["ShowTSList"]

    #Show or hide the 'Show TS List' checkbox (MDT Only)
    If(($sRuntime -eq "MDT" -or $sRuntime -eq "Standalone") -and ($ShowTSList -eq "True")){
        Show-GUI -WPFVariable $WPFcheckbox_ShowTS
    }
    Else{
        Hide-GUI -WPFVariable $WPFcheckbox_ShowTS
    }

    #Set variables for ShowOptions
    $ShowComputerName = $SettingsFile["Settings"]["ShowComputerName"]
    $Global:ShowOption1 = $SettingsFile["Settings"]["ShowOption1"]
    $Global:ShowOption2 = $SettingsFile["Settings"]["ShowOption2"]
    $Global:ShowOption3 = $SettingsFile["Settings"]["ShowOption3"]
    $Global:ShowOption4 = $SettingsFile["Settings"]["ShowOption4"]
    $Global:ShowOption5 = $SettingsFile["Settings"]["ShowOption5"]
    $Global:ShowOption6 = $SettingsFile["Settings"]["ShowOption6"]
    $Global:ShowOption7 = $SettingsFile["Settings"]["ShowOption7"]
    $Global:ShowOption8 = $SettingsFile["Settings"]["ShowOption8"]
    $Global:ShowOption9 = $SettingsFile["Settings"]["ShowOption9"]
    $Global:ShowOption10 = $SettingsFile["Settings"]["ShowOption10"]
    $Global:ShowOption11 = $SettingsFile["Settings"]["ShowOption11"]
    $Global:ShowOption12 = $SettingsFile["Settings"]["ShowOption12"]

    #Set variables for Options
    $Global:Option1 = $SettingsFile["Options"]["Option1"]
    $Global:Option2 = $SettingsFile["Options"]["Option2"]
    $Global:Option3 = $SettingsFile["Options"]["Option3"]
    $Global:Option4 = $SettingsFile["Options"]["Option4"]
    $Global:Option5 = $SettingsFile["Options"]["Option5"]
    $Global:Option6 = $SettingsFile["Options"]["Option6"]
    $Global:Option7 = $SettingsFile["Options"]["Option7"]
    $Global:Option8 = $SettingsFile["Options"]["Option8"]
    $Global:Option9 = $SettingsFile["Options"]["Option9"]
    $Global:Option10 = $SettingsFile["Options"]["Option10"]
    $Global:Option11 = $SettingsFile["Options"]["Option11"]
    $Global:Option12 = $SettingsFile["Options"]["Option12"]

    #Set variables for Option Task Sequence Variables
    $Global:OptionVar1 = $SettingsFile["TSVariables"]["OptionVar1"]
    $Global:OptionVar2 = $SettingsFile["TSVariables"]["OptionVar2"]
    $Global:OptionVar3 = $SettingsFile["TSVariables"]["OptionVar3"]
    $Global:OptionVar4 = $SettingsFile["TSVariables"]["OptionVar4"]
    $Global:OptionVar5 = $SettingsFile["TSVariables"]["OptionVar5"]
    $Global:OptionVar6 = $SettingsFile["TSVariables"]["OptionVar6"]
    $Global:OptionVar7 = $SettingsFile["TSVariables"]["OptionVar7"]
    $Global:OptionVar8 = $SettingsFile["TSVariables"]["OptionVar8"]
    $Global:OptionVar9 = $SettingsFile["TSVariables"]["OptionVar9"]
    $Global:OptionVar10 = $SettingsFile["TSVariables"]["OptionVar10"]
    $Global:OptionVar11 = $SettingsFile["TSVariables"]["OptionVar11"]
    $Global:OptionVar12 = $SettingsFile["TSVariables"]["OptionVar12"]
    
    #Set varibales for Option Defaults (Page 1 comboboxes only)
    $Global:OptionDefault2 = $SettingsFile["Defaults"]["OptionDefault2"]
    $Global:OptionDefault3 = $SettingsFile["Defaults"]["OptionDefault3"]
    $Global:OptionDefault4 = $SettingsFile["Defaults"]["OptionDefault4"]
    $Global:OptionDefault5 = $SettingsFile["Defaults"]["OptionDefault5"]
    $Global:OptionDefault6 = $SettingsFile["Defaults"]["OptionDefault6"]

    #Show or hide the options
    #Get the keys for options 2-6 to populate the drop-downs and set default value
    If(!($ShowOption1 -eq "True")){
        Hide-GUI -WPFVariable $WPFtextbox_Option1
        Hide-GUI -WPFVariable $WPFtextbox_Option1Text
    }

    If(!($ShowOption2 -eq "True")){
        Hide-GUI -WPFVariable $WPFcombobox_Option2
        Hide-GUI -WPFVariable $WPFtextbox_Option2Text
    }
    Else{
        ForEach($Key in $SettingsFile["Option2"].Keys | Sort-Object){
            Add-GUIText -WPFVariable $WPFcomboBox_Option2 -Text $Key
        }
        $WPFcomboBox_Option2.SelectedIndex = $OptionDefault2
    }

    If(!($ShowOption3 -eq "True")){
        Hide-GUI -WPFVariable $WPFcombobox_Option3
        Hide-GUI -WPFVariable $WPFtextbox_Option3Text
    }
    Else{
        ForEach($Key in $SettingsFile["Option3"].Keys | Sort-Object){
            Add-GUIText -WPFVariable $WPFcomboBox_Option3 -Text $Key
        }
        $WPFcomboBox_Option3.SelectedIndex = $OptionDefault3
    }

    If(!($ShowOption4 -eq "True")){
        Hide-GUI -WPFVariable $WPFcombobox_Option4
        Hide-GUI -WPFVariable $WPFtextbox_Option4Text
    }
    Else{
        ForEach($Key in $SettingsFile["Option4"].Keys | Sort-Object){
            Add-GUIText -WPFVariable $WPFcomboBox_Option4 -Text $Key
        }
        $WPFcomboBox_Option4.SelectedIndex = $OptionDefault4
    }

    If(!($ShowOption5 -eq "True")){
        Hide-GUI -WPFVariable $WPFcombobox_Option5
        Hide-GUI -WPFVariable $WPFtextbox_Option5Text
    }
    Else{
        ForEach($Key in $SettingsFile["Option5"].Keys | Sort-Object){
            Add-GUIText -WPFVariable $WPFcomboBox_Option5 -Text $Key
        }
        $WPFcomboBox_Option5.SelectedIndex = $OptionDefault5
    }
    If(!($ShowOption6 -eq "True")){
        Hide-GUI -WPFVariable $WPFcombobox_Option6
        Hide-GUI -WPFVariable $WPFtextbox_Option6Text
    }
    Else{
        ForEach($Key in $SettingsFile["Option6"].Keys | Sort-Object){
            Add-GUIText -WPFVariable $WPFcomboBox_Option6 -Text $Key
        }
        $WPFcomboBox_Option6.SelectedIndex = $OptionDefault6
    }
    If(!($ShowOption7 -eq "True")){
        Hide-GUI -WPFVariable $WPFtextbox_Option7
        Hide-GUI -WPFVariable $WPFtextbox_Option7Text
    }
    If(!($ShowOption8 -eq "True")){
        Hide-GUI -WPFVariable $WPFtextbox_Option8
        Hide-GUI -WPFVariable $WPFtextbox_Option8Text
    }
    If(!($ShowOption9 -eq "True")){
        Hide-GUI -WPFVariable $WPFtextbox_Option9
        Hide-GUI -WPFVariable $WPFtextbox_Option9Text
    }
    If(!($ShowOption10 -eq "True")){
        Hide-GUI -WPFVariable $WPFtextbox_Option10
        Hide-GUI -WPFVariable $WPFtextbox_Option10Text
    }
    If(!($ShowOption11 -eq "True")){
        Hide-GUI -WPFVariable $WPFtextbox_Option11
        Hide-GUI -WPFVariable $WPFtextbox_Option11Text
    }
    If(!($ShowOption12 -eq "True")){
        Hide-GUI -WPFVariable $WPFtextbox_Option12
        Hide-GUI -WPFVariable $WPFtextbox_Option12Text
    }

    #Add text to the option text boxes
    Add-GUIText -WPFVariable $WPFtextbox_Option1Text -Text $Option1
    Add-GUIText -WPFVariable $WPFtextbox_Option2Text -Text $Option2
    Add-GUIText -WPFVariable $WPFtextbox_Option3Text -Text $Option3
    Add-GUIText -WPFVariable $WPFtextbox_Option4Text -Text $Option4
    Add-GUIText -WPFVariable $WPFtextbox_Option5Text -Text $Option5
    Add-GUIText -WPFVariable $WPFtextbox_Option6Text -Text $Option6
    Add-GUIText -WPFVariable $WPFtextbox_Option7Text -Text $Option7
    Add-GUIText -WPFVariable $WPFtextbox_Option8Text -Text $Option8
    Add-GUIText -WPFVariable $WPFtextbox_Option9Text -Text $Option9
    Add-GUIText -WPFVariable $WPFtextbox_Option10Text -Text $Option10
    Add-GUIText -WPFVariable $WPFtextbox_Option11Text -Text $Option11
    Add-GUIText -WPFVariable $WPFtextbox_Option12Text -Text $Option12
            
    #Show page 1 as default, hide page 2 as default
    Show-GUI -WPFVariable $WPFGrid_Page1
    Hide-GUI -WPFVariable $WPFGrid_Page2

    #Display correct buttons for default page 1
    Enable-GUI -WPFVariable $WPFButton_Cancel
    Enable-GUI -WPFVariable $WPFButton_Next
    Disable-GUI -WPFVariable $WPFButton_Back
    Disable-GUI -WPFVariable $WPFButton_Finish

    #If options 7-12 are not to be shown, disable 'Next' button and enable 'Finish' button
    If(($ShowOption7 -eq "False") -and ($ShowOption8 -eq "False") -and ($ShowOption9 -eq "False") -and ($ShowOption10 -eq "False") -and ($ShowOption10 -eq "False") -and ($ShowOption11 -eq "False") -and ($ShowOption12 -eq "False")){
        Disable-GUI -WPFVariable $WPFbutton_Next
        Enable-GUI -WPFVariable $WPFbutton_Finish    
    }

    #Click Next
    $WPFbutton_Next.Add_Click({
        Hide-GUI -WPFVariable $WPFGrid_Page1
        Show-GUI -WPFVariable $WPFGrid_Page2
        Disable-GUI -WPFVariable $WPFButton_Next
        Enable-GUI -WPFVariable $WPFButton_Back
        Enable-GUI -WPFVariable $WPFButton_Finish
        $WPFtextbox_Option7.Focus()
    
    })

    #Click Back
    $WPFbutton_Back.Add_Click({
        Hide-GUI -WPFVariable $WPFGrid_Page2
        Show-GUI -WPFVariable $WPFGrid_Page1
        Enable-GUI -WPFVariable $WPFButton_Next
        Disable-GUI -WPFVariable $WPFButton_Back
        Disable-GUI -WPFVariable $WPFButton_Finish
    })
    
    
    #Click Finish
    $WPFbutton_Finish.Add_Click({
        
        #Get text/selections entered for Options
        If($ShowOption1 -eq "True"){
            $SelectedOption1 = $WPFtextbox_Option1.Text
            $SelectedOptionVariable1 = $WPFtextbox_Option1.Text
        }
        Else{$SelectedOptionVariable1 = " "}

        If($ShowOption2 -eq "True"){
            $SelectedOption2 = $WPFcomboBox_Option2.SelectedItem
            $SelectedOptionVariable2 = $SettingsFile["Option2"]["$SelectedOption2"]
        }
        Else{$SelectedOptionVariable2 = " "}

        If($ShowOption3 -eq "True"){
            $SelectedOption3 = $WPFcomboBox_Option3.SelectedItem
            $SelectedOptionVariable3 = $SettingsFile["Option3"]["$SelectedOption3"]
        }
        Else{$SelectedOptionVariable3 = " "}

        If($ShowOption4 -eq "True"){
            $SelectedOption4 = $WPFcomboBox_Option4.SelectedItem
            $SelectedOptionVariable4 = $SettingsFile["Option4"]["$SelectedOption4"]
        }
        Else{$SelectedOptionVariable4 = " "}

        If($ShowOption5 -eq "True"){
            $SelectedOption5 = $WPFcomboBox_Option5.SelectedItem
            $SelectedOptionVariable5 = $SettingsFile["Option5"]["$SelectedOption5"]
        }
        Else{$SelectedOptionVariable5 = " "}

        If($ShowOption6 -eq "True"){
            $SelectedOption6 = $WPFcomboBox_Option6.SelectedItem
            $SelectedOptionVariable6 = $SettingsFile["Option6"]["$SelectedOption6"]
        }
        Else{$SelectedOptionVariable6 = " "}

        If($ShowOption7 -eq "True"){
            $SelectedOption7 = $WPFtextbox_Option7.Text
            $SelectedOptionVariable7 = $WPFtextbox_Option7.Text
        }
        Else{$SelectedOptionVariable7 = " "}

        If($ShowOption8 -eq "True"){
            $SelectedOption8 = $WPFtextbox_Option8.Text
            $SelectedOptionVariable8 = $WPFtextbox_Option8.Text
        }
        Else{$SelectedOptionVariable8 = " "}

        If($ShowOption9 -eq "True"){
            $SelectedOption9 = $WPFtextbox_Option9.Text
            $SelectedOptionVariable9 = $WPFtextbox_Option9.Text
        }
        Else{$SelectedOptionVariable9 = " "}

        If($ShowOption10 -eq "True"){
            $SelectedOption10 = $WPFtextbox_Option10.Text
            $SelectedOptionVariable10 = $WPFtextbox_Option10.Text
        }
        Else{$SelectedOptionVariable10 = " "}

        If($ShowOption11 -eq "True"){
            $SelectedOption11 = $WPFtextbox_Option11.Text
            $SelectedOptionVariable11 = $WPFtextbox_Option11.Text
        }
        Else{$SelectedOptionVariable11 = " "}

        If($ShowOption12 -eq "True"){
            $SelectedOption12 = $WPFtextbox_Option12.Text
            $SelectedOptionVariable12 = $WPFtextbox_Option12.Text
        }
        Else{$SelectedOptionVariable12 = " "}

        #Ensure computername isn't blank
        If($SelectedOptionVariable1 -eq ""){
            $Msg =  MsgBox -Message "You must enter a valid computer name!" -Title "Computer Name?" -Buttons 0 -Icon "Exclamation"
            Return
        }

        #Check if 'Show TS List' checkbox is checked or not (for MDT only)
        If($WPFcheckBox_ShowTS.IsChecked -eq "True"){
            $cTSList = "NO"
        }
        Else{
            $cTSList = "YES"
        }
                
        Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Clicked Finish Button"

        #Get 'ShowConfirmation' value from Settings.ini
        $ShowConfirmation = $SettingsFile["Settings"]["ShowConfirmation"]

        #If 'ShowConfirmation' not equals 'False' then display confirmation dialog
        If(!($ShowConfirmation -eq "False")){
            $Msg =  MsgBox -Message "Continue Deployment with these Settings?" -Title "Continue Deployment?" -Buttons 4 -Icon "Information"
        }

        #If Confirmation button = 'Yes'
        If($Msg -eq "Yes"){
            #If Running in SCCM, set SCCM Task Sequence variables
            If($sRunTime -eq "SCCM"){
                If($OptionVar1 -ne ""){
                    Set-TSVar -Variable $OptionVar1 -Value $SelectedOptionVariable1
                }
                If($OptionVar2 -ne ""){
                    Set-TSVar -Variable $OptionVar2 -Value $SelectedOptionVariable2
                }
                If($OptionVar3 -ne ""){
                    Set-TSVar -Variable $OptionVar3 -Value $SelectedOptionVariable3
                }
                If($OptionVar4 -ne ""){
                    Set-TSVar -Variable $OptionVar4 -Value $SelectedOptionVariable4
                }
                If($OptionVar5 -ne ""){
                    Set-TSVar -Variable $OptionVar5 -Value $SelectedOptionVariable5
                }
                If($OptionVar6 -ne ""){
                    Set-TSVar -Variable $OptionVar6 -Value $SelectedOptionVariable6
                }
                If($OptionVar7 -ne ""){
                    Set-TSVar -Variable $OptionVar7 -Value $SelectedOptionVariable7
                }
                If($OptionVar8 -ne ""){
                    Set-TSVar -Variable $OptionVar8 -Value $SelectedOptionVariable8
                }
                If($OptionVar9 -ne ""){
                    Set-TSVar -Variable $OptionVar9 -Value $SelectedOptionVariable9
                }
                If($OptionVar10 -ne ""){
                    Set-TSVar -Variable $OptionVar10 -Value $SelectedOptionVariable10
                }
                If($OptionVar11 -ne ""){
                    Set-TSVar -Variable $OptionVar11 -Value $SelectedOptionVariable11
                }
                If($OptionVar12 -ne ""){
                    Set-TSVar -Variable $OptionVar12 -Value $SelectedOptionVariable12
                }
            }
            #Else assuming running standalone or MDT write results to ini file
            #Ini file will be read by UserExit.vbs script defined in CustomSettings.ini
            Else{
                #If running in MDT
                If($sRunTime -eq "MDT"){
                    $iniFile = "X:\MININT\SMSOSD\OSDLOGS\LTIAnswer.ini"
                }
                #Else, assuming running standalone
                ElseIf($sRunTime -eq "STANDALONE"){
                    $iniFile = "C:\MININT\SMSOSD\OSDLOGS\LTIAnswer.ini"
                }
                #Forcefully create ini file if it doesn't exist
                If(!(Test-Path -Path $iniFile)){
                    New-Item -ItemType file -Path $iniFile -Force
                }
                
                #Set content to be written to ini file
                $INISection = @{"cTSLIST".ToUpper()=$cTSList.ToUpper();$OptionVar1.ToUpper()=$SelectedOptionVariable1.ToUpper();$OptionVar2.ToUpper()=$SelectedOptionVariable2.ToUpper();$OptionVar3.ToUpper()=$SelectedOptionVariable3.ToUpper();$OptionVar4.ToUpper()=$SelectedOptionVariable4.ToUpper();$OptionVar5.ToUpper()=$SelectedOptionVariable5.ToUpper();$OptionVar6.ToUpper()=$SelectedOptionVariable6.ToUpper();$OptionVar7.ToUpper()=$SelectedOptionVariable7.ToUpper();$OptionVar8.ToUpper()=$SelectedOptionVariable8.ToUpper();$OptionVar9.ToUpper()=$SelectedOptionVariable9.ToUpper();$OptionVar10.ToUpper()=$SelectedOptionVariable10.ToUpper();$OptionVar11.ToUpper()=$SelectedOptionVariable11.ToUpper();$OptionVar12.ToUpper()=$SelectedOptionVariable12.ToUpper()}
                
                #Create 'MDTLTI' section in ini file
                $NewINIContent = @{"MDTLTI"=$INISection}

                #Write data to ini file with ASCII encoding
                Out-IniFile -InputObject $NewINIContent -FilePath $iniFile -Encoding ASCII -Force
            }
            End-Log
            Exit 0
            $Form.Close()
        }
    })

    #Click cancel
    $WPFbutton_Cancel.Add_Click({
        Write-Log -Message "  $($MyInvocation.MyCommand.Name):: Clicked Cancel Button"
        $Msg =  MsgBox -Message "Are you sure you want to cancel? `n`r`n`rThis will cause the deployment to fail." -Title "Cancel Deployment?" -Buttons 4
        If($Msg -eq "Yes"){
            End-Log
            Exit 42
            $Form.Close()
        }
    })
}

<#-------------------------------------------------------------------------------
'---    Start
'-------------------------------------------------------------------------------#>

Import-Module -WarningAction SilentlyContinue "$PSScriptRoot\ScriptLibrary1.15.psm1"
Set-GlobalVariables
Start-Log
Write-Log "  Runtime: $sRuntime" -Type 1
Set-Mode
$Null = $Global:Form.ShowDialog() #Uncomment for WPF Apps
End-Log


<#-------------------------------------------------------------------------------
'---    Function Templates
'-------------------------------------------------------------------------------#>

#IsSoftwareInstalled -sProduct "Microsoft*" -sVersion "*"

#Get-ComputerName

#MsgBox -sMessage "" -sTitle "" -Buttons ""

#Test-Ping -Hostname "DT2UAXXXXXX"

#RunInstall -sCMD "" -sArg ""

#Copy-File -sSource "$PSSCriptRoot\Install\Example.txt" -sDestination "C:\Users\Public\Desktop\Example.txt"

#Copy-Folder -sSource "$PSSCriptRoot\Install\Example\" -sDestination "C:\Users\Public\Desktop\Example\"

#Delete-Object -sPath "C:\Users\Public\Desktop\Example.url"

#Delete-Object -sPath "C:\Users\Public\Desktop\Example"

#AD_ManageGroupADSI -sDomain "contoso.org" -sFunction "Add" -sType "User" -sName "czt20b" -sGroup "GROUP_1" -sADUser "contoso\abc123a" -sADPass "P@ssw0rd"

#Get-XAML -sPath "$PSScriptRoot\Example\MainWindow.xaml"

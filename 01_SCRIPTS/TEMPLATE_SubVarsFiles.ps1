# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : Main Processing Script
# ===================================================================================


# -----------------------------------------------------------------------------------
#   STARTING TASKS
# -----------------------------------------------------------------------------------
# Grab Config Variables
. "$PSScriptRoot\_EPM_Config.ps1" -Process "SubVarsFilesTest_$(Get-Random -Maximum 1000)" -ExecEnvironment "TEST"

#Starts All Processing, Ensures Clean Logs, Logs into EPM Automate
EPM_Start-Process


# -----------------------------------------------------------------------------------
#   PROCESSING TASKS
# -----------------------------------------------------------------------------------

#Sub Var Testing
# -- Get Variables
$SV_EPMAutoTest = EPM_Get-SubVar -Name "EPMAutomate_TestingVar" -KeepQuotes

# -- Set Variables
EPM_Set-SubVar -Name "EPMAutomate_TestingVar" -Value "AutoTest_Run2" -WrapQuotes

# Export All Variables to custom file
EPM_Export-SubVars -Path "$EPM_PATH_FILES_OUT\SubVarTesting.txt"


#File Testing
# -- Upload
EPM_Upload-File -Path "$EPM_PATH_FILES_IN\Testing_General.txt"
EPM_Upload-File -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC.txt" -DataManagement

# -- Download
EPM_Get-File -Name "Testing_General.txt"
EPM_Get-File -Name "TEST_DM_NUMERIC.txt" -Path "$EPM_PATH_AUTO\ZZ_Testing" 
EPM_Get-File -Name "FileDoesNotExist.txt"

# -----------------------------------------------------------------------------------
#   FINISHING TASKS
# -----------------------------------------------------------------------------------

#End Processing, Archive Logs, Send Notification
EPM_End-Process -NotifyLevel SUCCESS
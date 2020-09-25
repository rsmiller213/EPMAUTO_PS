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
$SV = EPM_Export-SubVars
"Before : $($SV.ALL.EPMAutomate_TestingVar)" | EPM_Log-Item -WriteHost
# -- Set Variables & Re-Export
$SV = EPM_Set-SubVar -Name "EPMAutomate_TestingVar" -Value "AutoTest_Run$(Get-Random -Maximum 10)" -WrapQuotes
"After : $($SV.ALL.EPMAutomate_TestingVar)" | EPM_Log-Item  -WriteHost

#File Testing
# -- Upload
EPM_Upload-File -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC.txt" -DataManagement
# -- Upload Error
EPM_Upload-File -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC2.txt" -DataManagement

# -- Download
EPM_Get-File -Name "TEST_DM_NUMERIC.txt" -Path "$EPM_PATH_AUTO\ZZ_Testing" 
EPM_Get-File -Name "FileDoesNotExist.txt"

# -----------------------------------------------------------------------------------
#   FINISHING TASKS
# -----------------------------------------------------------------------------------

#End Processing, Archive Logs, Send Notification
EPM_End-Process -NotifyLevel "SUCCESS"
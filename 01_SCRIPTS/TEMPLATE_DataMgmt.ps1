# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : Data Management Testing
# ===================================================================================


# -----------------------------------------------------------------------------------
#   STARTING TASKS
# -----------------------------------------------------------------------------------

# Grab Config Variables
. "$PSScriptRoot\_EPM_Config.ps1" -Process "DataMgmtTest_$(Get-Random -Maximum 1000)" -ExecEnvironment "TEST"

#Ensure Clean Logs
EPM_Start-Process

# -----------------------------------------------------------------------------------
#   PROCESSING TASKS
# -----------------------------------------------------------------------------------

#Data Management Testing
    # Should be Successful
#EPMAPI_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-15" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC.txt"
    # Should have Kickouts and show a warning
EPMAPI_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-15" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC_ERR.txt"
    # Should have a fatal error
#EPMAPI_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-125" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC_ERR.txt" -StopOnError
#Write-Host "Above is run with StopOnError switch this should not display" -ForegroundColor Cyan

#epmautomate runDataRule "LR_OP_TEST_NUM" "Oct-125" "Mar-16" REPLACE STORE_DATA "TEST_DM_NUMERIC_ERR.txt"


# -----------------------------------------------------------------------------------
#   FINISHING TASKS
# -----------------------------------------------------------------------------------

#Close and Archive Log
EPM_End-Process -NotifyLevel SUCCESS

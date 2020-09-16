# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : BR Testing
# ===================================================================================


# -----------------------------------------------------------------------------------
#   STARTING TASKS
# -----------------------------------------------------------------------------------

# Grab Config Variables
. "$PSScriptRoot\_EPM_Config.ps1" -Process "ErrorHandling_$(Get-Random -Maximum 1000)" -ExecEnvironment "TEST"

#Ensure Clean Logs
EPM_Start-Process

# -----------------------------------------------------------------------------------
#   PROCESSING TASKS
# -----------------------------------------------------------------------------------

# DATA MANAGEMENT
# --- Success
EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-15" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC.txt"
# --- Kickouts - Numeric
#EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-15" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC_ERR.txt"
# --- Kickouts - Text
EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_TEXT" -StartPeriod "Oct-15" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_TEXT_ERR.txt"
# --- Bad Input File
#EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-151" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERICzzzz.txt"
# --- Bad Period
#EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-151" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC.txt"

# -----------------------------------------------------------------------------------
#   FINISHING TASKS
# -----------------------------------------------------------------------------------

#Close and Archive Log
EPM_End-Process -NotifyLevel SUCCESS

$global:EPM_TASK_LIST | Select TASK_ID,TASK_PARENT,TASK_LEVEL,TASK_STATUS,TASK_NAME,TASK_COMMAND,TIME_START,TIME_END,TIME_ELAPSED, TASK_ERROR_MSG | Format-Table -AutoSize

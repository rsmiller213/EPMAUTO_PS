# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : Full EPMAUTO_PS Testing
# ===================================================================================


# -----------------------------------------------------------------------------------
#   STARTING TASKS
# -----------------------------------------------------------------------------------

# Grab Config Variables
. "$PSScriptRoot\_EPM_Config.ps1" -Process "FullTesting_$(Get-Random -Maximum 1000)" -ExecEnvironment "TEST"

#Ensure Clean Logs
EPM_Start-Process

# -----------------------------------------------------------------------------------
#   PROCESSING TASKS
# -----------------------------------------------------------------------------------

# BUSINESS RULES 
# --- Success
EPM_Execute-EPMATask -TaskName "Run Rule TEST_RULE" -TaskCommand "runBusinessRule" -TaskDetails "TEST_RULE RTP_VERSION=`"Adopted Budget`" RTP_SCENARIO=`"Operating Budget`" RTP_YEARS=`"FY15`",`"FY16`""
# --- Wrong RTP Member
EPM_Execute-EPMATask -TaskName "Run Rule TEST_RULE Bad RTP" -TaskCommand "runBusinessRule" -TaskDetails "TEST_RULE RTP_VERSION=`"Adopted Budget1`" RTP_SCENARIO=`"Operating Budget`" RTP_YEARS=`"FY15`",`"FY16`""
# --- Non-Existing Rule
EPM_Execute-EPMATask -TaskName "Run Rule RULE_DOES_NOT_EXIST" -TaskCommand "runBusinessRule" -TaskDetails "RULE_DOES_NOT_EXIST RTP_VERSION=`"$Ver`" RTP_SCENARIO=$Scen RTP_YEARS=`"FY15`",`"FY16`""

# SUBSTITUTION VARIABLES
# --- Success
EPM_Set-SubVar -Name "EPMAutomate_TestingVar" -Value ("AutoTest_$(Get-Random -Maximum 1000)") -WrapQuotes
$SV_EPMAutoTest = EPM_Get-SubVar -Name "EPMAutomate_TestingVar" -KeepQuotes

# FILE HANDLING
# --- Upload : Success
EPM_Upload-File -Path "$EPM_PATH_FILES_IN\Testing_General.txt"
# --- Upload : File Does Not Exist
EPM_Upload-File -Path "$EPM_PATH_FILES_IN\UL_DoesNotExist.txt"
# --- Download : Success
EPM_Get-File -Name "Testing_General.txt"
# --- Download : File Does Not Exist
EPM_Get-File -Name "DL_DoesNotExist.txt"

# DATA MANAGEMENT
# --- Success
EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-15" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC.txt"
# --- Kickouts - Numeric
EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-15" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC_ERR.txt"
# --- Kickouts - Text
EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_TEXT" -StartPeriod "Oct-15" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_TEXT_ERR.txt"
# --- Bad Input File
EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-151" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERICzzzz.txt"
# --- Bad Period
EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-151" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC.txt"

# MAINTENANCE

# Backups
# --- Success
#EPM_Backup-Application -New

# Security
# --- Success
EPM_Export-Security
# --- Success
EPM_Import-Security -ImportFile "$EPM_FILE_SECURITY"
# --- Kickouts in File
EPM_Import-Security -ImportFile "$EPM_PATH_FILES_IN\CurrentAppSecurityWithError.csv"

# Admin Mode
# --- Success
EPM_Maintain-App -Task ModeAdmin
# --- Success
EPM_Maintain-App -Task ModeUser

# Cube Maint
# --- Success
EPM_Maintain-ASOCube -Cube "TestASO" -Task MergeSlicesRemoveZero
# --- Success
EPM_Maintain-BSOCube -Cube "TestBSO" -Task RestructureCube
# --- Invalid Cube
EPM_Maintain-BSOCube -Cube "TestBSO123" -Task RestructureCube


# -----------------------------------------------------------------------------------
#   FINISHING TASKS
# -----------------------------------------------------------------------------------

#Close and Archive Log
EPM_End-Process -NotifyLevel SUCCESS

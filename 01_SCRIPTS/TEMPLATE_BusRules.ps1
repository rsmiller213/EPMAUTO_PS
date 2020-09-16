# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : BR Testing
# ===================================================================================


# -----------------------------------------------------------------------------------
#   STARTING TASKS
# -----------------------------------------------------------------------------------

# Grab Config Variables
. "$PSScriptRoot\_EPM_Config.ps1" -Process "BRTest_$(Get-Random -Maximum 1000)" -ExecEnvironment "TEST"

#Ensure Clean Logs
EPM_Start-Process

# -----------------------------------------------------------------------------------
#   PROCESSING TASKS
# -----------------------------------------------------------------------------------

# With RTP
EPM_Execute-EPMATask -TaskName "Run Rule TEST_RULE" -TaskCommand "runBusinessRule TEST_RULE RTP_VERSION=`"Adopted Budget`" RTP_SCENARIO=`"Operating Budget`" RTP_YEARS=`"FY15`",`"FY16`""
$Scen = "`"Operating Budget`""
$Ver = "Adopted Budget"
#With RTP using PowerShell Vars
EPM_Execute-EPMATask -TaskName "Run Rule TEST_RULE" -TaskCommand "runBusinessRule TEST_RULE RTP_VERSION=`"$Ver`" RTP_SCENARIO=$Scen RTP_YEARS=`"FY15`",`"FY16`""
#With Error
EPM_Execute-EPMATask -TaskName "Run Rule RULE_DOES_NOT_EXIST" -TaskCommand "runBusinessRule RULE_DOES_NOT_EXIST RTP_VERSION=`"Adopted Budget`" RTP_SCENARIO=`"Operating Budget`" RTP_YEARS=`"FY15`",`"FY16`""


# -----------------------------------------------------------------------------------
#   FINISHING TASKS
# -----------------------------------------------------------------------------------

#Close and Archive Log
EPM_End-Process -NotifyLevel SUCCESS
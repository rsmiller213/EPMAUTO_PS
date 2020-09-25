# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : Example Full Process
# ===================================================================================


# -----------------------------------------------------------------------------------
#   STARTING TASKS
# -----------------------------------------------------------------------------------

# Grab Config Variables
. "$PSScriptRoot\_EPM_Config.ps1" -Process "FullProcess" -ExecEnvironment "TEST"


#$Task1 = $TaskList.addTask(@{name = "Task 1"})
#$Task2 = $TaskList.addTask(@{name = "Task 2"})
#$Task21 = $TaskList.addTask(@{name = "Task 2.1"; level = 1; parentId = 2})
#$TaskList.updateTask(3,@{status = "STARTING"})

$Task1 = $EPM_TASKLIST.addTask(@{name = "Hey";status = "STARTING"})
Start-Sleep -Seconds 5
$Task1.updateTask(@{status="SUCCESS"})
$EPM_TASKLIST.Tasks | ForEach-Object{$_.display()}
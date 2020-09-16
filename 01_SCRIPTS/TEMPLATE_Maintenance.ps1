# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : Main Processing Script
# ===================================================================================


# -----------------------------------------------------------------------------------
#   STARTING TASKS
# -----------------------------------------------------------------------------------

# Grab Config Variables
. "$PSScriptRoot\_EPM_Config.ps1" -Process "MaintTest_$(Get-Random -Maximum 1000)" -ExecEnvironment "TEST"

#Ensure Clean Logs
EPM_Start-Process

# -----------------------------------------------------------------------------------
#   PROCESSING TASKS
# -----------------------------------------------------------------------------------

#Security Testing
# -- Export Security
EPM_Export-Security
# -- Import Security
EPM_Import-Security -ImportFile "$EPM_FILE_SECURITY"
EPM_Import-Security -ImportFile "$EPM_PATH_FILES_IN\CurrentAppSecurityWithError.csv"

#Backup & Maintenance Testing
# -- Run & Download Backup
EPM_Backup-Application
# -- App Maint
# ---- Maint Mode Admin
EPM_Maintain-App -Task ModeAdmin
# ---- Maint Mode User
EPM_Maintain-App -Task ModeUser
# -- ASO Maint
# ---- Merge Slices Remove Zero
EPM_Maintain-ASOCube -Cube "CherryRd" -Task MergeSlicesRemoveZero
# -- BSO Maint
# ---- Restructure DB
EPM_Maintain-BSOCube -Cube "Operatin" -Task RestructureCube





# -----------------------------------------------------------------------------------
#   FINISHING TASKS
# -----------------------------------------------------------------------------------

#Close and Archive Log
EPM_End-Process -NotifyLevel SUCCESS

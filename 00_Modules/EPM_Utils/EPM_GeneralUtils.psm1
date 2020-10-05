# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 09-02-2020
#   Purpose : House general functions to use during EPM Automation (Oracle EPBCS)
# ===================================================================================
# -----------------------------------------------------------------------------------
#   EPM UTILITIES - PROCESSING
# -----------------------------------------------------------------------------------



function EPM_Start-Process{
    <#
        .SYNOPSIS
        Creates the Archive Folder
        Resets & Starts the LOG_FULL.log with time stamps
        Writes the Global Variables to LOG_FULL.log
        Logs into EPMAutomate
    
        .EXAMPLE
        EPM_Start-Log
    #>
        param(
            #Does not login
            [Switch]$NoLogin
        )
    
        #Check if Process Running, otherwise create the flag
        if (Test-Path $EPM_PROCESS_RUNNING_FLAG) {
            "$EPM_PROCESS is already running, please wait for it to finish to start again." | EPM_Log-Item -WriteHost -LogType "ERROR"
            "If you believe this is in error, ensure that you have called EPM_End-Process at the end of your script" | EPM_Log-Item -WriteHost -LogType "ERROR"
            "or delete $EPM_PROCESS_RUNNING_FLAG" | EPM_Log-Item -WriteHost -LogType "ERROR"
            break
        } else {
            Set-Content $EPM_PROCESS_RUNNING_FLAG -Value (Get-Date)
        }
    
        #Ensure Folder Structure Exists
        $PathVars = Get-Variable EPM_PATH* -Exclude EPM_PATH_CURRENT_ARCHIVE,EPM_PATH_AUTO -ValueOnly
        ForEach ($item in $PathVars) {
            New-Item -ItemType Directory -Force -Path $item | Out-Null
        }
    
        New-Item -ItemType Directory -Force -Path "$EPM_PATH_CURRENT_ARCHIVE\LOGS" | Out-Null
        New-Item -ItemType Directory -Force -Path "$EPM_PATH_CURRENT_ARCHIVE\FILES\INBOUND" | Out-Null
        New-Item -ItemType Directory -Force -Path "$EPM_PATH_CURRENT_ARCHIVE\FILES\OUTBOUND" | Out-Null
    
    
        #Write Starter Sequence to LOG_FULL.log
        Set-Content -Path $EPM_LOG_FULL -Value ""
        Set-Content -Path $EPM_LOG_ERROR -Value ""
        Remove-Item -Path "$EPM_PATH_SCRIPTS\*.log"
        "=================================" | EPM_Log-Item -Clean
        "==            START            ==" | EPM_Log-Item -Clean
        "=================================" | EPM_Log-Item -Clean
        "START TIME : $(EPM_Get-TimeStamp -StampType CLEAN)" | EPM_Log-Item -Clean
        "EXECUTING SCRIPT : $((Get-Variable MyInvocation -Scope 2).Value.MyCommand.Name)" | EPM_Log-Item -Clean
    
        #Display Starter Variables
        "======= STARTER VARIABLES =======" | EPM_Log-Item -Clean
        if ($EPMAPI_USED) {
            (Get-Variable EPM_ENV,EPM_PROCESS,EPM_USER,EPM_PASSFILE,EPM_DOMAIN,EPM_DATACENTER,EPM_URL,EPM_LOG_FULL,EPM_PATH_CURRENT_ARCHIVE,EPMAPI_PASSFILE,EPMAPI_PLN_BASE_URI,EPMAPI_MIG_BASE_URI,EPMAPI_DMG_BASE_URI | Format-Table -AutoSize | Out-String).trim() | EPM_Log-Item -Clean
        } else {
            (Get-Variable EPM_ENV,EPM_PROCESS,EPM_USER,EPM_PASSFILE,EPM_DOMAIN,EPM_DATACENTER,EPM_URL,EPM_LOG_FULL,EPM_PATH_CURRENT_ARCHIVE | Format-Table -AutoSize | Out-String).trim() | EPM_Log-Item -Clean
        }
        
        "" | EPM_Log-Item -Clean -IncludeSeparator
    
        #Login
        if(-not($NoLogin)) {EPM_Execute-EPMATask -TaskName "EPM Automate Login $EPM_ENV" -TaskCommand "login" -TaskDetails "$EPM_USER $EPM_PASSFILE $EPM_URL $EPM_DOMAIN" -StopOnError}
    }
    
    
    
function EPM_End-Process{
    <#
        .SYNOPSIS
        Logs out of EPM Automate (unless -NoLogout is specified)
        Adds the End Log with TimeStamp and Elapsed Time
        Copies Logs/Data to Archive Folder & Compresses Archive
        Applies Archive Retention Policy defined in 01_SCRIPTS\_EPMConfig.ps1
        Removes the Running Flag
    
        .EXAMPLE
        EPM_End-Process
    #>
    
        param(
            #Determines when to send notification
            [ValidateSet("NONE","ERROR","WARNING","SUCCESS")][String]$NotifyLevel = $EPM_EMAIL_NOTIFYLEVEL,
            #Will not logout
            [Switch]$NoLogout
        )
    
        #Logout of EPM Automate
        if(-not($NoLogout)) {EPM_Execute-EPMATask -TaskName "EPM Automate Logout $EPM_ENV" -TaskCommand "logout" -IgnoreError}
    
        #Write Ending Sequence to the Full Log
        "ELAPSED TIME : $(EPM_Get-ElapsedTime -StartTime $EPM_PROCESS_START)" | EPM_Log-Item -Clean
        "END TIME : $(EPM_Get-TimeStamp -StampType CLEAN)" | EPM_Log-Item -Clean
        "=================================" | EPM_Log-Item -Clean
        "==             END             ==" | EPM_Log-Item -Clean
        "=================================" | EPM_Log-Item -Clean
    
        # Create Task Log
        $TL = $EPM_TASKLIST.getTasks("status",@("ERROR","WARNING"))
        $TL = $TL.getTasks("command","ne","")
        if ($TL.Tasks.Count){
            $TL.Tasks | ForEach-Object{$_.display()} | EPM_Log-Item -Clean -IncludeSeparator -LogFile $EPM_LOG_ERROR
        } else {
            Remove-Item -Path "$EPM_LOG_ERROR" -Force
        }
        
        #Copy Logs/Data to Archive & Compress
        Copy-Item -Path "$EPM_PATH_LOGS\*" -Destination "$EPM_PATH_CURRENT_ARCHIVE\LOGS"
        Copy-Item -Path "$EPM_PATH_FILES_IN\*" -Destination "$EPM_PATH_CURRENT_ARCHIVE\FILES\INBOUND"
        Copy-Item -Path "$EPM_PATH_FILES_OUT\*" -Destination "$EPM_PATH_CURRENT_ARCHIVE\FILES\OUTBOUND"
        Move-Item -Path "$EPM_PATH_SCRIPTS\*.log" -Destination "$EPM_PATH_CURRENT_ARCHIVE\LOGS" -Force
        Compress-Archive -Path "$EPM_PATH_CURRENT_ARCHIVE\*" -DestinationPath "$EPM_PATH_CURRENT_ARCHIVE.zip" -Force
        #Cleanup
        Remove-Item -Path $EPM_PATH_CURRENT_ARCHIVE -Recurse -Force
        Remove-Item -Path "$EPM_PATH_LOGS\*" -Exclude LOG_*
    
        #Remove old Original Security if exists
        Get-ChildItem "$EPM_PATH_FILES_OUT\*OriginalSecurity.csv" | Sort-Object CreationTime -Descending | Select-Object -Skip 1 | Remove-Item -Force
    
        #Apply Archive Retention Policy set in 01_SCRIPTS\_EPM_Config.ps1
        if ($EPM_ARCHIVES_RETAIN_POLICY -eq "NUM"){
            Get-ChildItem "$EPM_PATH_ARCHIVES\*.zip" -Recurse | Sort-Object CreationTime -Descending | Select-Object -Skip $EPM_ARCHIVES_RETAIN_NUM | Remove-Item -Force
        } elseif ($EPM_ARCHIVES_RETAIN_POLICY -eq "DAYS") {
            Get-ChildItem "$EPM_PATH_ARCHIVES\*.zip" -Recurse | Where-Object {$_.LastWriteTime -lt  (Get-Date).AddDays(-$EPM_ARCHIVES_RETAIN_NUM)} | Remove-Item -Force
        }
    
        Remove-Item -Path $EPM_PROCESS_RUNNING_FLAG -Force
        if ($NotifyLevel -ne "NONE") { EPM_Send-Notification -NotifyLevel $NotifyLevel }
        Remove-Item -Path "$EPM_PATH_SCRIPTS\*.log"
    }



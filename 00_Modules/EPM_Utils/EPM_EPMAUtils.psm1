# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : House EPMAutomate functions to use during EPM Automation (Oracle EPBCS)
# ===================================================================================
# -----------------------------------------------------------------------------------
#   EPM UTILITIES - EPMAutomate Processes
# -----------------------------------------------------------------------------------
function EPM_Execute-EPMATask{
<#
    .SYNOPSIS
    Will check if logged in, if not will login
    Will properly log the Task / Sub-Task
    Will execute the supplied EPM Automate Command 
        (do not include "epmautomate" at the beginning)
    Will properly handle any errors

    .EXAMPLE
    $SubVarOut = EPM_Execute-EPMATask `
                    -TaskName "Export ALL Subvars" `
                    -TaskCommand "getSubstVar ALL" `
                    -ReturnOut -StopOnError
    Will execute the command "epmautomate getSubstVar ALL"
    Will Return the output to the $SubVarOut variable
    Will stop all processing on error

    .EXAMPLE
    EPM_Execute-EPMATask `
        -TaskName "Download Log"   `
        -TaskCommand ("downloadFile `"$LogWebPath.Trim()`"") `
        -TaskLevel 1
    Will execute the command "epmautomate downloadFile "<FilePath>"" as a Sub-Task

    .EXAMPLE
    EPM_Execute-EPMATask `
        -TaskName "EPMAutomate Login Task" `
        -TaskCommand "login $EPM_USER $EPM_PASSFILE $EPM_URL $EPM_DOMAIN" `
        -StopOnError
    Will execute the command "epmautomate $EPM_USER $EPM_PASSFILE $EPM_URL $EPM_DOMAIN"
    Will stop all processing on error
#>
    Param(
        #[MANDATORY] Name of the Task for Logging Purposes
        [parameter(Mandatory=$true)][String]$TaskName,
        #[MANDATORY] EPMAutomate base command to be executed 
        #   (do not include epmautomate)
        [parameter(Mandatory=$true)][String]$TaskCommand,
        #Command details to be included with the base command
        [String]$TaskDetails,
        #Allows for the output to be returned to the caller to be used,
        #   will also write to log
        [switch]$ReturnOut,
        #Will hide the task from the console
        [switch]$HideTask,
        #Will stop the entire process if there is an error
        [switch]$StopOnError,
        #Will ignore & not log any error if there is one
        [switch]$IgnoreError,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    if (-not (Test-Path $EPM_PROCESS_RUNNING_FLAG) ) {
        "Please run EPM_Start-Process to login" | EPM_Log-Item -WriteHost -LogType "ERROR"
        EPM_End-Process
        break
    }

    $Task = $EPM_TASKLIST.addTask(@{
        name = $TaskName;
        command = $TaskCommand;
        details = $TaskDetails;
        level = $TaskLevel;
        parentId = $ParentID;
        hideTask = $HideTask;
    })

    # Execute Command
    if (-not $ReturnOut) {
        if (!$HideTask){
            Invoke-Expression "$EPM_AUTO_CALL $TaskCommand $TaskDetails" | EPM_Log-Item
        } else {
            Invoke-Expression "$EPM_AUTO_CALL $TaskCommand $TaskDetails" | Out-Null
        }
    } else {
        $ReturnString = Invoke-Expression "$EPM_AUTO_CALL $TaskCommand $TaskDetails"
    }
    $LastStatus = $LASTEXITCODE
    # Check for Errors & Log Error Code
    if ( $LastStatus -ne 0 ) {
        if ($IgnoreError) {
            $Task.updateTask(@{status = "IGNORED"})
        } else {
            $Task.updateTask(@{status = "ERROR"})
        }
        
    } else {
        $Task.updateTask(@{status = "SUCCESS"})
    }

    if ( $StopOnError -and ($Task.status -ne "SUCCESS") ) {EPM_End-Process; break}
    if ( $ReturnOut ) { return $ReturnString }

}


function EPM_Export-SubVars{
<#
    .SYNOPSIS
    Will export substitution variables from the PBCS Instance and 
        put them in a hashtable (dictionary)

    .EXAMPLE
    EPM_Export-SubVars
    will export all subvars

    EPM_Export-SubVars -OutFile -PlanType "FINPLAN"
    will export all substitution variables for FINPLAN
#>
    Param(
        #Plan Type to export the substitution variables for, by default 
        #    it exports application level (ALL)
        [String]$PlanType = "ALL",
        #Hides the Task from Console
        [Switch]$HideTask,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0
    )

    $SVTemp = EPM_Execute-EPMATask `
                    -TaskName "Export ALL Subvars" `
                    -TaskCommand "getSubstVar" `
                    -TaskDetails "$PlanType" `
                    -TaskLevel $TaskLevel `
                    -ReturnOut -StopOnError -HideTask:($HideTask)
    $SV = [ordered]@{}
    $SVTemp = ($SVTemp | Select-Object -Skip 1 | Select-Object -SkipLast 1 | Sort-Object)
    foreach($row in $SVTemp) {
        $temp = $row.trim().split("=").split(".")
        $SV[$temp[0]] += [ordered]@{$temp[1] = $temp[2]}
    }
    $SV = $SV
    return $SV
}



function EPM_Set-SubVar{
<#
    .SYNOPSIS
    Will Set a Substitution variable to the value you provide

    .EXAMPLE
    EPM_Set-SubVar -Name "ACT_CUR_MO" -Value "Jan" -WrapQuotes
    Will set the ALL.ACT_CUR_MO sub var value to "Jan"

    .EXAMPLE
    EPM_Set-SubVar -Name "ACT_CUR_MO" -Value "Jan" -PlanType "FINPLAN"
    Will set the FINPLAN.ACT_CUR_MO sub var value to Jan (no quotes)
#>
    Param(
        #[MANDATORY] Substitution variable name to set
        [parameter(Mandatory=$true)][String]$Name,
        #[MANDATORY] Substitution variable value to set
        [parameter(Mandatory=$true)][String]$Value,
        #Plan Type for where the SV is stored (i.e. ALL = Global)
        [String]$PlanType ="ALL",
        #Forces double quotes around the variable
        [Switch]$WrapQuotes,
        #Hides the Task from Console
        [Switch]$HideTask,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0
    )

    if ($WrapQuotes) {
        EPM_Execute-EPMATask `
            -TaskName "Set SubVar $PlanType $Name to `"$Value`"" `
            -TaskCommand "setSubstVars" `
            -TaskDetails ("$PlanType $Name=" + '"\""' + $Value + '\"""') `
            -TaskLevel $TaskLevel `
            -StopOnError -HideTask:$HideTask
    } else {
        EPM_Execute-EPMATask `
            -TaskName "Set SubVar $PlanType $Name to $Value" `
            -TaskCommand "setSubstVars" `
            -TaskDetails "$PlanType $Name=$Value" `
            -TaskLevel $TaskLevel `
            -StopOnError -HideTask:$HideTask
    }
    return EPM_Export-SubVars -HideTask

}



function EPM_Get-File{
<#
    .SYNOPSIS
    Will test if a file exists in the Application, do not include pathing in file name
    If found will download it

    .EXAMPLE
    EPM_Get-File -Name "Testing123.txt"
    Will download file named "Testing123.txt" to $EPM_PATH_FILES_OUT\Testing123.txt

    .EXAMPLE
    EPM_Get-File -Name "FINPLAN_1000.txt"
    Will download file named "outbox/logs/FINPLAN_1000.txt" to 
        $EPM_PATH_FILES_OUT\Testing123.txt
#>
    param(
        #FileName to Download
        [String]$Name,
        #Path to move file To, by default will move to 03_Files\OUTBOUND
        [String]$Path = $EPM_PATH_FILES_OUT,
        #Hides the Task from Console
        [Switch]$HideTask,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $ListFiles = (EPM_Execute-EPMATask `
                        -TaskName "Export List of Files" `
                        -TaskCommand "listfiles" `
                        -ReturnOut `
                        -HideTask
                        )
    $DLCount = 0
    ForEach($line in $ListFiles) {
        if ($line.Trim().contains($Name)) {
            EPM_Execute-EPMATask `
                -TaskName "Download $($line.Trim())" `
                -TaskCommand "downloadFile" `
                -TaskDetails ("`"$($line.Trim())`"") `
                -TaskLevel $TaskLevel `
                -ParentID $ParentID `
                -HideTask:$HideTask
            if ($LASTEXITCODE -eq 0) 
                { Move-Item "$EPM_PATH_SCRIPTS\$Name*" `
                    -Destination "$Path" -Force -ErrorAction Ignore }
            $DLCount += 1
        }
    }

    if ($DLCount -eq 0) 
        {EPM_Execute-EPMATask `
            -TaskName "Download $Name" `
            -TaskCommand "downloadFile" `
            -TaskDetails ("`"$($Name.Trim())`"") `
            -TaskLevel $TaskLevel `
            -ParentID $ParentID `
            -HideTask:$HideTask
            }

}



function EPM_Upload-File{
<#
    .SYNOPSIS
    First checks if file exists, if it does it will delete
    Then uploads provided file

    .EXAMPLE
    EPM_Upload-File -Path "$EPM_PATH_FILES_IN\Testing123.txt"
    Will check if Testing123.txt exists, if it does it will delete & upload

    .EXAMPLE
    EPM_Upload-File -Path "$EPM_PATH_FILES_IN\Testing123.txt" -DataManagement
    Will check if Testing123.txt exists in DM Inbox, if it does it will delete & upload
#>
    Param(
        #[MANDATORY] File to Uplaod
        [parameter(Mandatory=$true)][String]$Path,
        #Will check / upload to data management inbox
        [Switch]$DataManagement,
        #Hides the Task from Console
        [Switch]$HideTask,
        #Stops on Error
        [Switch]$StopOnError,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    #Get File Name
    $FileName = $Path.Substring($Path.LastIndexOf('\')+1)
    
    if ($DataManagement){
        $TaskDetails = ("`"inbox\$FileName`"")
    } else {
        $TaskDetails = ("`"$FileName`"")
    }

    #Delete before upload (this may fail if its a new file)
    EPM_Execute-EPMATask `
        -TaskName "Delete Before Upload" `
        -TaskCommand "deleteFile" `
        -TaskDetails $TaskDetails `
        -TaskLevel $TaskLevel `
        -ParentID $ParentID `
        -HideTask:$true `
        -IgnoreError
    
    if ($DataManagement){
        $TaskDetails = ("`"$Path`" inbox")
    } else {
        $TaskDetails = ("`"$Path`"")
    }

    EPM_Execute-EPMATask `
        -TaskName "Upload the File $FileName" `
        -TaskCommand "uploadFile" `
        -TaskDetails $TaskDetails `
        -TaskLevel $TaskLevel `
        -ParentID $ParentID `
        -HideTask:$HideTask `
        -StopOnError:$StopOnError
}



function EPM_Move-FileToInstance{
<#
    .SYNOPSIS
    Will move a file from the provided source to the target 
        environment, will handle logins / logouts etc

    .EXAMPLE
    EPM_Move-FileToInstance `
        -SourceEnv "PROD" `
        -TargetEnv "TEST" `
        -FileName "20-04-05_PRD2TST_Testing" `
        -IsSnapshot
    Will move a snapshot named "20-04-05_PRD2TST_Testing" 
        from PROD to TEST
#>
    param(
        #The source environment where the file/snapshot resides
        [parameter(Mandatory=$true)]
        [ValidateSet("PROD","TEST")][String]$SourceEnv,
        #The target environment where you want to move it to
        [parameter(Mandatory=$true)]
        [ValidateSet("PROD","TEST")][String]$TargetEnv,
        #The name of the file/snapshot (Including Path)
        [parameter(Mandatory=$true)][String]$FileName,
        #The Source URL
        [String]$SourceURL,
        #The Source Env username
        [String]$SourceUser = $EPM_USER,
        #The Source Env password file
        [String]$SourcePassfile = $EPM_PASSFILE,
        #The Target URL
        [String]$TargetURL,
        #The Target Env username
        [String]$TargetUser = $EPM_USER,
        #The Target Env password file
        [String]$TargetPassfile = $EPM_PASSFILE,
        #True/False if this is a snapshot
        [switch]$IsSnapshot,
        #Hides the Task from Console
        [Switch]$HideTask,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0
    )

    $Task = $EPM_TASKLIST.addTask(@{
        name = "Moving File $FileName from $SourceEnv to $TargetEnv";
        level = $TaskLevel;
        parentId = $ParentID;
        hideTask = $HideTask;
    })

    # Set Proper URLs
    if (!$TargetURL){
        if ($TargetEnv -eq "PROD") {$TargetURL = $EPM_URL_PROD}
        elseif ($TargetEnv -eq "TEST") {$TargetURL = $EPM_URL_TEST}
    }
    if (!$SourceURL){
        if ($SourceEnv -eq "PROD") {$SourceURL = $EPM_URL_PROD}
        elseif ($SourceEnv -eq "TEST") {$SourceURL = $EPM_URL_TEST}
    }

    #Check Current Environment & Login if necessary
    if ( $EPM_ENV -ne $TargetEnv ) {
        # Logout 
        EPM_Execute-EPMATask `
            -TaskName "Logout of $EPM_ENV for File Move" `
            -TaskCommand "logout" `
            -TaskLevel ($Task.level + 1) `
            -ParentID ($Task.id) `
            -HideTask:$HideTask
        # Login to Target Env
        EPM_Execute-EPMATask `
            -TaskName "Log into $TargetEnv for File Move" `
            -TaskCommand "login" `
            -TaskDetails "$TargetUser $TargetPassfile $TargetURL $EPM_DOMAIN" `
            -TaskLevel ($Task.level + 1) `
            -ParentID ($Task.id) `
            -HideTask:$HideTask `
            -StopOnError 
    }

    #Delete from Target
    EPM_Execute-EPMATask `
        -TaskName "Delete $FileName from $TargetEnv" `
        -TaskCommand "deleteFile" `
        -TaskDetails "$FileName" `
        -TaskLevel ($Task.level + 1) `
        -ParentID ($Task.id) `
        -HideTask:$HideTask `
        -IgnoreError
    
    #Move File
    if ($IsSnapshot) {
        EPM_Execute-EPMATask `
            -TaskName "Moving Snapshot $FileName From $SourceEnv To $TargetEnv" `
            -TaskCommand "copySnapshotFromInstance" `
            -TaskDetails "$FileName $SourceUser $SourcePassfile $SourceURL $EPM_DOMAIN" `
            -TaskLevel ($Task.level + 1)  `
            -ParentID ($Task.id) `
            -HideTask:$HideTask `
            -StopOnError
    } else {
        EPM_Execute-EPMATask `
            -TaskName "Moving File $FileName From $SourceEnv To $TargetEnv" `
            -TaskCommand "copyFileFromInstance" `
            -TaskDetails "$FileName $SourceUser $SourcePassfile $SourceURL $EPM_DOMAIN $FileName" `
            -TaskLevel ($Task.level + 1)  `
            -ParentID ($Task.id) `
            -HideTask:$HideTask `
            -StopOnError 
    }

    # Return login status to what it was before
    if ($EPM_ENV -eq $SourceEnv) {
        EPM_Execute-EPMATask `
            -TaskName "Logout of $TargetEnv to restore access to $SourceEnv" `
            -TaskCommand "logout" `
            -TaskLevel ($Task.level + 1) `
            -ParentID ($Task.id) `
            -HideTask:$HideTask
        EPM_Execute-EPMATask `
            -TaskName "Restore Access to $SourceEnv" `
            -TaskCommand "login" `
            -TaskDetails "$SourceUser $SourcePassfile $SourceURL $EPM_DOMAIN" `
            -TaskLevel ($Task.level + 1)  `
            -ParentID ($Task.id) `
            -HideTask:$HideTask `
            -StopOnError 
    }

    $Task.updateTask(@{status = "SUCCESS"},(!$HideTask))
}


function EPM_Execute-LoadRule{
<#
    .SYNOPSIS
    Will execute a Data Mgmt load rule and download 
        the log after completion and parse for kickouts

    .EXAMPLE
    EPM_Execute-LoadRule `
        -LoadRule "LR_OP_TEST_NUM" `
        -StartPeriod "Oct-15" `
        -EndPeriod "Mar-16" `
        -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC.txt"
#>
    param(
        #[MANDATORY] Name of the Load Rule
        [parameter(Mandatory=$true)][String]$LoadRule,
        #[MANDATORY] Starting Period (DM Format, i.e. Oct-20 = October FY20)
        [parameter(Mandatory=$true)][String]$StartPeriod,
        #[MANDATORY] Ending Period (DM Format, i.e. Oct-20 = October FY20)
        [parameter(Mandatory=$true)][String]$EndPeriod,
        #Path to the Load file (including file)
        [String]$Path,
        #Import Mode to Use
        #   REPLACE - Will delete & replace all DM staging records
        #   APPEND - Will add the records to DM Staging
        #   RECALCULATE - Skip the import, but re-run the mappings
        #   NONE - Skip the import & do NOT re-run the mappings
        [ValidateSet("REPLACE","APPEND","RECALCULATE","NONE")]
        [String]$ImportMode = "REPLACE",
        #Export Mode to Use
        #   STORE_DATA - Will overwrite existing data with new data
        #   ADD_DATA - Will add new data to existing in essbase
        #   SUBTRACT_DATA - Will subtract the new data from existing in essbase
        #   REPLACE_DATA - Will clear existing data before importing new
        #       NOTE : By defualt will clear the POV (Scenario, Version, Year, Period, Entity) 
        #              You should specify a "Clear Region" in load rule options
        #   NONE - Skip the export
        [ValidateSet("STORE_DATA","ADD_DATA","SUBTRACT_DATA","REPLACE_DATA","NONE")]
        [String]$ExportMode = "STORE_DATA",
        #Hides the Task from Console
        [Switch]$HideTask,
        #Stops on Error
        [Switch]$StopOnError,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )
    if ($Path) {
        $HasParentTask = $true
        $FileName = "$($Path.Substring($Path.LastIndexOf('\')+1))"
        $TaskName = "Uploading $FileName and Executing Load of $LoadRule"
        $LoadTaskName = "Loading Data via $LoadRule for $StartPeriod to $EndPeriod"

        $Task = $EPM_TASKLIST.addTask(@{
            name = $TaskName;
            level = $TaskLevel;
            parentId = $ParentID;
            hideTask = $HideTask;
        })

        $NewTaskLevel = ($Task.level + 1)
        $NewParentID = ($Task.id)

        # Path Specified, need to upload file
        EPM_Upload-File `
            -Path "$Path" `
            -TaskLevel $NewTaskLevel  `
            -ParentID $NewParentID `
            -HideTask:$HideTask `
            -DataManagement
        if ( $LASTEXITCODE -ne 0 ) {
            $Task.updateTask(@{status = "ERROR"})
            #Exit Load Process
            Return 1 | Out-Null
        }

    } else {
        $HasParentTask = $false
        $FileName = ""
        $TaskName = "Executing Load of $LoadRule"
        $LoadTaskName = "Loading Data via $LoadRule for $StartPeriod to $EndPeriod"

        $NewTaskLevel = $TaskLevel
        $NewParentID = $ParentID
    }

    #Execute Load
    EPM_Execute-EPMATask `
        -TaskName $LoadTaskName `
        -TaskCommand "runDataRule" `
        -TaskDetails ("$LoadRule $StartPeriod $EndPeriod $ImportMode $ExportMode $FileName") `
        -TaskLevel $NewTaskLevel  `
        -ParentID $NewParentID `
        -HideTask:$HideTask
    $LastStatus = $LASTEXITCODE
    $LoadTask = $EPM_TASKLIST.getTask($LoadTaskName)

    if ( $LastStatus -ne 0 ) {
        #We had an error or kickouts, determine which.
        #Grab Error Log
        $ErrorLog = (Get-ChildItem "$EPM_PATH_SCRIPTS" -Filter runDataRule*.log `
                        | Sort-Object LastWriteTime `
                        | Select-Object -Last 1)
        if ( $ErrorLog ) {
            #Error Log found, Grab DM Log
            $DMLog = [regex]::Match((Get-Content $ErrorLog.FullName),`
                        "`"logFileName`":`"([a-zA-Z\/\.\:\-_0-9]+)`"").Groups[1].Value
            $KickoutLog = [regex]::Match((Get-Content $ErrorLog.FullName),`
                            "`"outputFileName`":`"([a-zA-Z\/\.\:\-_0-9]+)`"").Groups[1].Value
            if ( $DMLog ) {
                #Parse Just File Name
                $DMLog = $DMLog.Substring($DMLog.LastIndexOf('/') + 1)
                #Download DM Log
                EPM_Get-File `
                    -Name $DMLog `
                    -Path $EPM_PATH_LOGS `
                    -TaskLevel ($LoadTask.level + 1)  `
                    -ParentID ($LoadTask.id) `
                    -HideTask:$HideTask
                if ( $LASTEXITCODE -ne 0 ) {
                    #Error Log Not Found, Exit
                    $LoadTask.updateTask(@{status = "ERROR"})
                    if ( $HasParentTask ) {$Task.updateTask(@{status = "ERROR"})}
                    Return 1 | Out-Null
                }
                #Download Kickout Log
                if ($KickoutLog) {
                    $KickoutLog = $KickoutLog.Substring($KickoutLog.LastIndexOf('/') + 1)
                    EPM_Get-File `
                        -Name $KickoutLog `
                        -Path $EPM_PATH_LOGS `
                        -TaskLevel ($LoadTask.level + 1)  `
                        -ParentID ($LoadTask.id) `
                        -HideTask:$HideTask
                    if ( $LASTEXITCODE -ne 0 ) {
                        #No Kickouts, Exit
                        $LoadTask.updateTask(@{status = "ERROR"})
                        if ( $HasParentTask ) {$Task.updateTask(@{status = "ERROR"})}
                        Return 1 | Out-Null
                        }
                    #Parse Relevant Info
                    $LoadID = ($DMLog.Substring($DMLog.LastIndexOf('/') + 1).Split("_")[1].Replace(".log","")).trim()
                    $RuleName = ([regex]::Match((Get-Content "$EPM_PATH_LOGS\$DMLog"), `
                                    "Rule Name    : (.*? )").Groups[1].Value).trim()
                    $LoadFileName = ([regex]::Match((Get-Content "$EPM_PATH_LOGS\$DMLog"), `
                                    "File Name.*: (.*?txt)").Groups[1].Value).trim()
                    $LinePrefix = "$LoadID#$RuleName#$LoadFileName"

                    foreach ( $line in (Get-Content "$EPM_PATH_LOGS\$KickoutLog") ) {
                        $Kickout=""
                        if ( $line.Contains("Error: 3303") ) {
                            $arrLine = $line.split("|")
                            $Kickout = "$($arrLine[2].trim())#$($arrLine[3].trim())"
                            "$LinePrefix#$Kickout" | EPM_Log-Item -Clean -LogFile $EPM_LOG_KICKOUTS
                        } elseif ( $line.contains("The member ")) {
                            $Kickout = [regex]::Match(($line),"(The member )(.*)( does not exist)").Groups[2].Value
                            "$LinePrefix#$Kickout#Text Data Load" | EPM_Log-Item -Clean -LogFile $EPM_LOG_KICKOUTS
                        } elseif ( $line.contains("Fetch of Driver Member ")) {
                            $Kickout = $([regex]::Match(($line),'(.*\")(.*)(\".*)').Groups[2].Value)
                            "$LinePrefix#$Kickout#Text Data Load" | EPM_Log-Item -Clean -LogFile $EPM_LOG_KICKOUTS
                        }
                    }

                    #Update Warning
                    $LoadTask.updateTask(@{
                        status = "WARNING"; 
                        errorMsg = "Review Kickouts for Load ID : $LoadID"
                        })
                    if ( $HasParentTask ) {$Task.updateTask(@{status = "WARNING"})}
                } else {
                    #Update Error
                    $LoadTask.updateTask(@{status = "ERROR";})
                    if ( $HasParentTask ) {$Task.updateTask(@{status = "ERROR"})}
                }
                
            }
        }
    } else {
        if ( $HasParentTask ) {$Task.updateTask(@{status = "SUCCESS"})}
    }

    if ( ( $LastStatus -ne 0 ) -and ( !$DMLog ) ){
        if ( $HasParentTask ) {$Task.updateTask(@{status = "ERROR"})}
        #$LoadTask.updateTask(@{status = "ERROR"})
    } 
}


# -----------------------------------------------------------------------------------
#   EPM UTILITIES - MAINTENANCE
# -----------------------------------------------------------------------------------



function EPM_Backup-Application{
<#
    .SYNOPSIS
    Will download the Artifact Snapshot and move it to a specified file path. 
    If no path is provided will default to 
    04_BACKUPS\<CurrentDateTime>-<Environemnt>-<ApplicationName>-BACKUP.zip
    Also applies the Backup Retention Policy

    .EXAMPLE
    EPM_Backup-Application
    Will download "Artifact Snapshot"
    Will move file to 04_BACKUPS\<CurrentDateTime>_<ApplicationName>_BACKUP.zip

    .EXAMPLE
    EPM_Backup-Application -New -Path "C:\Testing.zip"
    Will re-run the Artifact Snapshot and download it to C:\Testing.zip
#>
    param(
        #Will Re-Run the Export for Artifact Snapshot
        [Switch]$New,
        #Will Move Backup to Defined Path & File Name
        [String]$Path = "$EPM_PATH_BACKUPS\$(EPM_Get-TimeStamp -StampType FILE)_$EPM_ENV-$EPM_APP-BACKUP.zip",
        #Hides the Task from Console
        [Switch]$HideTask,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    if ($New) {
        $Task = $EPM_TASKLIST.addTask(@{
            name = "Backup $EPM_FINPLAN in $EPM_ENV";
            level = $TaskLevel;
            parentId = $ParentID;
            hideTask = $HideTask;
        })

    
        EPM_Execute-EPMATask `
            -TaskName "Re-Export Snapshot" `
            -TaskCommand ("exportSnapshot") `
            -TaskDetails '"Artifact Snapshot"' `
            -TaskLevel ($Task.level + 1) `
            -ParentID ($Task.id) `
            -HideTask:$HideTask
        if ($LASTEXITCODE -ne 0) {
            $Task.updateTask(@{status = "ERROR"}) 
            return 1 | Out-Null
        }

        EPM_Get-File `
            -Name "Artifact Snapshot" `
            -Path $Path `
            -TaskLevel ($LoadTask.level + 1)  `
            -ParentID ($LoadTask.id) `
            -HideTask:$HideTask
        if ($LASTEXITCODE -ne 0) {
            $Task.updateTask(@{status = "ERROR"}) 
            return 1 | Out-Null
        }

    } else {
        EPM_Get-File `
            -Name "Artifact Snapshot" `
            -Path $Path `
            -TaskLevel $TaskLevel  `
            -ParentID $ParentID `
            -HideTask:$HideTask
        if ($LASTEXITCODE -ne 0) {
            return 1 | Out-Null
        }
    }

    #Apply Backup Retention Policy set in 01_SCRIPTS\_EPM_Config.ps1
    if ($EPM_BACKUPS_RETAIN_POLICY -eq "NUM"){
        Get-ChildItem "$EPM_PATH_BACKUPS\*.zip" -Recurse | `
            Sort-Object CreationTime -Descending | `
            Select-Object -Skip $EPM_BACKUPS_RETAIN_NUM | `
            Remove-Item -Force
    } elseif ($EPM_BACKUPS_RETAIN_POLICY -eq "DAYS") {
        Get-ChildItem "$EPM_PATH_BACKUPS\*.zip" -Recurse |`
            Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-$EPM_BACKUPS_RETAIN_NUM)} | `
            Remove-Item -Force
    }

    if ( $new ) { $Task.updateTask(@{status = "SUCCESS"}) }
}


function EPM_Maintain-ASOCube{
<#
    .SYNOPSIS
    Will perform various ASO cube maintenance tasks

    .EXAMPLE
    EPM_Maintain-ASOCube -Cube "REPORT" -Task MergeSlicesKeepZero
    Will merge the data slices but retain "0" values
#>
    param(
        #Name of the cube the action is to be performed on
        [parameter(Mandatory=$true)][String]$Cube,
        #Name of the task to be performed
        [parameter(Mandatory=$true)][ValidateSet("MergeSlicesKeepZero","MergeSlicesRemoveZero")]
        [String]$Task,
        #Hides the Task from Console
        [Switch]$HideTask,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    if ($Task -eq "MergeSlicesKeepZero" ) {
        EPM_Execute-EPMATask `
            -TaskName "Merge Slices - Keep Zero" `
            -TaskCommand "mergeDataSlices" `
            -TaskDetails "`"$Cube`" keepZeroCells=false" `
            -TaskLevel $TaskLevel `
            -ParentID $ParentID `
            -HideTask:$HideTask
    } elseif ($Task -eq "MergeSlicesRemoveZero") {
        EPM_Execute-EPMATask `
            -TaskName "Merge Slices - Remove Zero" `
            -TaskCommand "mergeDataSlices" `
            -TaskDetails "`"$Cube`" keepZeroCells=true" `
            -TaskLevel $TaskLevel `
            -ParentID $ParentID `
            -HideTask:$HideTask
    }
}



function EPM_Maintain-BSOCube{
<#
    .SYNOPSIS
    Will perform various BSO cube maintenance tasks

    .EXAMPLE
    EPM_Maintain-BSOCube -Cube "FINPLAN" -Task RestructureCube
#>
    param(
        #Name of the cube the action is to be performed on
        [parameter(Mandatory=$true)][String]$Cube,
        #Name of the task to be performed
        [parameter(Mandatory=$true)][ValidateSet("RestructureCube")]
        [String]$Task,
        #Hides the Task from Console
        [Switch]$HideTask,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    if ($Task -eq "RestructureCube") {
        EPM_Execute-EPMATask `
            -TaskName "Restructure Cube" `
            -TaskCommand "restructureCube" `
            -TaskDetails "`"$Cube`"" `
            -TaskLevel $TaskLevel `
            -ParentID $ParentID `
            -HideTask:$HideTask
    }


}



function EPM_Maintain-App{
<#
    .SYNOPSIS
    Will perform various Application Maintenance Tasks

    .EXAMPLE
    EPM_Maintain-App -Task ModeAdmin
    Will set the current application mode to Admin Only
#>
    param(
        #Name of the task to be performed
        [parameter(Mandatory=$true)][ValidateSet("ModeAdmin","ModeUser")]
        [String]$Task,
        #Hides the Task from Console
        [Switch]$HideTask,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    if ($Task -eq "ModeAdmin") {
        EPM_Execute-EPMATask `
            -TaskName "Set App Mode : Admin"   `
            -TaskCommand "applicationAdminMode" `
            -TaskDetails "$true" `
            -TaskLevel $TaskLevel `
            -ParentID $ParentID `
            -HideTask:$HideTask
    } elseif ($Task -eq "ModeUser") {
        EPM_Execute-EPMATask `
            -TaskName "Set App Mode : User" `
            -TaskCommand "applicationAdminMode" `
            -TaskDetails "$false" `
            -TaskLevel $TaskLevel `
            -ParentID $ParentID `
            -HideTask:$HideTask
    }

}



function EPM_Export-Security{
<#
    .SYNOPSIS
    Will export & download application security, requires .csv extension on file

    .EXAMPLE
    EPM_Export-Security
    Will export current App Security to $EPM_PATH_FILES_OUT\CurrentAppSecurity.csv

    .EXAMPLE
    EPM_Export-Security -FileName "Testing123.csv" -Path "C:\Security"
    Will export current App Security to C:\Security\Testing123.csv
#>
    param(
        #File Name of Security to be Exported
        [String]$FileName = "CurrentAppSecurity.csv",
        #Path of the local destination of the Security File
        [String]$Path = $EPM_PATH_FILES_OUT,
        #Hides the Task from Console
        [Switch]$HideTask,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    #Handle Pathing / File Names
    if ($FileName.Contains(".csv")) {
        $FileNameCSV = $FileName
    } else {
        if ($FileName.Contains(".")) {
            $FileNameCSV = $FileName.Substring(0,($FileName.IndexOf("."))) + ".csv"
        } else {
            $FileNameCSV = "$FileName.csv"
        }
    }

    if ($Path.Contains($FileName)) {
        if ($Path.Contains($FileNameCSV)) {
            $PathWithFileName = $Path
        } else {
            $PathWithFileName = ($Path.Replace($FileName,$FileNameCSV))
        }
    } else {
        if ($Path.EndsWith("\")) {
            $PathWithFileName = "$Path$FileNameCSV"
        } else {
            $PathWithFileName = "$Path\$FileNameCSV"
        }
    }

    #Process Security
    $Task = $EPM_TASKLIST.addTask(@{
        name = "Exporting Security to $($PathWithFileName.Replace($EPM_PATH_AUTO,''))";
        level = $TaskLevel;
        parentId = $ParentID;
        hideTask = $HideTask;
    })


    EPM_Execute-EPMATask `
        -TaskName "Generate Security Export" `
        -TaskCommand "exportAppSecurity" `
        -TaskDetails "$FileNameCSV" `
        -TaskLevel ($Task.level + 1) `
        -ParentID ($Task.id) `
        -HideTask:$HideTask

    if ($LASTEXITCODE -ne 0) {
        #Export Failed
        $Task.updateTask(@{status = "ERROR"})
    } else {
        #Export Success
        EPM_Get-File `
            -Name $FileNameCSV `
            -Path $PathWithFileName `
            -TaskLevel ($Task.level + 1) `
            -ParentID ($Task.id) `
            -HideTask:$HideTask

        if ($LASTEXITCODE -ne 0) {
            #Download Failed
            $Task.updateTask(@{status = "ERROR"})
        } else {
            #Download Success
            EPM_Execute-EPMATask `
                -TaskName "Delete Sec Export From Web" `
                -TaskCommand "deleteFile" `
                -TaskDetails "$FileNameCSV" `
                -TaskLevel ($Task.level + 1) `
                -ParentID ($Task.id) `
                -HideTask:$true `
                -IgnoreError 
            $Task.updateTask(@{status = "SUCCESS"})
        }
    }

}



function EPM_Import-Security{
<#
    .SYNOPSIS
    Will Export Security and store with timestamp in 
        $EPM_PATH_FILES_OUT\<TimeStamp>_OriginalSecurity.csv
    Will Import application security

    .EXAMPLE
    EPM_Import-Security -ImportFile "$EPM_PATH_FILES_IN\UpdatedSecurity.csv"
    Will export current App Security to 
        $EPM_PATH_FILES_OUT\<TimeStamp>_OriginalSecurity.csv
    Will import $EPM_PATH_FILES_IN\UpdatedSecurity.csv and 
        write errors to $EPM_PATH_LOGS\SecurityErrors.log

    .EXAMPLE
    EPM_Import-Security 
        -ImportFile "$EPM_PATH_FILES_IN\UpdatedSecurity.csv" `
        -ErrorFile "C:\Testing\ErrorTesting.log" `
        -ClearAll
    Will export current App Security to $EPM_PATH_FILES_OUT\<TimeStamp>_OriginalSecurity.csv
    Will remove all current security and import $EPM_PATH_FILES_IN\UpdatedSecurity.csv 
        and write errors to C:\Testing\ErrorTesting.log
#>
    param(
        #[MANDATORY] File Name & Path to Security File to Import
        [parameter(Mandatory=$true)][String]$ImportFile,
        #File Name & Path of Error File, by default writes 
        #   to $EPM_PATH_LOGS\SecurityErrors.log
        [String]$ErrorFileName = "SecurityErrors.csv",
        #Specifies whether this is a sub-task or not
        [Switch]$ClearAll = $false,
        #Hides the Task from Console
        [Switch]$HideTask,
        #Specifies the Level of Task Being Run
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    #Process Security
    $Task = $EPM_TASKLIST.addTask(@{
        name = "Import Security from $($ImportFile.Replace($EPM_PATH_AUTO,''))";
        level = $TaskLevel;
        parentId = $ParentID;
        hideTask = $HideTask;
    })

    $OrigSecurity = "$(EPM_Get-TimeStamp -StampType FILE)_OriginalSecurity.csv"
    $ImportFileName = $ImportFile.Substring($ImportFile.LastIndexOf('\')+1)

    #Export Original Security
    EPM_Export-Security `
        -FileName "$OrigSecurity" `
        -TaskLevel ($TaskLevel+1) `
        -ParentID $TaskID `
        -HideTask:$HideTask

    if ($LASTEXITCODE -ne 0) {
        #Export Failed
        $Task.updateTask(@{status = "ERROR"})
    } else {
        #Export Success
        $ImportTask = $EPM_TASKLIST.addTask(@{
            name = "Import New Security";
            level = ($Task.level + 1);
            parentId = ($Task.id);
            hideTask = $HideTask;
        })
        #Upload New Security File
        EPM_Upload-File `
            -Path "$ImportFile" `
            -TaskLevel ($ImportTask.level + 1) `
            -ParentID ($ImportTask.id) `
            -HideTask:$HideTask
        if ($LASTEXITCODE -ne 0) {
            #Upload Failed
            $ImportTask.updateTask(@{status = "ERROR"}) 
            $Task.updateTask(@{status = "ERROR"})
        } else {
            #Upload Success
            EPM_Execute-EPMATask `
                -TaskName "Import Security" `
                -TaskCommand "importAppSecurity" `
                -TaskDetails "`"$ImportFileName`" `"$ErrorFileName`" clearall=$ClearAll" `
                -TaskLevel ($ImportTask.level + 1) `
                -ParentID ($ImportTask.id) `
                -HideTask:$HideTask
            if ($LASTEXITCODE -ne 0) {
                #Import Failed
                EPM_Get-File `
                    -Name "$ErrorFileName" `
                    -Path "$EPM_LOG_SECURITY" `
                    -TaskLevel ($ImportTask.level + 1) `
                    -ParentID ($ImportTask.id) `
                    -HideTask:$HideTask
                
                EPM_Execute-EPMATask `
                    -TaskName "Delete Error File from Web $ErrorFileName" `
                    -TaskCommand "deleteFile" `
                    -TaskDetails "`"$ErrorFileName`"" `
                    -TaskLevel ($ImportTask.level + 1) `
                    -ParentID ($ImportTask.id) `
                    -HideTask:$true `
                    -IgnoreError

                $ImportTask.updateTask(@{status = "ERROR"}) 
                $Task.updateTask(@{status = "ERROR"})
            } else {
                #Import Success
                $ImportTask.updateTask(@{status = "SUCCESS"}) 
                $Task.updateTask(@{status = "SUCCESS"})
            }
        }
    }
}


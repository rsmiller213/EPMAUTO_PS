# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : House common functions to use during EPM Automation (Oracle EPBCS)
# ===================================================================================

# -----------------------------------------------------------------------------------
#   EPM UTILITIES - GENERAL / LOGGING
# -----------------------------------------------------------------------------------



function EPM_Get-TimeStamp{
<#
    .SYNOPSIS
    Gets the current time stamp formatted as a string

    .EXAMPLE
    EPM_Get-TimeStamp -StampType CLEAN
    This will return a string : 04/02/20 01:46:09 PM

    .EXAMPLE
    EPM_Get-TimeStamp -StampType FILE
    This will return a string : 2020-04-02_13-46-09

    .EXAMPLE
    EPM_Get-TimeStamp
    This will return a string : [04/02/20 13:46:09]

    .EXAMPLE
    "Embedded in String : $(EPM_Get-TimeStamp -StampType CLEAN)"
    This would output the string : "Embedded in String : 04/02/20 01:46:09 PM"
#>
    Param(
        #[VALUES = CLEAN, FILE] What type of TimeStamp to return or leave blank to use LOG. See examples for what they each return
        [ValidateSet("CLEAN","FILE")][String]$StampType
    )

    if ($StampType -eq 'CLEAN') {
        return "{0:MM/dd/yy} {0:hh:mm:ss tt}" -f (Get-Date)
    } elseif ($StampType -eq 'FILE') {
        return "{0:yyyy-MM-dd}_{0:HH-mm-ss}" -f (Get-Date)
    } else {
        return "[{0:MM/dd/yy} {0:hh:mm:ss}]" -f (Get-Date)
    }
}



function EPM_Log-Task{
    Param(
        #[MANDATORY] The task name that the error occured in
        [parameter(Mandatory=$true)][String]$TaskName,
        #The base command being executed
        [String]$TaskCommand,
        #The details of the command
        [String]$TaskDetails,
        #[VALUES = START, FINISH] Determines if it should display start/end log message
        [ValidateSet("START","FINISH")][String]$TaskStage = "START",
        #[VALUES = SUCCESS, ERROR, WARNING] Outcome of Task
        [ValidateSet("SUCCESS","ERROR","WARNING")][String]$TaskStatus = "SUCCESS",
        #The starting DateTime from which the elapsed time will calculate from
        $StartTime = (Get-Date),
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Will Stop All Processing on Error
        [Switch]$StopOnError,
        #Will Add Item to EPM_TASK_OBJECT
        [Switch]$ForceAdd,
        #Will Track The Task Parent
        [Int]$ParentID,
        #Will Force an Update to Task of provided Task ID
        [Int]$UpdateTask = 0,
        [String]$OverrideError,
        [switch]$NoLog
    )


    #Determine Logging Prefix
    if ($TaskStatus -eq "SUCCESS") {
        if ($TaskLevel -eq 0) {
            $TaskPrefix = "$(EPM_Get-TimeStamp) : ==="
        } else {
            $TaskPrefix = "$(EPM_Get-TimeStamp) : ---$("--" * $TaskLevel)"
        }
    } else {
        $TaskPrefix = "$(EPM_Get-TimeStamp) : !!!$("!!" * $TaskLevel)"
    }

    #Get Timers
    $EndTime = Get-Date
    $ElapsedTime = New-TimeSpan -Start $StartTime -End $EndTime
    $ElapsedTimeStr = EPM_Get-ElapsedTime -StartTime $StartTime

    #Determine Message & Color
    if ($TaskStage -eq "START") {
        $MessageColor = "Cyan"
        if ($TaskLevel -eq 0) {
            $TaskMessage = "$TaskPrefix Starting Task : $TaskName"
        } else {
            $TaskMessage = "$TaskPrefix Starting Sub-Task : $TaskName"
        }
    } else {
        if ($TaskStatus -eq "SUCCESS") {
            $MessageColor = "Green"
            if ($TaskLevel -eq 0) {
                $TaskMessage = "$TaskPrefix Finished Task : $TaskName [Elapsed Time : $ElapsedTimeStr]"
            } else {
                $TaskMessage = "$TaskPrefix Finished Sub-Task : $TaskName [Elapsed Time : $ElapsedTimeStr]"
            }
        } elseif ($TaskStatus -eq "WARNING") {
            $MessageColor = "Yellow"
            if ($TaskLevel -eq 0) {
                $TaskMessage = "$TaskPrefix WARNING in Task : $TaskName [Elapsed Time : $ElapsedTimeStr]"
            } else {
                $TaskMessage = "$TaskPrefix WARNING in Sub-Task : $TaskName [Elapsed Time : $ElapsedTimeStr]"
            }
        } else {
            $MessageColor = "Red"
            if ($TaskLevel -eq 0) {
                $TaskMessage = "$TaskPrefix ERROR in Task : $TaskName [Elapsed Time : $ElapsedTimeStr]"
            } else {
                $TaskMessage = "$TaskPrefix ERROR in Sub-Task : $TaskName [Elapsed Time : $ElapsedTimeStr]"
            }
        }
    }

    #Parse Error Msg
    $ErrorMsg = $OverrideError
    if (-not $ErrorMsg) {
        if ($TaskStatus -ne "SUCCESS") {
            #Parse Log
            $ErrorLog = Get-ChildItem "$EPM_PATH_SCRIPTS" -Filter $TaskCommand*.log | Sort-Object LastWriteTime | Select-Object -Last 1
            if ($ErrorLog) {
                $ErrorMsg = Get-Content -Path $ErrorLog.FullName | Select-String -Pattern "^EPM.*-(.*?)" -Context 0,1000 | Out-String
                $ErrorMsg = $ErrorMsg.Trim().Replace("> ","")
            } else {
                #Error Log Doesn't Exist, grab last line from $EPM_LOG_FULL
                $ErrorMsg = Get-Content -PATH $EPM_LOG_FULL -Tail 1
                $ErrorMsg = $ErrorMsg.Substring($ErrorMsg.IndexOf("] : ") + 3).Trim()
                #Check if this is an actual error
                if ($ErrorMsg.Substring(0,3) -ne "EPM") {
                    $ErrorMsg = ""
                    #If this is an update, grab the latest task error msg
                    if ($UpdateTask -ne 0) {
                        $ErrorMsg = $global:EPM_TASK_LIST[-1].TASK_ERROR_MSG
                    }

                }
            }
        }
    }

    #Write to Log
    if (-not $NoLog) {
        Write-Host $TaskMessage -ForegroundColor $MessageColor
        $TaskMessage | Add-Content -Path $EPM_LOG_FULL
        if ($TaskStage -eq "START") {
            if ( ($TaskCommand -ne "") -or ($TaskDetails -ne "") ) { "$(EPM_Get-TimeStamp) : [COMMAND] : $TaskCommand $TaskDetails" | Add-Content -Path $EPM_LOG_FULL }
        } else {
            if ($TaskLevel -eq 0) {"$EPM_TASK_SEPARATOR" | Add-Content -Path $EPM_LOG_FULL}
        }
    }

    #Add to Task Object
    if ($UpdateTask -eq 0) {
        if ( ($TaskStage -eq "FINISH") -or ($ForceAdd)  ) {
            $TaskID = $global:EPM_TASK_LIST.Count + 1
            $global:EPM_TASK_LIST += [PSCustomObject]@{
                TASK_ID = $TaskID; 
                TASK_STATUS = $TaskStatus; 
                TASK_NAME = $TaskName; 
                TASK_COMMAND = $TaskCommand; 
                TASK_DETAILS = $TaskDetails; 
                TIME_START = $StartTime; 
                TIME_END = $EndTime; 
                TIME_ELAPSED = $ElapsedTime;  
                TASK_LEVEL=$TaskLevel;
                TASK_ERROR_MSG=$ErrorMsg;
                TASK_PARENT=$ParentID;
                }
        }
    } else {
        ForEach ($task in $global:EPM_TASK_LIST){
            if ($task.TASK_ID -eq $UpdateTask) {

                if ($OverrideError) {
                    $task.TASK_STATUS = $TaskStatus
                    $task.TASK_ERROR_MSG=$ErrorMsg;
                } else {
                    #Write-Host ("Task ID [$($task.TASK_ID)] | UpdTask [$UpdateTask] | TaskName [$TaskName] | Before : $($task.TASK_PARENT) | After : $ParentID")
                    $task.TASK_STATUS = $TaskStatus
                    $task.TASK_NAME = $TaskName; 
                    $task.TASK_COMMAND = $TaskCommand; 
                    $task.TASK_DETAILS = $TaskDetails; 
                    $task.TIME_START = $StartTime; 
                    $task.TIME_END = $EndTime; 
                    $task.TIME_ELAPSED = $ElapsedTime; 
                    $task.TASK_LEVEL=$TaskLevel;
                    $task.TASK_ERROR_MSG=$ErrorMsg;
                    $task.TASK_PARENT=$ParentID;
                }
            }

        }
    }

    if ($TaskStatus -ne "SUCCESS") {
        if($StopOnError){
            EPM_End-Process
            break
        }
    }

}



filter EPM_Log-Item{
<#
    .SYNOPSIS
    Writes Item to EPM_LOG_FULL preceeded with the timestamp for anything that is piped to it

    .EXAMPLE
    "Testing 123" | EPM_Log-Item
    Will output to the LOG_FULL : [04/06/20 08:07:53] : Testing123

    .EXAMPLE
    "Download Log" | EPM_Log-Item -IncludeSeparator
    Will output to the LOG_FULL :     [04/06/20 08:07:56] : Download Log
    Will include the Task separator after it

#>
    Param(
        #Allows for a separator to be placed after the item is logged
        [switch]$IncludeSeparator
    )

    "$(EPM_Get-TimeStamp) : $_" | Add-Content -Path $EPM_LOG_FULL
    if ($IncludeSeparator) {"$EPM_TASK_SEPARATOR" | Add-Content -Path $EPM_LOG_FULL}
}



function EPM_Get-Function{
<#
    .SYNOPSIS
    Will return the current executing function

    .EXAMPLE
    $CurFnc = EPM_Get-Function

    .EXAMPLE
    $CurFnc = EPM_Get-Function -NoLog
    Will return the current executing function
#>
    param(
        #Does not log the current / parent function
        [Switch]$NoLog
    )

    $CurFnc = (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name
    $ParFnc = (Get-Variable MyInvocation -Scope 2).Value.MyCommand.Name

    if (-not ($NoLog)) {
        Write-Verbose -Message "Running... $CurFnc | Parent Function : $ParFnc"
    }

    return $CurFnc
}



function EPM_Get-ElapsedTime{
<#
    .SYNOPSIS
    Calculates the elapsed time from a provided start DataTime to now and returns a formatted string as "00:00:00"

    .EXAMPLE
    EPM_Get-ElapsedTime -StartTime $TaskStartTime
    will return the time between $TaskStartTime and now as "00:00:00"
#>
    Param(
        #The starting DateTime from which the elapsed time will calculate from
        $StartTime
    )

    $EndTime = Get-Date
    $ElapsedTime = New-TimeSpan -Start $StartTime -End $EndTime
    return ("{0:hh\:mm\:ss}" -f $ElapsedTime)
}



function EPM_Send-Notification{
<#
    .SYNOPSIS
    Will send an email notification after the process has finished.
    Will attach relevant logs
    Will popualte the body with relevant information (i.e. Errors & Kickouts)
    To Update the From/To/CC/SMTP Server & Port refer to the _EPM_Config.ps1 file

    .EXAMPLE
    EPM_Send-Notification
#>
    Param(
        #Determines when to send notification
        [ValidateSet("ERROR","WARNING","SUCCESS")][String]$NotifyLevel = "ERROR"
    )

    $Priority = "Normal"
    $EndTime = Get-Date
    $ElapsedTime = New-TimeSpan -Start $EPM_PROCESS_START -End $EndTime
    if ($NotifyLevel -eq "SUCCESS") {$Notify = $true} else {$Notify = $false}

    #Check if Errors / Warning Exist
    $ErrCount = 0
    $WarnCount = 0
    ForEach ($obj in $global:EPM_TASK_LIST) {
        if ($obj.TASK_STATUS -eq "ERROR") { $ErrCount += 1 }
        elseif ($obj.TASK_STATUS -eq "WARNING") { $WarnCount += 1 }
    }
   
    # ==================================
    # Build Email Body
    # ==================================
    $EmailBody = @"
        <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
        <html xmlns="http://www.w3.org/1999/xhtml">
        <head>
        <title>HTML TABLE</title>
        </head><body>
"@

    # GMail Doesn't Allow "Style" html tags in the headers of an email, so we have to rig this to do the style in line
    $ErrorColor = "#eb0738"
    $WarningColor = "#eff59d"
    $SuccessColor = "#40c23e"
    $HeaderBgColor = "#0070c0"
    $TitleBgColor = "#757171"
    $TableStyle="`n<table border='0' cellspacing='0' style='border: 3px solid black; text-align: left; border-collapse: collapse;'>"
    $HeaderStyle = "border: 1px solid black; padding: 4px 3px; font-weight: bold; color:white; background: $HeaderBgColor; border-bottom: 3px solid #000000;"
    $CellStyle = "border: 1px solid black; padding: 4px 3px;"
    

    # ----------------------------------
    # Build Process Table
    # ----------------------------------
    if ( $ErrCount -ge 1) {
        $ProcessStatus = "ERROR"
        $ProcessStyle = "background: $ErrorColor;"
    } elseif ($WarnCount -ge 1) {
        $ProcessStatus = "WARNING"
        $ProcessStyle = "background: $WarningColor"
    } else {
        $ProcessStatus = "SUCCESS"
        $ProcessStyle = "background: $SuccessColor;"
    }
    $EmailBody += $TableStyle
    $EmailBody += "`n<caption style='border: 3px solid black; border-bottom: 0px; text-align: center; padding: 4px 3px; font-weight: bold; color:white; background: $TitleBgColor;'>PROCESS SUMMARY</caption>"
    $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>ENVIRONMENT</td><td style='$CellStyle text-align: center;'>$EPM_ENV</td></tr>"
    $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>PROCESS</td><td style='$CellStyle text-align: center;'>$EPM_PROCESS</td></tr>"
    $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>STATUS</td><td style='$CellStyle text-align: center; $ProcessStyle'>$ProcessStatus</td></tr>"
    $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>START TIME</td><td style='$CellStyle text-align: center;'>$($EPM_PROCESS_START.ToString("MM/dd/yy hh\:mm\:ss tt"))</td></tr>"
    $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>END TIME</td><td style='$CellStyle text-align: center;'>$($EndTime.ToString("MM/dd/yy hh\:mm\:ss tt"))</td></tr>"
    $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>ELAPSED TIME</td><td style='$CellStyle text-align: center;'>$($ElapsedTime.ToString("hh\:mm\:ss"))</td></tr>"
    #Close Processes Table
    $EmailBody += "`n</table>"
    $EmailBody += "`n<br><br>"

    # ----------------------------------
    # Build Variable Table
    # ----------------------------------
    $EmailBody += $TableStyle
    $EmailBody += "`n<caption style='border: 3px solid black; border-bottom: 0px; text-align: center; padding: 4px 3px; font-weight: bold; color:white; background: $TitleBgColor;'>PROCESS VARIABLES</caption>"
    # Build Variable Table Header
    $EmailBody += @"
    `n   <tr>
          <th style='$HeaderStyle text-align: center;'>Name</th>
          <th style='$HeaderStyle text-align: center;'>Value</th>
       </tr>
"@
    ForEach ($obj in (Get-Variable EPM_USER,EPM_URL,EPM_LOG_FULL,EPM_PATH_CURRENT_ARCHIVE | Select-Object Name,Value)) {
        $EmailBody += "`n   <tr>`n      <td style='$CellStyle'>$($obj.Name)</td><td style='$CellStyle'>$($obj.Value)</td>"
    }
    #Close Variable Table
    $EmailBody += "`n</table>"
    $EmailBody += "`n<br><br>"


    # ----------------------------------
    # Build Tasks Table
    # ----------------------------------
    $EmailBody += $TableStyle
    $EmailBody += "`n<caption style='border: 3px solid black; border-bottom: 0px; text-align: center; padding: 4px 3px; font-weight: bold; color:white; background: $TitleBgColor;'>TASK SUMMARY</caption>"
    #Build Task Table Header
    $EmailBody += @"
    `n   <tr>
          <th style='$HeaderStyle text-align: center;'>Task ID</th>
          <th style='$HeaderStyle text-align: center;'>Status</th>
          <th style='$HeaderStyle text-align: left;'>Task</th>
          <th style='$HeaderStyle text-align: center;'>Start Time</th>
          <th style='$HeaderStyle text-align: center;'>End Time</th>
          <th style='$HeaderStyle text-align: center;'>Elapsed Time</th>
          <th style='$HeaderStyle text-align: center;'>Elapsed Time %</th>
       </tr>
"@
    #Add Tasks to Table as Rows
    ForEach ($obj in $global:EPM_TASK_LIST) {
        $EmailBody += "`n   <tr>`n      <td style='$CellStyle text-align: center;'>$($obj.TASK_ID)</td>"
        if ($obj.TASK_STATUS -eq "SUCCESS"){
            $EmailBody += "`n      <td style='$CellStyle background: $SuccessColor; text-align: center;'>$($obj.TASK_STATUS)</td>"
        } elseif ($obj.TASK_STATUS -eq "ERROR"){
            $EmailBody += "`n      <td style='$CellStyle background: $ErrorColor; text-align: center;'>$($obj.TASK_STATUS)</td>"
        } elseif ($obj.TASK_STATUS -eq "WARNING"){
            $EmailBody += "`n      <td style='$CellStyle background: $WarningColor; text-align: center;'>$($obj.TASK_STATUS)</td>"
        }
        if ($obj.TASK_LEVEL -eq 0) {
            $EmailBody += "`n      <td style='$CellStyle font-weight:bold;'>$($obj.TASK_NAME)</td>"
        } else {
            $EmailBody += "`n      <td style='$CellStyle font-weight:bold; color: grey;'>$("&nbsp;&nbsp;" * $obj.TASK_LEVEL)$($obj.TASK_NAME)</td>"
        }
        $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$($obj.TIME_START.ToString("hh\:mm\:ss tt"))</td>"
        $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$($obj.TIME_END.ToString("hh\:mm\:ss tt"))</td>"
        $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$($obj.TIME_ELAPSED.ToString("hh\:mm\:ss"))</td>"
        if ($obj.TASK_LEVEL -eq 0) {
            $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$( ($obj.TIME_ELAPSED.TotalMilliseconds / $ElapsedTime.TotalMilliseconds).ToString("P"))</td>"
        } else {
            $EmailBody += "`n      <td style='$CellStyle text-align: center;'></td>"
        }
        $EmailBody += "`n   </tr>"
    }
    #Close Task Table
    $EmailBody += "`n</table>"
    $EmailBody += "`n<br><br>"

    # ----------------------------------
    # Build Error / Warning Table
    # ----------------------------------    
    if (($ErrCount + $WarnCount) -ge 1) {
            #Build Error Table
        $EmailBody += $TableStyle
        $EmailBody += "`n<caption style='border: 3px solid black; border-bottom: 0px; text-align: center; padding: 4px 3px; font-weight: bold; color:white; background: $TitleBgColor;'>ERROR SUMMARY</caption>"
        #Build Error Table Header
        $EmailBody += @"
        `n   <tr>
              <th style='$HeaderStyle text-align: center;'>Task ID</th>
              <th style='$HeaderStyle text-align: center;'>Status</th>
              <th style='$HeaderStyle text-align: center;'>Task Command</th>
              <th style='$HeaderStyle text-align: left;'>Command Details</th>
              <th style='$HeaderStyle text-align: center;'>Error Message</th>
           </tr>
"@
        #Add Errors to Table as Rows
        ForEach ($obj in $global:EPM_TASK_LIST) {
            if ( ($obj.TASK_STATUS -ne "SUCCESS") -and ($obj.TASK_COMMAND -ne "") ) {
                $EmailBody += "`n   <tr>`n      <td style='$CellStyle text-align: center;'>$($obj.TASK_ID)</td>"
                if ($obj.TASK_STATUS -eq "ERROR"){
                    $EmailBody += "`n      <td style='$CellStyle text-align: center; background: $ErrorColor;'>$($obj.TASK_STATUS)</td>"
                } elseif ($obj.TASK_STATUS -eq "WARNING"){
                    $EmailBody += "`n      <td style='$CellStyle text-align: center; background: $WarningColor;'>$($obj.TASK_STATUS)</td>"
                }
                $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$($obj.TASK_COMMAND)</td>"
                $EmailBody += "`n      <td style='$CellStyle text-align: left;'>$($obj.TASK_DETAILS)</td>"
                $EmailBody += "`n      <td style='$CellStyle text-align: left;'>$($obj.TASK_ERROR_MSG)</td>"
                $EmailBody += "`n   </tr>"
            }
        }
        #Close Error Table
        $EmailBody += "`n</table>"
        $EmailBody += "`n<br><br>"
    }

    # ----------------------------------
    # Build Kickouts Table
    # ----------------------------------
    if (Test-Path $EPM_LOG_KICKOUTS) {
        #Build Kickouts Table
        $EmailBody += $TableStyle
        $EmailBody += "<caption style='border: 3px solid black; border-bottom: 0px; text-align: center; padding: 4px 3px; font-weight: bold; color:white; background: $TitleBgColor;'>FIRST 15 KICKOUTS</caption>"
        #Build Kickouts Table Header
        $EmailBody += @"
        `n   <tr>
              <th style='$HeaderStyle text-align: center;'>Load ID</th>
              <th style='$HeaderStyle text-align: center;'>Load Rule</th>
              <th style='$HeaderStyle text-align: center;'>Load File</th>
              <th style='$HeaderStyle text-align: center;'>Kickout Member</th>
              <th style='$HeaderStyle text-align: left;'>Kickout Record</th>
           </tr>
"@
        #Add Kickouts to Table as Rows
        ForEach ($line in (Get-Content -Path $EPM_LOG_KICKOUTS | Select-Object -First 15)) {
            $arr = $line.split("#")
            $EmailBody += "`n   <tr>`n      <td style='$CellStyle text-align: center;'>$($arr[0])</td>"
            $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$($arr[1].Trim())</td>"
            $EmailBody += "`n      <td style='$CellStyle text-align: left;'>$($arr[2].Trim())</td>"
            $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$($arr[3].Trim())</td>"
            $EmailBody += "`n      <td style='$CellStyle text-align: left;'>$($arr[4].Trim())</td>"
            $EmailBody += "`n   </tr>"
        }
        #Close Kickouts Table
        $EmailBody += "`n</table>"
    }


    $EmailBody += "`n</body></html>"

    # ==================================
    # Build Email Subject
    # ==================================
    if ($ProcessStatus -eq "ERROR") {
        $Priority = "High"
        $EmailSubject = "EPM Notifier : [FAILURE] - $EPM_ENV - $EPM_PROCESS"
        if ( ("ERROR","WARNING") -Contains $NotifyLevel) { $Notify = $true }
    } elseif ($ProcessStatus -eq "WARNING") {
        $EmailSubject = "EPM Notifier : [WARNING] - $EPM_ENV - $EPM_PROCESS"
        if ( $NotifyLevel -eq "WARNING" ) { $Notify = $true }
    } else {
        $EmailSubject = "EPM Notifier : $EPM_ENV - $EPM_PROCESS has Completed Successfully"
    }

    #Identify Params
    $param = @{
        From = $EPM_EMAIL_FROM
        To = $EPM_EMAIL_TO
        Subject = $EmailSubject
        Body = $EmailBody
        Attachment = (Get-ChildItem ("$EPM_PATH_LOGS\LOG_FULL.log","$EPM_PATH_LOGS\*.csv")).FullName
        Credential = $EPM_EMAIL_CREDENTIALS
        Port = $EPM_EMAIL_PORT
        SmtpServer = $EPM_EMAIL_SERVER
        Priority = $Priority
    }

    if ($EPM_EMAIL_CC.Count -gt 0) {$param.Add("CC",$EPM_EMAIL_CC)}
    #if ($Notify -and ($EPM_EMAIL_CREDENTIALS -ne "") ) {Send-MailMessage @param -UseSsl -BodyAsHtml}

    Set-Content -Path "$EPM_PATH_SCRIPTS\_EmailBody.htm" -Value $EmailBody
}



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
        Write-Host "$EPM_PROCESS is already running, please wait for it to finish to start again." -ForegroundColor Yellow
        Write-Host "If you believe this is in error, ensure that you have called EPM_End-Process at the end of your script" -ForegroundColor Yellow
        Write-Host "or delete $EPM_PROCESS_RUNNING_FLAG" -ForegroundColor Yellow
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
    Remove-Item -Path "$EPM_PATH_SCRIPTS\*.log"
    "=================================" | Add-Content -Path $EPM_LOG_FULL
    "==            START            ==" | Add-Content -Path $EPM_LOG_FULL
    "=================================" | Add-Content -Path $EPM_LOG_FULL
    "START TIME : $(EPM_Get-TimeStamp -StampType CLEAN)" | Add-Content -Path $EPM_LOG_FULL

    #Display Starter Variables
    "======= STARTER VARIABLES =======" | Add-Content -Path $EPM_LOG_FULL
    if ($EPMAPI_USED) {
        Get-Variable EPM_ENV,EPM_PROCESS,EPM_USER,EPM_PASSFILE,EPM_DOMAIN,EPM_DATACENTER,EPM_URL,EPM_LOG_FULL,EPM_PATH_CURRENT_ARCHIVE,EPMAPI_PASSFILE,EPMAPI_PLN_BASE_URI,EPMAPI_MIG_BASE_URI,EPMAPI_DMG_BASE_URI | Format-Table -AutoSize | Out-String | Add-Content -Path $EPM_LOG_FULL
    } else {
        Get-Variable EPM_ENV,EPM_PROCESS,EPM_USER,EPM_PASSFILE,EPM_DOMAIN,EPM_DATACENTER,EPM_URL,EPM_LOG_FULL,EPM_PATH_CURRENT_ARCHIVE | Format-Table -AutoSize | Out-String | Add-Content -Path $EPM_LOG_FULL
    }
    
    "$EPM_TASK_SEPARATOR" | Add-Content -Path $EPM_LOG_FULL

    #Login
    if(-not($NoLogin)) {EPM_Execute-EPMATask -TaskName "EPM Automate Login" -TaskCommand "login" -TaskDetails "$EPM_USER $EPM_PASSFILE $EPM_URL $EPM_DOMAIN" -StopOnError}
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
    if(-not($NoLogout)) {EPM_Execute-EPMATask -TaskName "EPM Automate Logout" -TaskCommand "logout" -IgnoreError}

    #Write Ending Sequence to the Full Log
    "ELAPSED TIME : $(EPM_Get-ElapsedTime -StartTime $EPM_PROCESS_START)" | Add-Content -Path $EPM_LOG_FULL
    "END TIME : $(EPM_Get-TimeStamp -StampType CLEAN)" | Add-Content -Path $EPM_LOG_FULL
    "=================================" | Add-Content -Path $EPM_LOG_FULL
    "==             END             ==" | Add-Content -Path $EPM_LOG_FULL
    "=================================" | Add-Content -Path $EPM_LOG_FULL

    #Copy Logs/Data to Archive & Compress
    Copy-Item -Path "$EPM_PATH_LOGS\*" -Destination "$EPM_PATH_CURRENT_ARCHIVE\LOGS"
    Move-Item -Path "$EPM_PATH_SCRIPTS\*.log" -Destination "$EPM_PATH_CURRENT_ARCHIVE\LOGS" -Force
    Copy-Item -Path "$EPM_PATH_FILES_IN\*" -Destination "$EPM_PATH_CURRENT_ARCHIVE\FILES\INBOUND"
    Copy-Item -Path "$EPM_PATH_FILES_OUT\*" -Destination "$EPM_PATH_CURRENT_ARCHIVE\FILES\OUTBOUND"
    Compress-Archive -Path "$EPM_PATH_CURRENT_ARCHIVE\*" -DestinationPath "$EPM_PATH_CURRENT_ARCHIVE.zip" -Force
    Remove-Item -Path $EPM_PATH_CURRENT_ARCHIVE -Recurse -Force
    Remove-Item -Path $EPM_FILE_SUBVARS -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$EPM_PATH_LOGS\*" -Exclude LOG_*,listfiles*,subvars*

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

    Start-Sleep -Seconds 5 | Out-Null
    Remove-Item -Path "$EPM_PATH_SCRIPTS\*.log"

}



function EPM_Execute-EPMATask{
<#
    .SYNOPSIS
    Will check if logged in, if not will login
    Will properly log the Task / Sub-Task
    Will execute the supplied EPM Automate Command (do not include "epmautomate" at the beginning)
    Will properly handle any errors

    .EXAMPLE
    $SubVarOut = EPM_Execute-EPMATask -TaskName "Export ALL Subvars" -TaskCommand "getSubstVar ALL" -ReturnOut -StopOnError
        Will execute the command "epmautomate getSubstVar ALL"
        Will Return the output to the $SubVarOut variable
        Will stop all processing on error

    .EXAMPLE
    EPM_Execute-EPMATask -TaskName "Download Log" -TaskCommand ("downloadFile `"$LogWebPath.Trim()`"") -TaskLevel 1
        Will execute the command "epmautomate downloadFile "<FilePath>"" as a Sub-Task

    .EXAMPLE
    EPM_Execute-EPMATask -TaskName "EPMAutomate Login Task" -TaskCommand "login $EPM_USER $EPM_PASSFILE $EPM_URL $EPM_DOMAIN" -StopOnError
        Will execute the command "epmautomate $EPM_USER $EPM_PASSFILE $EPM_URL $EPM_DOMAIN"
        Will stop all processing on error
#>
    Param(
        #[MANDATORY] Name of the Task for Logging Purposes
        [parameter(Mandatory=$true)][String]$TaskName,
        #[MANDATORY] EPMAutomate base command to be executed (do not include epmautomate)
        [parameter(Mandatory=$true)][String]$TaskCommand,
        #Command details to be included with the base command
        [String]$TaskDetails,
        #Allows for the output to be returned to the caller to be used, will also write to log
        [switch]$ReturnOut,
        #Does not write command output to log
        [switch]$NoLog,
        #Will stop the entire process if there is an error
        [switch]$StopOnError,
        #Will ignore & not log any error if there is one
        [switch]$IgnoreError,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        [Int]$ParentID = 0
    )

    if (-not (Test-Path $EPM_PROCESS_RUNNING_FLAG) ) {
        Write-Host "Please run EPM_Start-Process to login" -ForegroundColor Red
        break
    }

    $TaskStartTime = Get-Date
    if (-not ($NoLog)) {EPM_Log-Task -TaskName $TaskName -TaskCommand $TaskCommand -TaskDetails $TaskDetails -TaskStage START -TaskLevel $TaskLevel -ParentID $ParentID}

    # Execute Command
    if ($ReturnOut) {
        $ReturnString = Invoke-Expression "$EPM_AUTO_CALL $TaskCommand $TaskDetails"
        if (-not ($NoLog) ) {
            $ReturnString | EPM_Log-Item
        }
    } else {
        if (-not ($NoLog) ) {
            Invoke-Expression "$EPM_AUTO_CALL $TaskCommand $TaskDetails" | EPM_Log-Item
        } else {
            Invoke-Expression "$EPM_AUTO_CALL $TaskCommand $TaskDetails" | Out-Null
        }
    }
    $LastStatus = $LASTEXITCODE
    # Check for Errors & Log Error Code
    if ( ($LastStatus -ne 0) -and (-not ($IgnoreError)) ) {
        if ($StopOnError) {
            if (-not ($NoLog)) {EPM_Log-Task -TaskName $TaskName -TaskCommand $TaskCommand -TaskDetails $TaskDetails -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -StopOnError -ParentID $ParentID}
        } else {
            if (-not ($NoLog)) {EPM_Log-Task -TaskName $TaskName -TaskCommand $TaskCommand -TaskDetails $TaskDetails -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -ParentID $ParentID}
        }

    } else {
        #Success
        if (-not ($NoLog)) {EPM_Log-Task -TaskName $TaskName -TaskCommand $TaskCommand -TaskDetails $TaskDetails -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus SUCCESS -StartTime $TaskStartTime -ParentID $ParentID}
    }

    if ($ReturnOut) { return $ReturnString }


}



function EPM_Export-SubVars{
<#
    .SYNOPSIS
    Will export substitution variables from the PBCS Instance and write them to a file to be used

    .EXAMPLE
    EPM_Export-SubVars
    will export application level substitution variables to 02_LOGS\Subvars.txt

    EPM_Export-SubVars -OutFile "C:\Test.txt" -PlanType "FINPLAN"
    will export all substitution variables for FINPLAN to C:\Test.txt
#>
    Param(
        #Output file path to export the substitution variables to, by default it is 02_LOGS\Subvars.txt
        [String]$Path = $EPM_FILE_SUBVARS,
        #Plan Type to export the substitution variables for, by default it exports application level (ALL)
        [String]$PlanType = "ALL",
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0
    )

    $CurFnc = EPM_Get-Function

    $SubVarOut = EPM_Execute-EPMATask -TaskName "Export $PlanType Subvars" -TaskCommand "getSubstVar" -TaskDetails "$PlanType" -ReturnOut -StopOnError -TaskLevel $TaskLevel -NoLog
    Set-Content $Path -Value $SubVarOut
}



function EPM_Get-SubVar{
<#
    .SYNOPSIS
    Will return the provided substitution variable value as a string (without quotes)
    If the requested $InFile doesn't exist or the 02_LOGS\Subvars.txt doesn't exist it will run the export

    .EXAMPLE
    EPM_Get-SubVar -Name "ACT_CUR_MO"
    will search 02_LOGS\Subvars.txt for ALL.ACT_CUR_MO and return the value

    EPM_Get-SubVar -Name "ACT_CUR_MO" -InFile "C:\Test.txt" -PlanType "FINPLAN"
    will search C:\Test.txt for FINPLAN.ACT_CUR_MO and return the value
#>
    Param(
        #[MANDATORY] Substitution variable name to retrieve
        [parameter(Mandatory=$true)][String]$Name,
        #Path to File to look for substitution variables, by default uses 02_LOGS\Subvars.txt
        [String]$Path = $EPM_FILE_SUBVARS,
        #Plan Type to Search for Substitution variables, by default uses ALL
        [String]$PlanType ="ALL",
        #Switch to keep double quotes or not
        [Switch]$KeepQuotes
    )

    $CurFnc = EPM_Get-Function

    #Export Sub Vars if the file doesn't exist
    if (-not (Test-Path -Path $Path)) {
        EPM_Export-SubVars -Path $Path -TaskLevel 1
    }
    #Parse out the substitution variable
    Get-ChildItem $Path | Select-String -Pattern "$PlanType." | Select-String -Pattern "$Name" | ForEach-Object { $SubVarOut = $_ }
    if ($KeepQuotes) {
        $SubVarOut = $SubVarOut -split '='
    } else {
        $SubVarOut = $SubVarOut -split '=' -replace '"',''
    }

    "$SubVarOut" | EPM_Log-Item

    return $SubVarOut
}


function EPM_Set-SubVar{
<#
    .SYNOPSIS
    Will Set a Substitution variable to the value you provide

    .EXAMPLE
    EPM_Set-SubVar -Name "ACT_CUR_MO" -Value "Jan" -WrapQuotes
    wW

    .EXAMPLE
    EPM_Get-SubVar -Name "ACT_CUR_MO" -InFile "C:\Test.txt" -PlanType "FINPLAN"
    will search C:\Test.txt for FINPLAN.ACT_CUR_MO and return the value
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
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0
    )

    $CurFnc = EPM_Get-Function

    if ($WrapQuotes) {
        EPM_Execute-EPMATask -TaskName "Set SubVar $PlanType $Name to `"$Value`"" -TaskCommand "setSubstVars" -TaskDetails ("$PlanType $Name=" + '"\""' + $Value + '\"""') -StopOnError -TaskLevel $TaskLevel
    } else {
        EPM_Execute-EPMATask -TaskName "Set SubVar $PlanType $Name to $Value" -TaskCommand "setSubstVars" -TaskDetails "$PlanType $Name=$Value" -StopOnError -TaskLevel $TaskLevel
    }
    EPM_Export-SubVars

}



function EPM_Test-File{
<#
    .SYNOPSIS
    Will test if a file exists in the Application
    returns True/False by default
    Use -ReturnPath switch to return the full filepath

    .EXAMPLE
    EPM_Test-File -FileName "Testing123.txt"
    Will return true if a file exists named "Testing123.txt"

    .EXAMPLE
    EPM_Test-File -FileName "FINPLAN_1000.log"
    Will list all files in the application to C:\Testing.zip
#>
    param(
        #List File Path, default is 02_LOGS\listfiles.txt
        [String]$Path = "$EPM_PATH_LOGS\listfiles.txt",
        #FileName to Check
        [String]$Name,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0
    )

    $CurFnc = EPM_Get-Function

    #Export File List
    $ListFiles = EPM_Execute-EPMATask -TaskName "Export List of Files" -TaskCommand "listfiles" -ReturnOut -NoLog
    Set-Content $Path -Value $ListFiles

    #Parse out the filename
    foreach($line in Get-Content ($Path)) {
        if($line.Contains($Name)){
            Write-Verbose -Message "$CurFnc | Returning Path : $line"
            return $line
        }
    }

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
    Will download file named "outbox/logs/FINPLAN_1000.txt" to $EPM_PATH_FILES_OUT\Testing123.txt

    .EXAMPLE
    EPM_Test-File -Name "Testing123.txt" -Path "C:\OutFiles"
    Will download file named "Testing123.txt" to "C:\OutFiles\Testing123.txt
#>
    param(
        #FileName to Download
        [String]$Name,
        #Path to move file To, by default will move to 03_Files\OUTBOUND
        [String]$Path = $EPM_PATH_FILES_OUT,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $ListFiles = (EPM_Execute-EPMATask -TaskName "Export List of Files" -TaskCommand "listfiles" -ReturnOut -NoLog)
    $DLCount = 0
    ForEach($line in $ListFiles) {
        if ($line.Trim().contains($Name)) {
            EPM_Execute-EPMATask -TaskName "Download $($line.Trim())" -TaskCommand "downloadFile" -TaskDetails ("`"$($line.Trim())`"") -TaskLevel $TaskLevel -ParentID $ParentID
            if ($LASTEXITCODE -eq 0) { Move-Item "$EPM_PATH_SCRIPTS\$Name*" -Destination "$Path" -Force -ErrorAction Ignore }
            $DLCount += 1
        }
    }

    if ($DLCount -eq 0) {EPM_Execute-EPMATask -TaskName "Download $Name" -TaskCommand "downloadFile" -TaskDetails ("`"$($Name.Trim())`"") -TaskLevel $TaskLevel -ParentID $ParentID}

}



function EPM_Get-AllFiles{
<#
    .SYNOPSIS
    Will download any files matching a pattern from the Inbox/Outbox

    .EXAMPLE
    EPM_Get-AllFile -Like "Testing123"
    Will download files named "Testing123.txt" and "Testing1234567.txt" to $EPM_PATH_FILES_OUT\Testing123.txt

#>
    param(
        #FileName to Download
        [String]$Name,
        #Path to move file To, by default will move to 03_Files\OUTBOUND
        [String]$Path = $EPM_PATH_FILES_OUT,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $ListFiles = EPM_Execute-EPMATask -TaskName "Export List of Files" -TaskCommand "listfiles" -ReturnOut -NoLog
    Set-Content -Path "$EPM_PATH_LOGS\listfiles.txt" -Value $ListFiles

    #Write-Host (Get-Content "$EPM_PATH_LOGS\listfiles.txt")

    ForEach ($line in (Get-Content "$EPM_PATH_LOGS\listfiles.txt")) {
        #Write-Host $line
        if ($line.Trim() -like "$Name") {
            Write-Host "Found : $line"
        }
    }


    #$WebPath = EPM_Test-File -Name $Name -TaskLevel $TaskLevel
    #if (-not $WebPath) { $WebPath = $Name}
    #EPM_Execute-EPMATask -TaskName "Download $Name" -TaskCommand "downloadFile" -TaskDetails ("`"$($WebPath.Trim())`"") -TaskLevel $TaskLevel -ParentID $ParentID
    #if ($LASTEXITCODE -eq 0) { Move-Item "$EPM_PATH_SCRIPTS\$Name*" -Destination "$Path" -Force -ErrorAction Ignore }
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
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Stops on Error
        [Switch]$StopOnError,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $CurFnc = EPM_Get-Function

    #Get File Name
    $FileName = $Path.Substring($Path.LastIndexOf('\')+1)
    $LastStatus = 0


    $TaskStartTime = Get-Date
    $TaskName = "Upload File : $FileName"
    EPM_Log-Task -TaskName $TaskName -TaskStage START -TaskLevel $TaskLevel -ParentID $ParentID -ForceAdd
    #Get Task ID
    $TaskID = ($global:EPM_TASK_LIST[-1].TASK_ID)

    if ($DataManagement) {
        EPM_Execute-EPMATask -TaskName "Delete Before Upload" -TaskCommand "deleteFile" -TaskDetails ("`"inbox\$FileName`"") -IgnoreError -TaskLevel ($TaskLevel+1) -ParentID $TaskID
        if ($StopOnError) {
            EPM_Execute-EPMATask -TaskName "Upload the File $FileName" -TaskCommand "uploadFile" -TaskDetails ("`"$Path`" inbox") -StopOnError -TaskLevel ($TaskLevel+1) -ParentID $TaskID
            
        } else { 
            EPM_Execute-EPMATask -TaskName "Upload the File $FileName" -TaskCommand "uploadFile" -TaskDetails ("`"$Path`" inbox") -TaskLevel ($TaskLevel+1) -ParentID $TaskID
            $LastStatus = $LASTEXITCODE
        }
        Write-Verbose -Message "$CurFnc | Uploaded to DM Inbox : $FileName"
    } else {
        EPM_Execute-EPMATask -TaskName "Delete Before Upload" -TaskCommand "deleteFile" -TaskDetails ("`"$FileName`"") -IgnoreError -TaskLevel ($TaskLevel+1) -ParentID $TaskID
        if ($StopOnError) {
            EPM_Execute-EPMATask -TaskName "Upload the File $FileName" -TaskCommand "uploadFile" -TaskDetails ("`"$Path`"") -StopOnError -TaskLevel ($TaskLevel+1) -ParentID $TaskID
        } else {
            EPM_Execute-EPMATask -TaskName "Upload the File $FileName" -TaskCommand "uploadFile" -TaskDetails ("`"$Path`"") -TaskLevel ($TaskLevel+1) -ParentID $TaskID
            $LastStatus = $LASTEXITCODE
        }
        Write-Verbose -Message "$CurFnc | Uploaded to General Inbox : $FileName"
    }

    if ( $LastStatus -eq 0) {
        EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus SUCCESS -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
    } else {
        EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
    }

}



function EPM_Move-FileToInstance{
<#
    .SYNOPSIS
    Will move a file from the provided source to the target environment, will handle logins / logouts etc

    .EXAMPLE
    EPM_Move-FileToInstance -SourceEnv "PROD" -TargetEnv "TEST" -FileName "20-04-05_PRD2TST_Testing" -IsSnapshot
    Will move a snapshot named "20-04-05_PRD2TST_Testing" from PROD to TEST
#>
    param(
        #[MANDATORY][VALUES = PROD,TEST] The source environment where the file/snapshot resides
        [parameter(Mandatory=$true)][ValidateSet("PROD","TEST")][String]$SourceEnv,
        #[MANDATORY][VALUES = PROD,TEST] The target environment where you want to move the file/snapshot to
        [parameter(Mandatory=$true)][ValidateSet("PROD","TEST")][String]$TargetEnv,
        #[MANDATORY] The name of the file/snapshot, if it is a file the path & file extension needs to be provided
        [parameter(Mandatory=$true)][String]$FileName,
        #The username to login to the source environment if $EPM_USER does not have access
        [String]$SourceUser = $EPM_USER,
        #The password file to login to the source environment if $EPM_USER does not have access
        [String]$SourcePassfile = $EPM_PASSFILE,
        #The username to login to the target environment if $EPM_USER does not have access
        [String]$TargetUser = $EPM_USER,
        #The password file to login to the target environment if $EPM_USER does not have access
        [String]$TargetPassfile = $EPM_PASSFILE,
        #True/False if this is a snapshot
        [switch]$IsSnapshot,
        #Assumes process is logged in, but allows the ability to login before moving snapshot
        [switch]$Login,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0
    )

    $CurFnc = EPM_Get-Function

    $TaskStartTime = Get-Date
    $TaskName = "Moving File $FileName from $SourceEnv to $TargetEnv"
    EPM_Log-Task -TaskName $TaskName -TaskStage START -TaskLevel $TaskLevel -ForceAdd
    $TaskID = ($global:EPM_TASK_LIST[-1].TASK_ID)

    # Set Proper URL
    if ($TargetEnv -eq "PROD") {
        $TargetURL = $EPM_URL_PROD
        $SourceURL = $EPM_URL_TEST
    } elseif ($TargetEnv -eq "TEST") {
        $TargetURL = $EPM_URL_TEST
        $SourceURL = $EPM_URL_PROD
    }

    #Check Current Environment & Login if necessary
    if ( ($EPM_ENV -ne $TargetEnv) -or ($Login) ) {

        if ( ($EPM_ENV -ne $TargetEnv) -and (-not ($Login)) ) {
            #Logout so we can login to $TargetEnv
            EPM_Execute-EPMATask -TaskName "Logout of $EPM_ENV for File Move" -TaskCommand "logout" -TaskLevel ($TaskLevel+1) -ParentID $TaskID
        }

        EPM_Execute-EPMATask -TaskName "Log into $TargetEnv for File Move" -TaskCommand "login" -TaskDetails "$TargetUser $TargetPassfile $TargetURL $EPM_DOMAIN" -StopOnError -TaskLevel ($TaskLevel+1) -ParentID $TaskID
    }

    #Move File
    EPM_Execute-EPMATask -TaskName "Delete $FileName from $TargetEnv" -TaskCommand "deleteFile" -TaskDetails "$FileName" -IgnoreError -TaskLevel ($TaskLevel+1) -ParentID $TaskID
    if ($IsSnapshot) {
        EPM_Execute-EPMATask -TaskName "Moving Snapshot $FileName From $SourceEnv To $TargetEnv" -TaskCommand "copySnapshotFromInstance" -TaskDetails "$FileName $SourceUser $SourcePassfile $SourceURL $EPM_DOMAIN" -StopOnError -TaskLevel ($TaskLevel+1) -ParentID $TaskID
    } else {
        EPM_Execute-EPMATask -TaskName "Moving File $FileName From $SourceEnv To $TargetEnv" -TaskCommand "copyFileFromInstance" -TaskDetails "$FileName $SourceUser $SourcePassfile $SourceURL $EPM_DOMAIN $FileName" -StopOnError -TaskLevel ($TaskLevel+1) -ParentID $TaskID
    }

    # Return login status to what it was before
    if ( ($EPM_ENV -eq $SourceEnv) -or ($Login) ) {

        if ( ($EPM_ENV -eq $SourceEnv) -and (-not ($Login)) ) {
            EPM_Execute-EPMATask -TaskName "Logout of $TargetEnv to restore access to $SourceEnv" -TaskCommand "logout" -TaskLevel ($TaskLevel+1) -ParentID $TaskID
            EPM_Execute-EPMATask -TaskName "Restore Access to $SourceEnv" -TaskCommand "login" -TaskDetails "$SourceUser $SourcePassfile $SourceURL $EPM_DOMAIN" -StopOnError -TaskLevel ($TaskLevel+1) -ParentID $TaskID
        }
        if ($Login) {
            EPM_Execute-EPMATask -TaskName "Logout of $TargetEnv to restore not logged in status" -TaskCommand "logout" -TaskLevel ($TaskLevel+1) -ParentID $TaskID
        }
    }

    EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus SUCCESS -StartTime $TaskStartTime -UpdateTask $TaskID
}


function EPM_Execute-LoadRule{
<#
    .SYNOPSIS
    Will execute a Data Mgmt load rule and download the log after completion and parse for kickouts

    .EXAMPLE
    EPM_Execute-LoadRule -LoadRule "LR_OP_TEST_NUM" -StartPeriod "Oct-15" -EndPeriod "Mar-16" -Path "$EPM_PATH_FILES_IN\TEST_DM_NUMERIC.txt"
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
        #Will stop process on error
        [Switch]$StopOnError,
        #Import Mode to Use
        #   REPLACE - Will truncate the DM Import table and then import the records in the provided file
        #   APPEND - Will add the records in the file to the records that are currently in the DM Import Table
        #   RECALCULATE - Skip the import, but re-run the mappings
        #   NONE - Skip the import & do NOT re-run the mappings
        [ValidateSet("REPLACE","APPEND","RECALCULATE","NONE")][String]$ImportMode = "REPLACE",
        #Export Mode to Use
        #   STORE_DATA - Will load the data into Essbase will overwrite intersections, but will not clear none-specified intersections
        #   ADD_DATA - Will load the data into essbase but will add the data to existing intersections
        #   SUBTRACT_DATA - Will load the data into essbase but will subtract the data from existing intersections
        #   REPLACE_DATA - Will clear the POV (Scenario, Version, Year, Period, Entity) before importing the data
        #   NONE - Skip the export
        [ValidateSet("STORE_DATA","ADD_DATA","SUBTRACT_DATA","REPLACE_DATA","NONE")][String]$ExportMode = "STORE_DATA",
        #Task Level for Logging Purposes
        [Int]$TaskLevel = 0
    )

    $TaskStartTime = Get-Date
    if ($Path) {
        $TaskName = "Executing Load of $FileName"
        $FileName = "$($Path.Substring($Path.LastIndexOf('\')+1))"
        $LoadTask = "Loading $FileName via $LoadRule for $StartPeriod to $EndPeriod"
    } else {
        $TaskName = "Executing Load of $LoadRule"
        $FileName = ""
        $LoadTask = "Loading Data via $LoadRule for $StartPeriod to $EndPeriod"
    }

    EPM_Log-Task -TaskName $TaskName -TaskStage START -TaskLevel $TaskLevel -ForceAdd
    $TaskID = ($global:EPM_TASK_LIST[-1].TASK_ID)

    if ($Path) {
        EPM_Upload-File -Path "$Path" -DataManagement -TaskLevel ($TaskLevel+1) -ParentID $TaskID
        if ($LASTEXITCODE -ne 0) {
            if ($StopOnError) {
                EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -UpdateTask $TaskID -StopOnError 
            } else {
                EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -UpdateTask $TaskID
                Return 1 | Out-Null
            }
        }
    }
    
    EPM_Execute-EPMATask -TaskName $LoadTask -TaskCommand "runDataRule" -TaskDetails ("$LoadRule $StartPeriod $EndPeriod $ImportMode $ExportMode $FileName") -TaskLevel ($TaskLevel+1) -ParentID $TaskID 
    $LastStatus = $LASTEXITCODE
    $LoadTaskID = ($global:EPM_TASK_LIST[-1].TASK_ID)

    #$DMError = 0
    if ($LastStatus -ne 0) {
        #We had an error or kickouts, determine which.
        #Parse Log
        $ErrorLog = Get-ChildItem "$EPM_PATH_SCRIPTS" -Filter runDataRule*.log | Sort-Object LastWriteTime | Select-Object -Last 1
        $throwError = $false
        if ($ErrorLog) {
            $DMLog = [regex]::Match((Get-Content $ErrorLog.FullName),"`"logFileName`":`"([a-zA-Z\/\.\:\-_0-9]+)`"").Groups[1].Value
            if ($DMLog){
                #Kickout Log Found
                $DMLog = $DMLog.Substring($DMLog.LastIndexOf('/')+1)
                $KickoutLog = $DMLog.Replace(".log",".out")
                EPM_Get-File -Name $DMLog -Path $EPM_PATH_LOGS -TaskLevel ($TaskLevel+1) -ParentID $TaskID 
                EPM_Get-File -Name $KickoutLog -Path $EPM_PATH_LOGS -TaskLevel ($TaskLevel+1) -ParentID $TaskID 

                $LoadID = $DMLog.Substring($DMLog.LastIndexOf('/')+1).Split("_")[1].Replace(".log","")
                $RuleName = [regex]::Match((Get-Content "$EPM_PATH_LOGS\$DMLog"),"Rule Name    : (.*? )").Groups[1].Value
                $LoadFileName = [regex]::Match((Get-Content "$EPM_PATH_LOGS\$DMLog"),"File Name.*: (.*?txt)").Groups[1].Value

                foreach ($line in Get-Content "$EPM_PATH_LOGS\$KickoutLog") {
                    if($line.Contains("Error: 3303")) {
                        $arrLine = $line.split("|")
                        "$LoadID#$RuleName#$LoadFileName#$($arrLine[2].trim())#$($arrLine[3].trim())" | Add-Content -Path $EPM_LOG_KICKOUTS
                    } elseif ($line.Contains("The member ")) {
                        $member = [regex]::Match(($line),"(The member )(.*)( does not exist)").Groups[2].Value
                        "$LoadID#$RuleName#$LoadFileName#$member#Not Available" | Add-Content -Path $EPM_LOG_KICKOUTS
                    }
                }
                
                EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus WARNING -StartTime $TaskStartTime -UpdateTask $TaskID
                EPM_Log-Task -TaskName $LoadTask -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus WARNING -StartTime $TaskStartTime -UpdateTask $LoadTaskID -OverrideError "Review Kickouts for Load ID : $LoadID" -NoLog
            } else {
                $throwError = $true
            }
        } else {
            $throwError = $true
        }

        if ($throwError) {
            if ($StopOnError) {
                EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -UpdateTask $TaskID -StopOnError
            } else {
                EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -UpdateTask $TaskID
            }
        }
    } else {
        EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus SUCCESS -StartTime $TaskStartTime -UpdateTask $TaskID
    }
    
}


# -----------------------------------------------------------------------------------
#   EPM UTILITIES - MAINTENANCE
# -----------------------------------------------------------------------------------



function EPM_Backup-Application{
<#
    .SYNOPSIS
    Will download the Artifact Snapshot and move it to a specified file path. If no path is provided will default to 04_BACKUPS\<CurrentDateTime>-<Environemnt>-<ApplicationName>-BACKUP.zip
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
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $CurFnc = EPM_Get-Function

    $TaskStartTime = Get-Date
    $TaskName = "Backup $EPM_FINPLAN in $EPM_ENV"
    $ArtifactSnapshot = '"Artifact Snapshot"'
    EPM_Log-Task -TaskName $TaskName -TaskStage START -TaskLevel $TaskLevel -ForceAdd -ParentID $ParentID
    $TaskID = ($global:EPM_TASK_LIST[-1].TASK_ID)

    if ($New) {
        EPM_Execute-EPMATask -TaskName "Re-Export Snapshot" -TaskCommand ("exportSnapshot $ArtifactSnapshot") -TaskLevel ($TaskLevel+1) -ParentID $TaskID
        if ($LASTEXITCODE -ne 0) {
            EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
            return ""
        }
    }

    EPM_Execute-EPMATask -TaskName "Download Snapshot" -TaskCommand ("downloadFile $ArtifactSnapshot") -TaskLevel ($TaskLevel+1) -ParentID $TaskID
    if ($LASTEXITCODE -ne 0) {
        EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
        return ""
    } else {

        Move-Item -Path "$EPM_PATH_SCRIPTS\Artifact Snapshot.zip" -Destination "$Path" | EPM_Log-Item
        "Backup Moved to $Path" | EPM_Log-Item

        #Apply Backup Retention Policy set in 01_SCRIPTS\_EPM_Config.ps1
        if ($EPM_BACKUPS_RETAIN_POLICY -eq "NUM"){
            Get-ChildItem "$EPM_PATH_BACKUPS\*.zip" -Recurse | Sort-Object CreationTime -Descending | Select-Object -Skip $EPM_BACKUPS_RETAIN_NUM | Remove-Item -Force
        } elseif ($EPM_BACKUPS_RETAIN_POLICY -eq "DAYS") {
            Get-ChildItem "$EPM_PATH_BACKUPS\*.zip" -Recurse | Where-Object {$_.LastWriteTime -lt  (Get-Date).AddDays(-$EPM_BACKUPS_RETAIN_NUM)} | Remove-Item -Force
        }

        EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus SUCCESS -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID 
    }

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
        [parameter(Mandatory=$true)][ValidateSet("MergeSlicesKeepZero","MergeSlicesRemoveZero")][String]$Task,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $CurFnc = EPM_Get-Function

    if ($Task -eq "MergeSlicesKeepZero" ) {
        EPM_Execute-EPMATask -TaskName "Merge Slices - Keep Zero" -TaskCommand "mergeDataSlices" -TaskDetails "`"$Cube`" keepZeroCells=false" -TaskLevel $TaskLevel -ParentID $ParentID
    } elseif ($Task -eq "MergeSlicesRemoveZero") {
        EPM_Execute-EPMATask -TaskName "Merge Slices - Remove Zero" -TaskCommand "mergeDataSlices" -TaskDetails "`"$Cube`" keepZeroCells=true" -TaskLevel $TaskLevel -ParentID $ParentID
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
        [parameter(Mandatory=$true)][ValidateSet("RestructureCube")][String]$Task,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $CurFnc = EPM_Get-Function

    if ($Task -eq "RestructureCube") {
        EPM_Execute-EPMATask -TaskName "Restructure Cube" -TaskCommand "restructureCube" -TaskDetails "`"$Cube`"" -TaskLevel $TaskLevel -ParentID $ParentID
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
        [parameter(Mandatory=$true)][ValidateSet("ModeAdmin","ModeUser")][String]$Task,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $CurFnc = EPM_Get-Function

    if ($Task -eq "ModeAdmin") {
        EPM_Execute-EPMATask -TaskName "Set App Mode : Admin" -TaskCommand "applicationAdminMode" -TaskDetails "$true" -TaskLevel $TaskLevel -ParentID $ParentID
    } elseif ($Task -eq "ModeUser") {
        EPM_Execute-EPMATask -TaskName "Set App Mode : User" -TaskCommand "applicationAdminMode" -TaskDetails "$false" -TaskLevel $TaskLevel -ParentID $ParentID
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
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $CurFnc = EPM_Get-Function

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
    Write-Verbose -Message "$CurFnc | Export to : $FileNameCSV"

    #Process Security
    $TaskStartTime = Get-Date
    $TaskName = "Exporting Security to $PathWithFileName"
    EPM_Log-Task -TaskName $TaskName -TaskStage START -TaskLevel $TaskLevel -ForceAdd -ParentID $ParentID
    $TaskID = ($global:EPM_TASK_LIST[-1].TASK_ID)

    EPM_Execute-EPMATask -TaskName "Generate Security Export in Outbox : $FileNameCSV" -TaskCommand "exportAppSecurity" -TaskDetails "$FileNameCSV" -TaskLevel ($TaskLevel+1) -ParentID $TaskID
    if ($LASTEXITCODE -ne 0) {
        #Export Failed
        EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
    } else {
        #Export Success
        EPM_Get-File -Name $FileNameCSV -Path $PathWithFileName -TaskLevel ($TaskLevel+1) -ParentID $TaskID
        if ($LASTEXITCODE -ne 0) {
            #Download Failed
            EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
        } else {
            #Download Success
            EPM_Execute-EPMATask -TaskName "Delete Sec Export From Web $FileNameCSV" -TaskCommand "deleteFile" -TaskDetails "$FileNameCSV" -IgnoreError -TaskLevel ($TaskLevel+1) -NoLog -ParentID $TaskID
            EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus SUCCESS -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
        }
    }

}



function EPM_Import-Security{
<#
    .SYNOPSIS
    Will Export Security and store with timestamp in $EPM_PATH_FILES_OUT\<TimeStamp>_OriginalSecurity.csv
    Will Import application security

    .EXAMPLE
    EPM_Import-Security -ImportFile "$EPM_PATH_FILES_IN\UpdatedSecurity.csv"
    Will export current App Security to $EPM_PATH_FILES_OUT\<TimeStamp>_OriginalSecurity.csv
    Will import $EPM_PATH_FILES_IN\UpdatedSecurity.csv and write errors to $EPM_PATH_LOGS\SecurityErrors.log

    .EXAMPLE
    EPM_Import-Security -ImportFile "$EPM_PATH_FILES_IN\UpdatedSecurity.csv" -ErrorFile "C:\Testing\ErrorTesting.log" -ClearAll
    Will export current App Security to $EPM_PATH_FILES_OUT\<TimeStamp>_OriginalSecurity.csv
    Will remove all current security and import $EPM_PATH_FILES_IN\UpdatedSecurity.csv and write errors to C:\Testing\ErrorTesting.log
#>
    param(
        #[MANDATORY] File Name & Path to Security File to Import
        [parameter(Mandatory=$true)][String]$ImportFile,
        #File Name & Path of Error File, by default writes to $EPM_PATH_LOGS\SecurityErrors.log
        [String]$ErrorFileName = "SecurityErrors.csv",
        #Specifies whether this is a sub-task or not
        [Switch]$ClearAll = $false,
        #Specifies the Level of Task Being Run
        [Int]$TaskLevel = 0,
        #Parent Task ID
        [Int]$ParentID = 0
    )

    $CurFnc = EPM_Get-Function

    $OrigSecurity = "$(EPM_Get-TimeStamp -StampType FILE)_OriginalSecurity.csv"
    $ImportFileName = $ImportFile.Substring($ImportFile.LastIndexOf('\')+1)

    $TaskStartTime = Get-Date
    $TaskName = "Import Security from $ImportFile"
    EPM_Log-Task -TaskName $TaskName -TaskStage START -TaskLevel $TaskLevel -ForceAdd -ParentID $ParentID
    $TaskID = ($global:EPM_TASK_LIST[-1].TASK_ID)

    #Export Original Security
    EPM_Export-Security -FileName "$OrigSecurity" -TaskLevel ($TaskLevel+1) -ParentID $TaskID
    if ($LASTEXITCODE -ne 0) {
        #Export Failed
        EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
    } else {
        #Export Success
        $ImpSecTaskName = "Importing New Security"
        EPM_Log-Task -TaskName $ImpSecTaskName -TaskStage START -TaskLevel ($TaskLevel+1) -ForceAdd -ParentID $TaskID
        $ImpSecTaskID = ($global:EPM_TASK_LIST[-1].TASK_ID)
        $ImportTime = Get-Date
        EPM_Upload-File -Path "$ImportFile" -TaskLevel ($TaskLevel+2) -ParentID $ImpSecTaskID
        if ($LASTEXITCODE -ne 0) {
            #Upload Failed
            EPM_Log-Task -TaskName $ImpSecTaskName -TaskStage FINISH -TaskLevel ($TaskLevel+1) -TaskStatus ERROR -StartTime $ImportTime -ParentID $TaskID -UpdateTask $ImpSecTaskID
            EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
        } else {
            #Upload Success
            EPM_Execute-EPMATask -TaskName "Import Security" -TaskCommand "importAppSecurity" -TaskDetails "`"$ImportFileName`" `"$ErrorFileName`" clearall=$ClearAll" -TaskLevel ($TaskLevel+2) -ParentID $ImpSecTaskID
            if ($LASTEXITCODE -ne 0) {
                #Import Failed
                EPM_Get-File -Name "$ErrorFileName" -Path "$EPM_LOG_SECURITY" -TaskLevel ($TaskLevel+2) -ParentID $ImpSecTaskID
                EPM_Execute-EPMATask -TaskName "Delete Error File from Web $ErrorFileName" -TaskCommand "deleteFile" -TaskDetails "`"$ErrorFileName`"" -IgnoreError -TaskLevel ($TaskLevel+2) -NoLog -ParentID $ImpSecTaskID
                EPM_Log-Task -TaskName $ImpSecTaskName -TaskStage FINISH -TaskLevel ($TaskLevel+1) -TaskStatus ERROR -StartTime $ImportTime -ParentID $TaskID -UpdateTask $ImpSecTaskID
                EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus ERROR -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
                
            } else {
                #Import Success
                EPM_Log-Task -TaskName $ImpSecTaskName -TaskStage FINISH -TaskLevel ($TaskLevel+1) -TaskStatus SUCCESS -StartTime $ImportTime -ParentID $TaskID -UpdateTask $ImpSecTaskID
                EPM_Log-Task -TaskName $TaskName -TaskStage FINISH -TaskLevel $TaskLevel -TaskStatus SUCCESS -StartTime $TaskStartTime -ParentID $ParentID -UpdateTask $TaskID
            }
        }
    }
}





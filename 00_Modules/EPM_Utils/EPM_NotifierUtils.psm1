# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : House notification functions to use during EPM Automation (Oracle EPBCS)
# ===================================================================================

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
        $ErrCount = $EPM_TASKLIST.countTasks("status","ERROR")
        $WarnCount = $EPM_TASKLIST.countTasks("status","WARNING")
       
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
        $TitleStyle="border: 3px solid black; border-bottom: 0px; text-align: center; padding: 4px 3px; font-weight: bold; color:white; background: $TitleBgColor;"
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
        $EmailBody += "`n<caption style='$TitleStyle'>PROCESS SUMMARY</caption>"
        $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>ENVIRONMENT</td>" +
                            "<td style='$CellStyle text-align: center;'>$EPM_ENV</td></tr>"
        $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>PROCESS</td>" + 
                            "<td style='$CellStyle text-align: center;'>$EPM_PROCESS</td></tr>"
        $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>STATUS</td>" +
                            "<td style='$CellStyle text-align: center; $ProcessStyle'>$ProcessStatus</td></tr>"
        $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>START TIME</td>" + 
                            "<td style='$CellStyle text-align: center;'>$($EPM_PROCESS_START.ToString("MM/dd/yy hh\:mm\:ss tt"))</td></tr>"
        $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>END TIME</td>" + 
                            "<td style='$CellStyle text-align: center;'>$($EndTime.ToString("MM/dd/yy hh\:mm\:ss tt"))</td></tr>"
        $EmailBody += "`n   <tr><td style='$HeaderStyle text-align: left; border-bottom: 0px;'>ELAPSED TIME</td>" + 
                            "<td style='$CellStyle text-align: center;'>$($ElapsedTime.ToString("hh\:mm\:ss"))</td></tr>"
        #Close Processes Table
        $EmailBody += "`n</table>"
        $EmailBody += "`n<br><br>"
    
        # ----------------------------------
        # Build Variable Table
        # ----------------------------------
        $EmailBody += $TableStyle
        $EmailBody += "`n<caption style='$TitleStyle'>PROCESS VARIABLES</caption>"
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
        # Build Tasks & Error Table Table
        # ----------------------------------
        # -- Build Tasks Table
        $TaskTable = $TableStyle
        $TaskTable += "`n<caption style='$TitleStyle'>TASK SUMMARY</caption>"
        #Build Task Table Header
        $TaskTable += "`n   <tr>"
        ForEach($hdr in @("Task ID","Status","Task","Start Time","End Time","Elapsed Time","Elapsed Time %")){
            $TaskTable += "`n    <th style='$HeaderStyle text-align: center;'>$hdr</th>"
        }
        $TaskTable += "`n   </tr>"

        # -- Build Error Table
        $ErrorTable = $TableStyle
        $ErrorTable += "`n<caption style='$TitleStyle'>ERROR SUMMARY</caption>"
        #Build Error Table Header
        $ErrorTable += "`n   <tr>"
        ForEach($hdr in @("Task ID","Status","Task Command","Command Details","Error Message","Call Stack")){
            $ErrorTable += "`n    <th style='$HeaderStyle text-align: center;'>$hdr</th>"
        }
        $ErrorTable += "`n   </tr>"


        #Add Tasks to Table as Rows
        ForEach ($task in $EPM_TASKLIST.Tasks) {

            if ( !$task.hideTask ) {
                Switch($task.status){
                    SUCCESS {$StatusColor = $SuccessColor}
                    ERROR {$StatusColor = $ErrorColor}
                    Default {$StatusColor = $WarningColor}
                }

                if ( $task.level -eq 0 ){
                    $indentColor = ""
                    $PoT = ($task.elapsedTime.TotalMilliseconds / $ElapsedTime.TotalMilliseconds).ToString("P")
                } else {
                    $indentColor = "color: grey;"
                    $PoT = ""
                }

                $cstack = ""
                ForEach($item in $task.callstack){
                    $cstack += "$item<br>"
                }
                

                $TaskTable += "`n   <tr>`n      <td style='$CellStyle text-align: center;'>$($task.id)</td>"
                $TaskTable += "`n      <td style='$CellStyle background: $StatusColor; text-align: center;'>$($task.status)</td>"
                $TaskTable += "`n      <td style='$CellStyle font-weight:bold; $indentColor'>$("&nbsp;&nbsp;" * $task.level)$($task.name)</td>"
                $TaskTable += "`n      <td style='$CellStyle text-align: center;'>$($task.startTime.ToString("hh\:mm\:ss tt"))</td>"
                $TaskTable += "`n      <td style='$CellStyle text-align: center;'>$($task.endTime.ToString("hh\:mm\:ss tt"))</td>"
                $TaskTable += "`n      <td style='$CellStyle text-align: center;'>$("{0:hh\:mm\:ss}" -f $task.elapsedTime)</td>"
                $TaskTable += "`n      <td style='$CellStyle text-align: center;'>$PoT</td>"
                $TaskTable += "`n   </tr>"

                if ( (@("ERROR","WARNING").contains($task.status)) -and ($task.errorMsg) ){
                    $ErrorTable += "`n   <tr>`n      <td style='$CellStyle text-align: center;'>$($task.id)</td>"
                    $ErrorTable += "`n      <td style='$CellStyle text-align: center; background: $StatusColor;'>$($task.status)</td>"
                    $ErrorTable += "`n      <td style='$CellStyle text-align: center;'>$($task.command)</td>"
                    $ErrorTable += "`n      <td style='$CellStyle text-align: left;'>$($task.details)</td>"
                    $ErrorTable += "`n      <td style='$CellStyle text-align: left;'>$($task.errorMsg)</td>"
                    $ErrorTable += "`n      <td style='$CellStyle text-align: left;'>$cstack</td>"
                    $ErrorTable += "`n   </tr>"
                }
            }

        }
        #Close Task Table
        $TaskTable += "`n</table>"
        $TaskTable += "`n<br><br>"
        #Close Error Table
        $ErrorTable += "`n</table>"
        $ErrorTable += "`n<br><br>"

        # Add Tables to Body
        $EmailBody += $TaskTable
        if (($ErrCount + $WarnCount) -ge 1){
            $EmailBody += $ErrorTable
        }
        
    
        # ----------------------------------
        # Build Kickouts Table
        # ----------------------------------
        if (Test-Path $EPM_LOG_KICKOUTS) {
    
            $Kickouts = Get-Content -Path $EPM_LOG_KICKOUTS
        
            #Build Kickouts Table
            $EmailBody += $TableStyle
            $EmailBody += "`n<caption style='$TitleStyle'>KICKOUT SUMMARY ($($Kickouts.Length) Total Kickouts)</caption>"
            #Build Kickouts Table Header
            $EmailBody += @"
            `n   <tr>
                  <th style='$HeaderStyle text-align: center;'>Load ID</th>
                  <th style='$HeaderStyle text-align: center;'>Load Rule</th>
                  <th style='$HeaderStyle text-align: center;'>Load File</th>
                  <th style='$HeaderStyle text-align: center;'>Kickout Member</th>
                  <th style='$HeaderStyle text-align: left;'>Record Count</th>
               </tr>
"@
            #Add Kickouts to Table as Rows
            $uniqueKickouts = @()
            $uniqueMembers = @()
            ForEach ($line in $Kickouts) {
                $arr = $line.split("#")
                $uniqueKickouts += "$($arr[0])#$($arr[1])#$($arr[2])#$($arr[3])"
                $uniqueMembers += $arr[3]
            }
        
            $uniqueKickouts = ($uniqueKickouts | Sort-Object | Get-Unique)
            ForEach ($line in $uniqueKickouts) {
                $arr = $line.split("#")
                $EmailBody += "`n   <tr>`n      <td style='$CellStyle text-align: center;'>$($arr[0])</td>" #Load Id
                $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$($arr[1].Trim())</td>" #Load Rule
                $EmailBody += "`n      <td style='$CellStyle text-align: left;'>$($arr[2].Trim())</td>" #Load File
                $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$($arr[3].Trim())</td>" #Kickout Member
                $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$(($Kickouts | Select-String -pattern "$line").length)</td>" #Kickout Count
                $EmailBody += "`n   </tr>"
            }
            #Close Kickouts Table
            $EmailBody += "`n</table>"
            $EmailBody += "`n<br><br>"
        
            # Unique Member Kickout Table
            $EmailBody += $TableStyle
            $EmailBody += "`n<caption style='$TitleStyle'>UNIQUE MEMBER KICKOUTS</caption>"
            #Build Kickouts Table Header
            $EmailBody += @"
            `n   <tr>
                  <th style='$HeaderStyle text-align: center;'>Kickout Member</th>
                  <th style='$HeaderStyle text-align: left;'>Record Count</th>
               </tr>
"@
        
            $uniqueMembers = ($uniqueMembers | Sort-Object | Get-Unique)
            ForEach ($member in $uniqueMembers) {
                $EmailBody += "`n   <tr>`n      <td style='$CellStyle text-align: center;'>$member</td>" #Member
                $EmailBody += "`n      <td style='$CellStyle text-align: center;'>$(($Kickouts | Select-String -pattern "$member").length)</td>" #Count
                $EmailBody += "`n   </tr>"
            }
            #Close Kickouts Table
            $EmailBody += "`n</table>"
        
       
        }
    
    
        $EmailBody += "`n</body></html>"
        $EmailBody = $EmailBody.replace("$EPM_PATH_AUTO","")
    
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
            Attachment = (Get-ChildItem ("$EPM_PATH_LOGS\LOG_*.log","$EPM_PATH_LOGS\*.csv")).FullName
            Credential = $EPM_EMAIL_CREDENTIALS
            Port = $EPM_EMAIL_PORT
            SmtpServer = $EPM_EMAIL_SERVER
            Priority = $Priority
        }
    
        if ($EPM_EMAIL_CC.Count -gt 0) {$param.Add("CC",$EPM_EMAIL_CC)}
        if ($Notify -and ($EPM_EMAIL_CREDENTIALS -ne "") ) {Send-MailMessage @param -UseSsl -BodyAsHtml}
        Set-Content -Path "$EPM_PATH_SCRIPTS\_EmailBody.htm" -Value $EmailBody
    }
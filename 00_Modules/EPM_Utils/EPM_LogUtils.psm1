# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 09-02-2020
#   Purpose : House logging functions to use during EPM Automation (Oracle EPM)
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
            return "{0:MM/dd/yyyy} {0:hh:mm:ss tt}" -f (Get-Date)
        } elseif ($StampType -eq 'FILE') {
            return "{0:yyyy-MM-dd}_{0:HH-mm-ss}" -f (Get-Date)
        } else {
            return "[{0:yyyy-MM-dd} {0:hh:mm:ss tt}]" -f (Get-Date)
        }
    }
    

filter EPM_Log-Item{
    <#
        .SYNOPSIS
        Meant for messages to be piped in and will write to host / EPM_LOG_FULL with timestamp
    
        .EXAMPLE
        "Testing 123" | EPM_Log-Item
        Will output to the LOG_FULL : [04/06/20 08:07:53] : Testing123
    
        .EXAMPLE
        "Download Log" | EPM_Log-Item -IncludeSeparator -
        Will output to the LOG_FULL :     [04/06/20 08:07:56] : Download Log
        Will include the Task separator after it
    
    #>
        Param(
            #Control the Log it is Written To
            [String]$LogFile = $EPM_LOG_FULL,
            #The type of log item we are writing
            [ValidateSet("VERBOSE","INFO","WARN","ERROR","IGNORE")][String]$LogType = "INFO",
            #Whether or note the message is written to host
            [switch]$WriteHost,
            #Control the Color of the text written to host
            [ValidateSet("Cyan","Gray","Yellow","Red","Green")][String]$HostColor,
            #Hides TimeStamp & LogType
            [switch]$Clean,
            #Allows for a separator to be placed after the item is logged
            [switch]$IncludeSeparator
        )

        $LogLevel = switch ($LogType)
            {
                VERBOSE {0; $Color = "Cyan"}
                INFO    {1; $Color = "Gray"}
                WARN    {2; $Color = "Yellow"}
                IGNORE  {2; $Color = "Yellow"}
                ERROR   {3; $Color = "Red"}
            }
        $GlobalLogLevel = switch ($EPM_LOG_LEVEL)
            {
                VERBOSE {0}
                INFO    {1}
                WARN    {2}
                ERROR    {3}
            }
        
        if (-not $HostColor){$HostColor = $Color}

        # $_ in this case is whatever is piped to this
        if (-not $Clean){
            $EPMAutoError = [regex]::Match(($_),"^EPM.*-[0-9]*:.*$").Value
            if ($EPMAutoError){
                $LogType = "ERROR"
            }
            $LogType2 = "[$LogType]".PadRight(8,' ')
            $message =  "$(EPM_Get-TimeStamp) $LogType2 : $_" 
        } else {
            $message =  "$_" 
        }
        
        if ($LogLevel -ge $GlobalLogLevel) {
            $message | Add-Content -Path $LogFile
            if ($IncludeSeparator) {"$EPM_TASK_SEPARATOR" | Add-Content -Path $LogFile}
            if ($WriteHost) {Write-Host $message -ForegroundColor $HostColor}
        }
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
        return ("{0:hh\:mm\:ss}" -f (New-TimeSpan -Start $StartTime -End (Get-Date)))
    }
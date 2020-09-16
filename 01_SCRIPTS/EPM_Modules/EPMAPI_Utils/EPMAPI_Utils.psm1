# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : House common functions to use during EPM Automation with REST API (Oracle EPBCS)
# ===================================================================================

<#
    TODO Utilities : 
        - API For Data Loads
#>

# -----------------------------------------------------------------------------------
#   EPM API UTILITIES
# -----------------------------------------------------------------------------------



function EPMAPI_Create-AuthFile {
<#
    .SYNOPSIS
    Creates an Authentication file in $EPM_PATH_SCRIPTS\APIEncodedAuthCR.txt to be used for API

    .EXAMPLE
    EPAPI_Create-AuthFile -Pass "Testing123"
#>
    Param(
        #Domain to use, by default it is $EPM_DOMAIN
        [String]$Domain = $EPM_DOMAIN,
        #Username to use, by default it is $EPM_USER
        [String]$User = $EPM_USER,
        #Password to encrypt
        [parameter(Mandatory=$true)][String]$Pass
    )
    
    $tempEncName = [System.Text.Encoding]::UTF8.GetBytes("$domain" + '.' + $user + ':' + "$pass")
    $tempEncAuth = [System.Convert]::ToBase64String($tempEncName)
    Set-Content "$EPMAPI_PASSFILE" -Value $tempEncAuth -Force

    #check if we need EPM Automate Encryption
    if (-not (Test-Path -Path $EPM_PASSFILE) ) {
        epmautomate encrypt "$Pass" "EPMAutoEncKey$(Get-Random -Maximum 1000)" $EPM_PASSFILE
    }
}



function EPMAPI_Create-APIVerURIFile{
<#
    .SYNOPSIS
    Grabs the API 

    .EXAMPLE
    EPAPI_Create-AuthFile -Pass "Testing123"
#>

    # Get Planning Version
    $RestResponse = Invoke-RestMethod -Uri "$EPM_URL/HyperionPlanning/rest/" -Method GET -Headers $EPMAPI_HEADER
    $RestResponse.items | ForEach-Object { if ($_.isLatest -eq $true) { $TempVer = $_.version } }
    $TempOut = "PLAN_VERSION=$TempVer`n"

    # Get Migration Version
    $RestResponse = Invoke-RestMethod -Uri "$EPM_URL/interop/rest/" -Method GET -Headers $EPMAPI_HEADER
    $RestResponse.items | ForEach-Object { if ($_.Latest -eq $true) { $TempVer = $_.version } }
    $TempOut += "MIG_VERSION=$TempVer`n"

    # Get Data Management Version
    $RestResponse = Invoke-RestMethod -Uri "$EPM_URL/aif/rest/" -Method GET -Headers $EPMAPI_HEADER
    $RestResponse | ForEach-Object { if ($_.isLatest -eq $true) { $TempVer = $_.version } }
    $TempOut += "DMG_VERSION=$TempVer`n"

    Set-Content "$EPM_PATH_SCRIPTS\EPMAPI_Info.txt" -Value $tempOut
}



function EPMAPI_Execute-Request{
<#
    .SYNOPSIS
    Executes an API Request that is passed in

    .EXAMPLE
    EPAPI_Create-AuthFile -Pass "Testing123"
#>
    param(
        #[MANDATORY] Name of the Task for Logging Purposes
        [parameter(Mandatory=$true)][String]$TaskName,
        #Method of the API Request
        #   POST - Generally when telling the API to run / execute something
        #   GET - Generally retrieving information
        [parameter(Mandatory=$true)][ValidateSet("POST","GET")][String]$Method,
        #URI of the request, where the API resource can be found
        [parameter(Mandatory=$true)][String]$URI,
        #ContentType of the rreturned response
        [ValidateSet("application/json")][String]$ContentType = "application/json",
        #Authentication Header
        $Header = $EPMAPI_HEADER,
        #Body of the request if there are parameters to the request
        $Body,
        #Task Level for Logging Purposes
        [Int]$TaskLevel = 0,
        #Turns off the Logging
        [Switch]$NoLog,
        #Exit on error
        [Switch]$StopOnError,
        #Will not log or act on any errors
        [Switch]$IgnoreErrorHandling

    )

    if (-not (Test-Path $EPM_PROCESS_RUNNING_FLAG) ) {
        $ErrorMsg = "Please run EPM_Start-Process to login"
        Write-Host "$ErrorMsg" -ForegroundColor Yellow
        "$ErrorMsg" | EPM_Log-Item -LogType ERROR
        break
    }

    #Logging
    $TaskStartTime = Get-Date
    $TaskCommand = "$Method $($URI.Replace("$EPM_URL"," ")) $Body"
    if (-not ($NoLog)) {"$TaskName" | EPM_Log-Item -LogType TASK -LogStatus START -TaskLevel $TaskLevel}
    if (-not ($NoLog)) {"$TaskCommand" | EPM_Log-Item -LogType TASKCOMMAND}

    #Execute Request
    <#if ($Method -eq "POST") {
        $RestResponse = Invoke-RestMethod -Uri $URI -Body $Body -Headers $Header -Method $Method -ContentType "$ContentType"
        $GetURI = $RestResponse.links[0].href
        while ($RestResponse.Status -eq -1) {
            $RestResponse = Invoke-RestMethod -Uri $GetURI -Headers $EPMAPI_HEADER -Method GET -ContentType "$ContentType"
            if ($RestResponse.Status -eq -1) {
                Write-Verbose -Message "Sleeping 5 seconds..."
                Start-Sleep -Seconds 5 
            }
        }
    } elseif ($Method -eq "GET") {
        $RestResponse = Invoke-RestMethod -Uri $URI -Headers $Header -Method GET -ContentType "$ContentType"
    }
    #>
    if ($Body) {
        $RestResponse = Invoke-RestMethod -Uri $URI -Body $Body -Headers $Header -Method $Method -ContentType "$ContentType"
        $GetURI = $RestResponse.links[0].href
    } elseif ($Method -eq "GET") {
        $RestResponse = Invoke-RestMethod -Uri $URI -Headers $Header -Method GET -ContentType "$ContentType"
        $GetURI = $RestResponse.links[0].href
    }

    # Wait until process is finished
    while ($RestResponse.Status -eq -1) {
        $RestResponse = Invoke-RestMethod -Uri $GetURI -Headers $EPMAPI_HEADER -Method GET -ContentType "$ContentType"
        if ($RestResponse.Status -eq -1) {
            Write-Verbose -Message "Sleeping 5 seconds..."
            Start-Sleep -Seconds 5 
        }
    }
    #Error Handling
    if (-not $IgnoreErrorHandling) {
        if ( $RestResponse.Status -gt 0 ) {
            Write-Verbose -Message "$TaskName ERROR : $($RestResponse.Status)"
            if ($StopOnError) {
                EPM_Log-Error -TaskName "$TaskName" -TaskCommand $TaskCommand -TaskStatus $RestResponse.Status -StopOnError -TaskLevel $TaskLevel -StartTime $TaskStartTime
            } else {
                EPM_Log-Error -TaskName "$TaskName" -TaskCommand $TaskCommand -TaskStatus $RestResponse.Status -TaskLevel $TaskLevel -StartTime $TaskStartTime
            }
        } else {
            Write-Verbose -Message "$TaskName COMPLETE : $($RestResponse.Status)"
            if (-not ($NoLog)) {"$TaskName" | EPM_Log-Item -LogType TASK -LogStatus FINISH -TaskLevel $TaskLevel -StartTime $TaskStartTime}
        }
    } else {
        if (-not ($NoLog)) {"$TaskName" | EPM_Log-Item -LogType TASK -LogStatus FINISH -TaskLevel $TaskLevel -StartTime $TaskStartTime}
    }

    if (-not ($NoLog)) {"REST Response : $RestResponse" | EPM_Log-Item}
    return $RestResponse

}



function EPMAPI_Execute-LoadRule{
<#
    .SYNOPSIS
    Will execute a Data Mgmt load rule via REST API, and download the log after completion or parse

    .EXAMPLE
    EPM_Send-Notification
#>
    param(
        #[MANDATORY] Name of the Load Rule
        [parameter(Mandatory=$true)][String]$LoadRule,
        #[MANDATORY] Starting Period (DM Format, i.e. Oct-20 = October FY20)
        [parameter(Mandatory=$true)][String]$StartPeriod,
        #[MANDATORY] Ending Period (DM Format, i.e. Oct-20 = October FY20)
        [parameter(Mandatory=$true)][String]$EndPeriod,
        #[MANDATORY] Path to the Load file (including file)
        [parameter(Mandatory=$true)][String]$Path,
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
    $FileName = "$($Path.Substring($Path.LastIndexOf('\')+1))"
    $TaskName = "Loading $FileName via $LoadRule for $StartPeriod to $EndPeriod"
    "$TaskName" | EPM_Log-Item -LogType TASK -LogStatus START -TaskLevel $TaskLevel

    EPM_Upload-File -Path "$Path" -DataManagement -TaskLevel ($TaskLevel+1)

    $PayLoad = @{}
    $PayLoad.Add("jobType","DATARULE")
    $PayLoad.Add("jobName","$LoadRule")
    $PayLoad.Add("startPeriod","$StartPeriod")
    $PayLoad.Add("endPeriod","$EndPeriod")
    $PayLoad.Add("importMode","$ImportMode")
    $PayLoad.Add("exportMode","$ExportMode")
    $PayLoad.Add("fileName","$FileName")

    $PayLoad = $PayLoad | ConvertTo-Json
    Write-Verbose -Message "Payload : $Payload"

    $URI = "$EPMAPI_DMG_BASE_URI/jobs"
    Write-Verbose -Message "URI : $URI"


    $RestResponse = EPMAPI_Execute-Request -TaskName "Run $LoadRule" -Method POST -URI $URI -Body $PayLoad -TaskLevel ($TaskLevel+1) -IgnoreErrorHandling
    
    #Error Handling
    


    if ($RestResponse.Status -gt 0) {
        # If the "details" of the rest response are blank then it at least processed, if they are not blank then it could not be processed
        If ($RestResponse.details -eq "") {

            $LogFileName = "$($RestResponse.logFileName.Substring($RestResponse.logFileName.LastIndexOf('/')+1))"
            EPM_Get-File -Name $LogFileName -Path $EPM_PATH_LOGS -TaskLevel ($TaskLevel+1)
            $LogError = EPMAPI_Process-DataMgmtLog -Path "$EPM_PATH_LOGS\$LogFileName" -TaskLevel ($TaskLevel+1)

            if ($LogError -eq 0) {
                #No Reportable Errors / Kickouts
                "$TaskName" | EPM_Log-Item -LogType TASK -LogStatus FINISH -TaskLevel $TaskLevel -StartTime $TaskStartTime
            } elseif ( ($LogError -gt 0) -and ($LogError -lt 99) ) {
                # Kickouts
                EPM_Log-Error -ErrorType WARNING -TaskName $TaskName -TaskCommand ("POST $($URI.Replace("$EPM_URL"," ")) $Payload") -TaskStatus $LogError -TaskLevel $TaskLevel -StartTime $TaskStartTime
            } else {
                # Errors
                if ($StopOnError) {
                    EPM_Log-Error -ErrorType ERROR -TaskName $TaskName -TaskCommand ("POST $($URI.Replace("$EPM_URL"," ")) $Payload") -TaskStatus $LogError -TaskLevel $TaskLevel -StartTime $TaskStartTime -StopOnError
                } else {
                    EPM_Log-Error -ErrorType ERROR -TaskName $TaskName -TaskCommand ("POST $($URI.Replace("$EPM_URL"," ")) $Payload") -TaskStatus $LogError -TaskLevel $TaskLevel -StartTime $TaskStartTime
                }
            }

        } else {
            if ($StopOnError) {
                EPM_Log-Error -ErrorType ERROR -TaskName $TaskName -TaskCommand $RestResponse.details -TaskStatus $RestResponse.Status -TaskLevel $TaskLevel -StartTime $TaskStartTime -StopOnError
            } else {
                EPM_Log-Error -ErrorType ERROR -TaskName $TaskName -TaskCommand $RestResponse.details -TaskStatus $RestResponse.Status -TaskLevel $TaskLevel -StartTime $TaskStartTime
            }
        }
        
    } else {
        "$TaskName" | EPM_Log-Item -LogType TASK -LogStatus FINISH -TaskLevel $TaskLevel -StartTime $TaskStartTime
    }
    
}

function EPMAPI_Process-DataMgmtLog{
<#
    .SYNOPSIS
    Uses EPMAutomate to parse the log provided
    Parses the Log for Import & Export Errors and writes them to a LOG_KICKOUTS as well as the LOG_FULL
    Will return a integer based on the outcome of parsing
       0 = No Errors / Kickouts
       1 = Import Errors
       2 = Export Errors
       3 = Import & Export Errors
       99 = General / Fatal Error

    .EXAMPLE
    EPM_Process-DMGLog -LogPath "$EPM_PATH_LOGS\FINPLAN_1000.log"
    Will parse the FINPLAN_1000.log from Data Management
#>
    Param(
        #[MANDATORY] Name of the Log to Parse
        [parameter(Mandatory=$true)][String]$Path,
        #Level of Task Being Executed for Logging Purposes
        [Int]$TaskLevel = 0,
        #Turns off Logging
        [Switch]$NoLog
    )

    $CurFnc = EPM_Get-Function
    
    $TaskStartTime = Get-Date
    $Name = "$($Path.Substring($Path.LastIndexOf('\')+1))"
    $TaskName = "Process Data Mgmt Log : $Name"
    "$TaskName" | EPM_Log-Item -LogType TASK -LogStatus START -TaskLevel $TaskLevel

    if (Test-Path $Path) {
        #Parse Log for Errors
        $ErrImport = $false
        $FirstRow = $true
        foreach($line in Get-Content ($Path)) {
            foreach($ErrorMsg in @("[BLANK]","[NN]","[TC]","[NULL ACCOUNT VALUE]","[ERROR_INVALID_PERIOD]")){
                if($line.Contains($ErrorMsg)){
                    if ($FirstRow) { 
                        "" | EPM_Log-Item
                        "[ IMPORT ERRORS ]" | EPM_Log-Item
                        $FirstRow = $false
                    }
                    $ErrImport = $true
                    $line.Substring($line.IndexOf($ErrorMsg)) | EPM_Log-Item
                    #Write to Kickout File
                    "IMPORT : $($line.Substring($line.IndexOf($ErrorMsg)))" | Add-Content -Path $EPM_LOG_KICKOUTS
                }
            }
        
        }

        $ErrExport = $false
        $FirstRow = $true
        foreach($line in Get-Content ($Path)) {
            foreach($ErrorMsg in @("Error: 3303","The member")){
                if($line.Contains($ErrorMsg)){
                    if ($FirstRow) {
                        "" | EPM_Log-Item
                        "[ EXPORT ERRORS ]" | EPM_Log-Item 
                        $FirstRow = $false
                    }
                    $ErrExport = $true
                    $line.Substring($line.IndexOf($ErrorMsg)) | EPM_Log-Item
                    #Write to Kickout File
                    "EXPORT : $($line.Substring($line.IndexOf($ErrorMsg)))" | Add-Content -Path $EPM_LOG_KICKOUTS
                }
            }
        }
        
        $ErrGeneral = $false
        $FirstRow = $true
        foreach($line in Get-Content ($Path)) {
            foreach($ErrorMsg in @(" ERROR "," FATAL ", " WARN ")){
                if($line.Contains($ErrorMsg)){
                    if ($FirstRow) { 
                        "" | EPM_Log-Item
                        "[ GENERAL ERRORS ]" | EPM_Log-Item
                        $FirstRow = $false
                    }
                    $ErrGeneral = $true                    
                    $line.Substring($line.IndexOf($ErrorMsg)) | EPM_Log-Item
                }
            }
        }

    } else {
        Write-Verbose -Message "$CurFnc | Error in EPM_Get-File for $Name"
    }

    #Return Error Code
    #   0 = No Errors / Kickouts
    #   1 = Import Errors
    #   2 = Export Errors
    #   3 = Import & Export Errors
    #   99 = General / Fatal Error

    if ( $ErrImport -or $ErrExport -or $ErrGeneral) {
        #Error was found

        if ($ErrImport) {
            if ($ErrExport) {
                #Import & Export Errors
                $RetVal = 3
            } else {
                #Import Errors Only
                $RetVal = 1
            }
        } elseif ($ErrExport) {
            #Export Errors Only
            $RetVal = 2
        } else {
            #General / Fatal Error
            $RetVal = 99
        }

    } else {
        $RetVal = 0
    }


    "$TaskName" | EPM_Log-Item -LogType TASK -LogStatus FINISH -TaskLevel $TaskLevel -StartTime $TaskStartTime

    return $RetVal

}
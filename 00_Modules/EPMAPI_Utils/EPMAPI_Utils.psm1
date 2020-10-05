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